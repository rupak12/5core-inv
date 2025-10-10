<?php

namespace App\Services;

use App\Models\EbayDataMetricData;
use Carbon\Carbon;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use ZipArchive;

class FetchEbayDataMetricServiceClass
{
    public function fetchAndInsertEbayMetrics(): bool
    {
        try {
            $token = $this->generateEbayToken();
            if (!$token) {
                Log::error('Failed to generate Ebay token.');
                return false;
            }

            $rateLimit = $this->getRateLimit($token);
            if (!$rateLimit) {
                Log::error('Failed to get rate limit data.');
                return false;
            }

            $dateRanges = $this->getDateRanges();

            // Fetch Active Inventory
            $listingData = $this->fetchAndParseReport('LMS_ACTIVE_INVENTORY_REPORT', null, $token);

            // Check if we got any data
            if (empty($listingData)) {
                Log::error('❌ No listing data received from eBay API. This could be due to rate limiting or empty inventory.');
                return false;
            }

            Log::info('📦 Successfully fetched ' . count($listingData) . ' listing records.');

            // Map item_id => sku
            $itemIdToSku = [];
            foreach ($listingData as $row) {
                if (!empty($row['item_id']) && !empty($row['sku'])) {
                    $itemIdToSku[$row['item_id']] = $row['sku'];
                }
            }

            // Fetch views from Analytics API
            $viewsByItemId = $this->fetchViews($itemIdToSku, $token);

            // Insert or update EbayDataMetricData table
            $insertedCount = 0;
            foreach ($listingData as $row) {
                $itemId = $row['item_id'] ?? null;
                if (!$itemId) continue;

                EbayDataMetricData::updateOrCreate(
                    ['item_id' => $itemId, 'sku' => $row['sku'] ?? ''],
                    [
                        'ebay_data_price' => $row['price'] ?? null,
                        'ebay_data_l30' => 0,
                        'ebay_data_l60' => 0,
                        'ebay_data_views' => 0,
                    ]
                );
                $insertedCount++;
            }

            Log::info('✅ Inserted/Updated ' . $insertedCount . ' records in EbayDataMetricData table.');

            // Update Views
            foreach ($viewsByItemId as $itemId => $views) {
                EbayDataMetricData::where('item_id', $itemId)->update(['ebay_data_views' => $views]);
            }

            Log::info('✅ Views updated for ' . count($viewsByItemId) . ' items.');

            // Get last 30 and 60 day quantities
            $existingSkus = EbayDataMetricData::pluck('sku')->filter()->toArray();
            $l30Qty = $this->getQuantityBySkuFromOrders($token, $dateRanges['l30']['start'], $dateRanges['l30']['end'], $existingSkus);
            $l60Qty = $this->getQuantityBySkuFromOrders($token, $dateRanges['l60']['start'], $dateRanges['l60']['end'], $existingSkus);

            foreach ($existingSkus as $sku) {
                EbayDataMetricData::where('sku', $sku)->update([
                    'ebay_data_l30' => $l30Qty[$sku] ?? 0,
                    'ebay_data_l60' => $l60Qty[$sku] ?? 0,
                ]);
            }

            Log::info('✅ eBay metrics fetched and stored successfully.');
            return true;

        } catch (\Throwable $e) {
            Log::error('❌ eBay metrics fetch failed: ' . $e->getMessage(), ['trace' => $e->getTraceAsString()]);
            return false;
        }
    }

    // ---------------------- Helper Methods ----------------------

    private function getDateRanges(): array
    {
        $today = Carbon::today();
        return [
            'l30' => [
                'start' => $today->copy()->subDays(30)->addDay(),
                'end' => $today->copy()->subDay(),
            ],
            'l60' => [
                'start' => $today->copy()->subDays(60)->addDay(),
                'end' => $today->copy()->subDays(31),
            ],
        ];
    }

    private function fetchViews(array $itemIdToSku, string $token): array
    {
        $viewsByItemId = [];

        if (empty($itemIdToSku)) return $viewsByItemId;

        $itemIdChunks = array_chunk(array_keys($itemIdToSku), 20);

        foreach ($itemIdChunks as $chunk) {
            $ids = implode('|', $chunk);
            $dateRange = now()->subDays(30)->format('Ymd') . '..' . now()->format('Ymd');
            $url = "https://api.ebay.com/sell/analytics/v1/traffic_report?dimension=LISTING&filter=listing_ids:%7B{$ids}%7D,date_range:[{$dateRange}]&metric=LISTING_VIEWS_TOTAL&sort=LISTING_VIEWS_TOTAL";

            $response = Http::withToken($token)->get($url);
            if ($response->ok()) {
                foreach ($response->json()['records'] ?? [] as $record) {
                    $itemId = $record['dimensionValues'][0]['value'] ?? null;
                    $views = $record['metricValues'][0]['value'] ?? null;
                    if ($itemId && $views !== null) {
                        $viewsByItemId[$itemId] = $views;
                    }
                }
            }
        }

        return $viewsByItemId;
    }

    private function fetchAndParseReport($reportType, $range, $token): array
    {
        $apiUrl = 'https://api.ebay.com/sell/feed/v1/inventory_task';

        $payload = [
            'feedType' => $reportType,
            'format' => 'TSV_GZIP',
            'schemaVersion' => '1.0',
        ];

        $response = Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
            'Content-Type' => 'application/json',
        ])->post($apiUrl, $payload);

        if (!$response->successful()) {
            Log::error("Task creation failed: " . $response->body());
            return [];
        }

        $location = $response->header('Location');
        if (!$location) return [];

        $taskId = basename($location);

        do {
            sleep(10);
            $statusResponse = Http::withHeaders([
                'Authorization' => 'Bearer ' . $token,
            ])->get("https://api.ebay.com/sell/feed/v1/inventory_task/{$taskId}");

            $status = $statusResponse['status'] ?? 'PENDING';
        } while (!in_array($status, ['COMPLETED', 'COMPLETED_WITH_ERROR', 'FAILED']));

        if ($status === 'FAILED') return [];

        return $this->downloadAndParseEbayReport($taskId, $token);
    }

    private function downloadAndParseEbayReport(string $taskId, string $token): array
    {
        $baseTaskUrl = "https://api.ebay.com/sell/feed/v1/task/{$taskId}/download_result_file";
        $filePath = storage_path("app/inventory_{$taskId}");
        $zipPath = "{$filePath}.zip";

        $response = Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
        ])->get($baseTaskUrl);

        $content = $response->body();

        if (substr($content, 0, 2) === "PK") {
            file_put_contents($zipPath, $content);
            $zip = new ZipArchive;

            if ($zip->open($zipPath) === TRUE) {
                $zip->extractTo(storage_path('app/'));
                $zip->close();

                $xmlFiles = glob(storage_path('app/*.xml'));
                if (empty($xmlFiles)) return [];

                $xmlPath = $xmlFiles[0];
                $xml = simplexml_load_file($xmlPath);
                if (!$xml) return [];

                $data = [];
                foreach ($xml->ActiveInventoryReport->SKUDetails as $item) {
                    $itemId = (string) $item->ItemID ?? null;
                    if (!$itemId) continue;

                    $data[] = [
                        'item_id' => $itemId,
                        'sku' => (string) ($item->SKU ?? ''),
                        'price' => (float) ($item->Price ?? 0),
                    ];

                    // Handle variations
                    foreach ($item->Variations->Variation ?? [] as $variation) {
                        $data[] = [
                            'item_id' => $itemId,
                            'sku' => (string) ($variation->SKU ?? ''),
                            'price' => (float) ($variation->Price ?? 0),
                        ];
                    }
                }

                @unlink($zipPath);
                @unlink($xmlPath);
                return $data;
            }
        }

        return [];
    }

    private function getQuantityBySkuFromOrders($token, Carbon $from, Carbon $to, array $onlyTheseSkus = []): array
    {
        $allQuantities = [];

        $url = "https://api.ebay.com/sell/fulfillment/v1/order?filter=creationdate:[{$from->format('Y-m-d\TH:i:s.000\Z')}..{$to->format('Y-m-d\TH:i:s.000\Z')}]&limit=200";

        do {
            $response = Http::withToken($token)->get($url);

            if (!$response->ok()) {
                Log::error("Fulfillment fetch failed: " . $response->body());
                break;
            }

            $data = $response->json();
            foreach ($data['orders'] ?? [] as $order) {
                foreach ($order['lineItems'] ?? [] as $line) {
                    $sku = $line['sku'] ?? null;
                    $qty = (int) ($line['quantity'] ?? 0);
                    if (!$sku || !in_array($sku, $onlyTheseSkus)) continue;

                    $allQuantities[$sku] = ($allQuantities[$sku] ?? 0) + $qty;
                }
            }

            $url = $data['next'] ?? null;
        } while ($url);

        return $allQuantities;
    }

    private function generateEbayToken(): ?string
    {
        $clientId = env('EBAY_APP_ID');
        $clientSecret = env('EBAY_CERT_ID');

        $scope = implode(' ', [
            'https://api.ebay.com/oauth/api_scope',
            'https://api.ebay.com/oauth/api_scope/sell.account',
            'https://api.ebay.com/oauth/api_scope/sell.inventory',
            'https://api.ebay.com/oauth/api_scope/sell.fulfillment',
            'https://api.ebay.com/oauth/api_scope/sell.analytics.readonly',
        ]);

        try {
            $response = Http::asForm()
                ->withBasicAuth($clientId, $clientSecret)
                ->post('https://api.ebay.com/identity/v1/oauth2/token', [
                    'grant_type' => 'refresh_token',
                    'refresh_token' => env('EBAY_REFRESH_TOKEN'),
                    'scope' => $scope,
                ]);

            if ($response->successful()) {
                return $response->json()['access_token'];
            }

            Log::error('Ebay token refresh error', ['response' => $response->json()]);
        } catch (\Exception $e) {
            Log::error('Ebay token exception: ' . $e->getMessage());
        }

        return null;
    }

    private function getRateLimit($token)
    {
        try {
            $response = Http::withToken($token)->get('https://api.ebay.com/developer/analytics/v1_beta/rate_limit/');
            return $response->successful() ? $response->json() : null;
        } catch (\Exception $e) {
            Log::error('Ebay getRateLimit exception: ' . $e->getMessage());
            return null;
        }
    }
}
