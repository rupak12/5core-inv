<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use App\Models\FbaReportsMaster;
use App\Models\FbaTable;
use App\Models\FbaFees;


use Carbon\Carbon;

class FetchFbaReportsMaster extends Command
{
    protected $signature = 'app:fetch-fba-reports';
    protected $description = 'Fetch FBA sales, views and fees and insert into fba_reports_master';

    public function handle()
    {
        $this->info("🔍 Fetching Amazon FBA reports...");

        // 1️⃣ Get access token
        $accessToken = $this->getAccessToken();
        if (!$accessToken) {
            $this->error("❌ Cannot proceed without access token");
            return 1;
        }
        $this->info("✅ Access token obtained");

                // 2️⃣ Fetch valid ASINs from FBA table (for sales reports)
        $validAsins = FbaTable::whereNotNull('asin')->where('asin', '!=', '')->pluck('asin')->toArray();
        
        // 2️⃣ Also get SKU to ASIN mapping for fees matching
        $skuToAsinMap = FbaTable::whereNotNull('seller_sku')
            ->whereNotNull('asin')
            ->where('seller_sku', '!=', '')
            ->where('asin', '!=', '')
            ->pluck('asin', 'seller_sku')
            ->toArray();
            
        // Create reverse mapping: ASIN to SKU for sales reports
        $asinToSkuMap = FbaTable::whereNotNull('seller_sku')
            ->whereNotNull('asin')
            ->where('seller_sku', '!=', '')
            ->where('asin', '!=', '')
            ->pluck('seller_sku', 'asin')
            ->toArray();
        
        if (empty($validAsins)) {
            $this->warn("⚠️ No ASINs found in FBA table");
            return 0;
        }
        
        $this->info("📋 Found " . count($validAsins) . " ASINs, " . count($skuToAsinMap) . " SKU mappings, and " . count($asinToSkuMap) . " ASIN mappings in FBA table");

        $marketplaceIds = env('SPAPI_MARKETPLACE_ID', 'ATVPDKIKX0DER');
        $endpointBase = 'https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports';

        // 3️⃣ Fetch latest Sales & Traffic report
        $reportIdData = Http::withHeaders([
            'x-amz-access-token' => $accessToken,
            'Content-Type' => 'application/json',
        ])->get($endpointBase, [
            'reportTypes' => 'GET_SALES_AND_TRAFFIC_REPORT',
            'marketplaceIds' => $marketplaceIds,
            'pageSize' => 1
        ])->json();

        if (empty($reportIdData['reports'])) {
            $this->warn("⚠️ No Sales & Traffic reports found");
            return 0;
        }

        $report = $reportIdData['reports'][0];
        $documentId = $report['reportDocumentId'] ?? null;
        if (!$documentId) {
            $this->error("❌ No document ID found");
            return 1;
        }

        // 4️⃣ Download report document
        $csvData = $this->downloadReportDocument($accessToken, $documentId);
        if (!$csvData) return 1;

        $inserted = 0;
        $jsonData = json_decode($csvData, true);

        if (json_last_error() === JSON_ERROR_NONE && isset($jsonData['salesAndTrafficByAsin'])) {
            $this->info("📊 Processing JSON Sales & Traffic report...");
            $salesTrafficData = $jsonData['salesAndTrafficByAsin'];

            foreach ($salesTrafficData as $asinData) {
                $parentAsin = $asinData['parentAsin'] ?? null;
                $childAsin = $asinData['childAsin'] ?? null;
                
                // Check if either parent or child ASIN matches our FBA inventory
                $matchingAsin = null;
                if ($parentAsin && in_array($parentAsin, $validAsins)) {
                    $matchingAsin = $parentAsin;
                } elseif ($childAsin && in_array($childAsin, $validAsins)) {
                    $matchingAsin = $childAsin;
                }
                
                if (!$matchingAsin) continue;
                
                // Get the actual SKU from ASIN mapping
                $sku = $asinToSkuMap[$matchingAsin] ?? $matchingAsin;
                
                if (!$sku) continue;

                $traffic = $asinData['trafficByAsin'] ?? [];
                $pageViews = (int)($traffic['pageViews'] ?? 0);

                $currentMonth = strtolower(date('M'));
                $monthlyViews = [];
                foreach (range(1, 12) as $m) {
                    $monthName = strtolower(Carbon::create()->month($m)->format('M'));
                    $monthlyViews[$monthName.'_views'] = $currentMonth === $monthName ? $pageViews : 0;
                }

                FbaReportsMaster::updateOrCreate(
                    ['seller_sku' => $sku, 'year' => date('Y')],
                    array_merge([
                        'asin' => $childAsin ?: $parentAsin,
                        'current_month_views' => $pageViews,
                        'total_views' => $pageViews,
                        'fulfillment_fee' => 0,
                        'referral_fee' => 0,
                        'storage_fee' => 0,
                        'total_fee' => 0,
                    ], $monthlyViews)
                );

                $inserted++;
            }
        } else {
            // CSV fallback
            $lines = explode("\n", trim($csvData));
            if (count($lines) <= 1) {
                $this->warn("⚠️ No data lines found in report");
                return 0;
            }

            $headers = str_getcsv(array_shift($lines), "\t");
            foreach ($lines as $line) {
                if (!trim($line)) continue;
                $row = array_combine($headers, str_getcsv($line, "\t"));

                $sku = $row['sku'] ?? $row['seller-sku'] ?? null;
                $asin = $row['asin'] ?? null;
                
                // For CSV fallback, we need to check if ASIN matches our inventory, then use proper SKU
                if (!$asin || !in_array($asin, $validAsins)) continue;
                
                // Get proper SKU from ASIN mapping, fallback to found SKU
                $properSku = $asinToSkuMap[$asin] ?? $sku;
                if (!$properSku) continue;

                $monthlyViews = [];
                foreach (range(1,12) as $m) {
                    $monthName = strtolower(Carbon::create()->month($m)->format('M'));
                    $viewKey = $monthName.'_views';
                    $monthlyViews[$viewKey] = (int)($row[$monthName] ?? 0);
                }

                $totalViews = array_sum($monthlyViews);
                $currentViews = (int)($row['pageViews'] ?? 0);

                FbaReportsMaster::updateOrCreate(
                    ['seller_sku' => $properSku, 'year' => date('Y')],
                    array_merge([
                        'asin' => $asin,
                        'current_month_views' => $currentViews,
                        'total_views' => $totalViews,
                        'fulfillment_fee' => 0,
                        'referral_fee' => 0,
                        'storage_fee' => 0,
                        'total_fee' => 0,
                    ], $monthlyViews)
                );

                $inserted++;
            }
        }

        $this->info("✅ Inserted/updated {$inserted} records for valid SKUs");

        // 5️⃣ Fetch FBA Estimated Fees
        $reportIdFees = Http::withHeaders([
            'x-amz-access-token' => $accessToken,
            'Content-Type' => 'application/json',
        ])->get($endpointBase, [
            'reportTypes' => 'GET_FBA_ESTIMATED_FBA_FEES_TXT_DATA',
            'marketplaceIds' => $marketplaceIds,
            'pageSize' => 1
        ])->json();

        if (!empty($reportIdFees['reports'])) {
            $feesReport = $reportIdFees['reports'][0];
            $feesDocId = $feesReport['reportDocumentId'] ?? null;

            if ($feesDocId) {
                // Truncate FbaFees table to avoid duplicates
                $this->info("🗑️ Clearing existing FBA fees data...");
                FbaFees::truncate();
                
                $csvFees = $this->downloadReportDocument($accessToken, $feesDocId);
                $linesFees = explode("\n", trim($csvFees));
                $headersFees = str_getcsv(array_shift($linesFees), "\t");
                
                $this->info("📊 Fees report has " . count($linesFees) . " data lines");
                
                $feesInserted = 0;
                $feesMatched = 0;
                $skusChecked = [];
                $reportGeneratedAt = now();

                foreach ($linesFees as $line) {
                    if (!trim($line)) continue;
                    $row = array_combine($headersFees, str_getcsv($line, "\t"));
                    $sku = $row['sku'] ?? null;
                    
                    if ($sku) {
                        $skusChecked[] = $sku;
                        
                        // Get ASIN from our mapping
                        $asin = $skuToAsinMap[$sku] ?? null;
                        
                        // Insert all fee data into FbaFees table
                        FbaFees::create([
                            'seller_sku' => $sku,
                            'report_generated_at' => $reportGeneratedAt,
                            'fnsku' => $this->cleanString($row['fnsku'] ?? null),
                            'asin' => $asin ?: ($this->cleanString($row['asin'] ?? null)),
                            'amazon_store' => $this->cleanString($row['amazon-store'] ?? null),
                            'product_name' => $this->cleanString($row['product-name'] ?? null),
                            'product_group' => $this->cleanString($row['product-group'] ?? null),
                            'brand' => $this->cleanString($row['brand'] ?? null),
                            'fulfilled_by' => $this->cleanString($row['fulfilled-by'] ?? null),
                            'your_price' => $this->cleanDecimal($row['your-price'] ?? null),
                            'sales_price' => $this->cleanDecimal($row['sales-price'] ?? null),
                            'estimated_fee_total' => $this->cleanDecimal($row['estimated-fee-total'] ?? null),
                            'estimated_referral_fee_per_unit' => $this->cleanDecimal($row['estimated-referral-fee-per-unit'] ?? null),
                            'expected_fulfillment_fee_per_unit' => $this->cleanDecimal($row['expected-fulfillment-fee-per-unit'] ?? null),
                        ]);
                        
                        $feesInserted++;
                        
                        // If we have ASIN mapping, also update FbaReportsMaster
                        if ($asin) {
                            $fulfillmentFee = (float)($row['expected-fulfillment-fee-per-unit'] ?? 0);
                            $referralFee = (float)($row['estimated-referral-fee-per-unit'] ?? 0);
                            $totalFee = (float)($row['estimated-fee-total'] ?? 0);
                            
                            FbaReportsMaster::where('seller_sku', $sku)
                                ->orWhere('asin', $asin)
                                ->update([
                                    'fulfillment_fee' => $fulfillmentFee,
                                    'referral_fee' => $referralFee,
                                    'storage_fee' => 0,
                                    'total_fee' => $totalFee
                                ]);
                            
                            $feesMatched++;
                        }
                    }
                }

                $this->info("✅ FBA fees: {$feesInserted} records inserted, {$feesMatched} matched with inventory");
                $this->info("🔍 Sample SKUs from fees report: " . implode(', ', array_slice($skusChecked, 0, 5)));
                $this->info("🔍 Total unique SKU mappings available: " . count($skuToAsinMap));
            }
        }

        $this->info("🏁 FBA reports fetch complete!");
        return 0;
    }

    private function getAccessToken()
    {
        $res = Http::asForm()->post('https://api.amazon.com/auth/o2/token', [
            'grant_type' => 'refresh_token',
            'refresh_token' => env('SPAPI_REFRESH_TOKEN'),
            'client_id' => env('SPAPI_CLIENT_ID'),
            'client_secret' => env('SPAPI_CLIENT_SECRET'),
        ]);

        if ($res->failed()) {
            $this->error('❌ Failed to get access token');
            return null;
        }
        return $res->json()['access_token'] ?? null;
    }

    private function downloadReportDocument($token, $docId)
    {
        $endpoint = 'https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/documents/'.$docId;
        $info = Http::withHeaders(['x-amz-access-token'=>$token])->get($endpoint)->json();
        $url = $info['url'] ?? null;
        if (!$url) return null;

        $data = Http::timeout(120)->get($url)->body();
        if (($info['compressionAlgorithm'] ?? null) === 'GZIP') $data = gzdecode($data);
        return $data;
    }

    private function cleanString($value)
    {
        if ($value === null) return null;
        
        // Convert to UTF-8 and remove problematic characters
        $cleaned = mb_convert_encoding($value, 'UTF-8', 'UTF-8');
        
        // Remove or replace specific problematic characters
        $cleaned = str_replace([chr(150), chr(151), chr(147), chr(148)], ['-', '-', '"', '"'], $cleaned);
        
        // Remove any remaining non-printable characters except basic ASCII
        $cleaned = preg_replace('/[^\x20-\x7E]/', '', $cleaned);
        
        return trim($cleaned);
    }

    private function cleanDecimal($value)
    {
        if ($value === null || $value === '' || $value === '--') return null;
        
        // Remove any non-numeric characters except decimal point
        $cleaned = preg_replace('/[^0-9.-]/', '', $value);
        
        return is_numeric($cleaned) ? (float)$cleaned : null;
    }
}
