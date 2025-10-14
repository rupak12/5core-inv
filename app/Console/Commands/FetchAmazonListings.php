<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use App\Models\AmazonDatasheet;
use Illuminate\Support\Facades\Log;
use Carbon\Carbon;

class FetchAmazonListings extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:fetch-amazon-listings';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Command description';

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $accessToken = $this->getAccessToken();
        info('Access Token', [$accessToken]);

        $marketplaceId = env('SPAPI_MARKETPLACE_ID');

        // Step 1: Request the report
        $response = Http::withHeaders([
            'x-amz-access-token' => $accessToken,
        ])->post('https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports', [
            'reportType' => 'GET_MERCHANT_LISTINGS_ALL_DATA',
            'marketplaceIds' => [$marketplaceId],
        ]);

        $this->error('Report Request Response: ' . $response->body());
        $reportId = $response['reportId'] ?? null;
        if (!$reportId) {
            $this->error('Failed to request report.');
            return;
        }

        // Step 2: Wait for report generation
        do {
            sleep(15);
            $status = Http::withHeaders([
                'x-amz-access-token' => $accessToken,
            ])->get("https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports/{$reportId}");
            $processingStatus = $status['processingStatus'] ?? 'UNKNOWN';
            $this->info("Waiting... Status: $processingStatus");
        } while ($processingStatus !== 'DONE');

        $documentId = $status['reportDocumentId'];
        $doc = Http::withHeaders([
            'x-amz-access-token' => $accessToken,
        ])->get("https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/documents/{$documentId}");

        $url = $doc['url'] ?? null;
        if (!$url) {
            $this->error('Document URL not found.');
            return;
        }

        // Step 3: Download and parse the data
        $csv = file_get_contents($url);
        $lines = explode("\n", $csv);
        $headers = explode("\t", array_shift($lines));

        foreach ($lines as $line) {
            $row = str_getcsv($line, "\t");
            if (count($row) < count($headers)) continue;

            $data = array_combine($headers, $row);

            // Fulfillment channel filter
            if (($data['fulfillment-channel'] ?? '') !== 'DEFAULT') continue;

            $asin = $data['asin1'] ?? null;
            $sku = isset($data['seller-sku']) ? preg_replace('/[^\x20-\x7E]/', '', trim($data['seller-sku'])) : null;
            $price = isset($data['price']) && is_numeric($data['price']) ? $data['price'] : null;

            if ($asin) {                
                AmazonDatasheet::updateOrCreate(
                    ['asin' => $asin],                    [
                        'sku' => $sku,
                        'price' => $price,
                    ]
                );
            }
        }

        $this->info('ASIN and price and sku data imported successfully.');

        $this->getUnitOrderedAndSessionsData();
    }

    private function getAccessToken()
    {
        $res = Http::asForm()->post('https://api.amazon.com/auth/o2/token', [
            'grant_type' => 'refresh_token',
            'refresh_token' => env('SPAPI_REFRESH_TOKEN'),
            'client_id' => env('SPAPI_CLIENT_ID'),
            'client_secret' => env('SPAPI_CLIENT_SECRET'),
        ]);

        return $res['access_token'] ?? null;
    }

    private function getUnitOrderedAndSessionsData(){
        $accessToken = $this->getAccessToken();
        $marketplaceId = env('SPAPI_MARKETPLACE_ID');

        $l30End = Carbon::today()->copy()->subDay()->endOfDay();           
        $l30Start = $l30End->copy()->subDays(30)->startOfDay();   

        $l60End = $l30Start->copy()->subDay(2)->endOfDay();
        $l60Start = $l60End->copy()->subDays(30)->startOfDay();

        $dateRanges = [
            'l30' => [$l30Start->toIso8601ZuluString(), $l30End->toIso8601ZuluString()],
            'l60' => [$l60Start->toIso8601ZuluString(), $l60End->toIso8601ZuluString()],
        ];

        info('$dateRanges', [$dateRanges]);

        foreach ($dateRanges as $key => [$start, $end]) {
            // 1. Create report
            $response = Http::withHeaders([
                'x-amz-access-token' => $accessToken,
            ])->post('https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports', [
                'reportType' => 'GET_SALES_AND_TRAFFIC_REPORT',
                'marketplaceIds' => [$marketplaceId],
                'dataStartTime' => $start,
                'dataEndTime' => $end,
                "reportOptions" => ['asinGranularity' => 'CHILD'],
            ]);
    
            $reportId = $response['reportId'] ?? null;
            if (!$reportId) {
                $this->error("Failed to create report for $key");
                continue;
            }
    
            // 2. Wait for report to be ready
            do {
                sleep(15);
                $status = Http::withHeaders([
                    'x-amz-access-token' => $accessToken,
                ])->get("https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports/{$reportId}");
                $processingStatus = $status['processingStatus'] ?? 'UNKNOWN';
            } while ($processingStatus !== 'DONE');
    
            // 3. Download document
            $documentId = $status['reportDocumentId'] ?? null;
            if (!$documentId) {
                $this->error("No documentId for report $reportId");
                continue;
            }
            $doc = Http::withHeaders([
                'x-amz-access-token' => $accessToken,
            ])->get("https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/documents/{$documentId}");
    
            $url = $doc['url'] ?? null;
            if (!$url) {
                Log::error("No document URL for $key");
                continue;
            }
        
            $rows = $this->getGZConvertedData($url);
            if (!isset($rows[0])) {
                Log::warning("Empty rows returned for $key");
                continue;
            }
        
            $decoded = json_decode($rows[0], true);
            info('decoded', [$decoded]);
            if (json_last_error() !== JSON_ERROR_NONE) {
                Log::error('JSON decode error: ' . json_last_error_msg());
                continue;
            }
        
            $salesAndTraffic = $decoded['salesAndTrafficByAsin'] ?? [];

            info("Decoded structure for $key", $decoded);
            if (empty($salesAndTraffic)) {
                Log::warning("No salesAndTrafficByAsin data for $key");
                continue;
            }
        
            foreach ($salesAndTraffic as $asinData) {
                $asin = $asinData['childAsin'] ?? null;
                $unitsOrdered = $asinData['salesByAsin']['unitsOrdered'] ?? 0;
                $sessions = $asinData['trafficByAsin']['sessions'] ?? 0;
        
                if (!$asin) {
                    Log::warning("ASIN missing in entry");
                    continue;
                }
    
                $updated = AmazonDatasheet::where('asin', $asin)->update([
                    "units_ordered_{$key}" => $unitsOrdered,
                    "sessions_{$key}"      => $sessions,
                ]);
        
                if ($updated) {
                    Log::info("Updated ASIN: $asin | Units Ordered: $unitsOrdered | Sessions: $sessions");
                } else {
                    Log::info("ASIN not found in database: $asin");
                }
            }
        }
    
        $this->info('Sales and traffic data updated for L30 and L60.');
    }

    private function getGZConvertedData($url)
    {
        try {
            $response = Http::timeout(120)->get($url);
            if (!$response->ok()) {
                Log::error("Failed to download GZ file from: " . $url);
                return null;
            }
    
            $gzPath = storage_path('app/temp_' . uniqid() . '.gz');
            $extractedPath = storage_path('app/extracted_' . uniqid() . '.tsv');
    
            file_put_contents($gzPath, $response->body());
    
            // Extract
            $gz = gzopen($gzPath, 'rb');
            $out = fopen($extractedPath, 'wb');
            while (!gzeof($gz)) {
                fwrite($out, gzread($gz, 4096));
            }
            fclose($out);
            gzclose($gz);
    
            $content = file_get_contents($extractedPath);
    
            // Cleanup temp files
            @unlink($gzPath);
            @unlink($extractedPath);
    
            return explode("\n", trim($content)); // Return rows including header
        } catch (\Exception $e) {
            Log::error('Error extracting GZ: ' . $e->getMessage());
            return null;
        }
    }
}