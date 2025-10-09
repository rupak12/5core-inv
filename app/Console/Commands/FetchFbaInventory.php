<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use App\Models\FbaTable;
use App\Models\FbaPrice;

class FetchFbaInventory extends Command
{
    protected $signature = 'app:fetch-fba-inventory 
        {--debug : Show detailed debug info} 
        {--insert : Insert data into fba_table after preview} 
        {--prices : Also fetch prices using GET_MERCHANT_LISTINGS_ALL_DATA} 
        {--prices-only : Only fetch prices without inventory data}';
    protected $description = 'Fetch FBA inventory data and optionally prices';

    public function handle()
    {
        $this->info('🚀 Starting FBA fetch...');

        $accessToken = $this->getAccessToken();
        if (!$accessToken) {
            $this->error('❌ Failed to get access token');
            return 1;
        }

        $marketplaceId = env('SPAPI_MARKETPLACE_ID', 'ATVPDKIKX0DER');
        $endpoint = env('SPAPI_ENDPOINT', 'https://sellingpartnerapi-na.amazon.com');

        // If prices-only, skip inventory
        if ($this->option('prices-only')) {
            $this->fetchPrices($accessToken, $endpoint, $marketplaceId);
            return 0;
        }

        // Step 1: Create inventory report
        $reportId = $this->createReport($accessToken, $endpoint, $marketplaceId, 'GET_FBA_MYI_ALL_INVENTORY_DATA');
        if (!$reportId) return 1;

        // Step 2: Wait for report completion
        $statusData = $this->waitForReport($accessToken, $endpoint, $reportId);
        if (!$statusData) return 1;

        $documentId = $statusData['reportDocumentId'] ?? null;
        if (!$documentId) {
            $this->error('❌ No document ID found');
            return 1;
        }

        $csvData = $this->downloadReportDocument($accessToken, $endpoint, $documentId);
        if (!$csvData) return 1;

        $this->saveReportData($csvData);
        $this->showDataPreview($csvData);

        if ($this->option('insert')) {
            $this->insertInventory($csvData);
        }

        if ($this->option('prices')) {
            $this->fetchPrices($accessToken, $endpoint, $marketplaceId);
        }

        return 0;
    }

    private function getAccessToken()
    {
        try {
            $res = Http::asForm()->post('https://api.amazon.com/auth/o2/token', [
                'grant_type' => 'refresh_token',
                'refresh_token' => env('SPAPI_REFRESH_TOKEN'),
                'client_id' => env('SPAPI_CLIENT_ID'),
                'client_secret' => env('SPAPI_CLIENT_SECRET'),
            ]);

            if ($res->failed()) {
                Log::error('Access token request failed', $res->json());
                return null;
            }

            return $res->json()['access_token'] ?? null;
        } catch (\Throwable $e) {
            Log::error('Access token error: ' . $e->getMessage());
            return null;
        }
    }

    private function createReport($token, $endpoint, $marketplaceId, $reportType)
    {
        $res = Http::withHeaders([
            'x-amz-access-token' => $token,
            'Content-Type' => 'application/json',
        ])->post("{$endpoint}/reports/2021-06-30/reports", [
            'reportType' => $reportType,
            'marketplaceIds' => [$marketplaceId]
        ]);

        if ($res->failed()) {
            $this->error('❌ Failed to create report: ' . $res->body());
            return null;
        }

        return $res->json()['reportId'] ?? null;
    }

    private function waitForReport($token, $endpoint, $reportId)
    {
        $maxWait = 300; $wait = 0;

        while ($wait < $maxWait) {
            sleep(15); $wait += 15;

            $res = Http::withHeaders(['x-amz-access-token' => $token])
                ->get("{$endpoint}/reports/2021-06-30/reports/{$reportId}");

            if ($res->failed()) {
                $this->error('❌ Report status check failed');
                return null;
            }

            $data = $res->json();
            $status = $data['processingStatus'] ?? 'UNKNOWN';
            $this->line("Status: {$status} ({$wait}s)");

            if ($status === 'DONE') return $data;
            if (in_array($status, ['FATAL','CANCELLED'])) {
                $this->error("❌ Report failed with status: {$status}");
                if ($this->option('debug')) $this->line(json_encode($data, JSON_PRETTY_PRINT));
                return null;
            }
        }

        $this->error("❌ Report timeout after {$maxWait}s");
        return null;
    }

    private function downloadReportDocument($token, $endpoint, $docId)
    {
        $res = Http::withHeaders(['x-amz-access-token' => $token])
            ->get("{$endpoint}/reports/2021-06-30/documents/{$docId}");

        if ($res->failed()) {
            $this->error('❌ Failed to download document');
            return null;
        }

        $info = $res->json();
        $url = $info['url'] ?? null;
        if (!$url) return null;

        $data = Http::timeout(120)->get($url)->body();
        if (($info['compressionAlgorithm'] ?? null) === 'GZIP') $data = gzdecode($data);
        return $data;
    }

    private function saveReportData($csv)
    {
        $path = storage_path('app/fba_report_cache.csv');
        file_put_contents($path, $csv);
        $this->info("💾 Cached at {$path}");
    }

    private function showDataPreview($csv)
    {
        $lines = explode("\n", trim($csv));
        if (!$lines) { $this->warn('⚠️ No data'); return; }

        $headers = str_getcsv(array_shift($lines), "\t");
        $this->info("📊 Records: " . count($lines) . " | Columns: " . count($headers));
        foreach (array_slice($lines, 0, 5) as $i=>$line) {
            $data = array_combine($headers, str_getcsv($line,"\t"));
            $this->comment("Sample {$i}: ".json_encode($data));
        }
    }

    private function insertInventory($csv)
    {
        $lines = explode("\n", trim($csv));
        $headers = str_getcsv(array_shift($lines), "\t");

        $this->info('🗑️ Truncating fba_table...');
        FbaTable::truncate();

        $inserted = 0;
        foreach ($lines as $line) {
            if (!trim($line)) continue;
            $row = str_getcsv($line,"\t");
            if (count($row)<count($headers)) continue;

            $data = array_combine($headers,$row);

            FbaTable::create([
                'seller_sku' => $data['sku'] ?? '',
                'fulfillment_channel_sku' => $data['fnsku'] ?? '',
                'asin' => $data['asin'] ?? '',
                'condition_type' => $data['condition'] ?? '',
                'quantity_available' => (int)($data['afn-fulfillable-quantity'] ?? 0)
            ]);
            $inserted++;
        }
        $this->info("✅ Inserted {$inserted} records");
    }

    private function fetchPrices($token, $endpoint, $marketplaceId)
    {
        $this->info('💰 Fetching prices...');

        $fbaSkus = FbaTable::pluck('seller_sku')->toArray();
        if (!$fbaSkus) { $this->warn('⚠️ No SKUs to fetch prices for'); return; }

        $reportId = $this->createReport($token, $endpoint, $marketplaceId, 'GET_MERCHANT_LISTINGS_ALL_DATA');
        if (!$reportId) return;

        $statusData = $this->waitForReport($token, $endpoint, $reportId);
        if (!$statusData) return;

        $documentId = $statusData['reportDocumentId'] ?? null;
        if (!$documentId) return;

        $csvData = $this->downloadReportDocument($token, $endpoint, $documentId);
        if (!$csvData) return;

        $lines = explode("\n", trim($csvData));
        $headers = str_getcsv(array_shift($lines), "\t");

        $this->info('🗑️ Clearing fba_prices...');
        FbaPrice::truncate();

        $inserted = 0;
        foreach ($lines as $line) {
            if (!trim($line)) continue;
            $row = str_getcsv($line,"\t");
            if (count($row)<count($headers)) continue;

            $data = array_combine($headers,$row);
            $sku = $data['seller-sku'] ?? $data['sku'] ?? '';
            $price = $data['price'] ?? $data['standard-price'] ?? 0;

            if ($sku && in_array($sku,$fbaSkus) && $price>0) {
                FbaPrice::updateOrCreate(['seller_sku'=>$sku], ['price'=>$price]);
                $inserted++;
            }
        }

        $this->info("✅ Inserted {$inserted} price records");
    }
}
