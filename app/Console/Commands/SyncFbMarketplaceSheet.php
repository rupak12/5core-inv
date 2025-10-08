<?php

namespace App\Console\Commands;

use App\Models\FbMarketplaceSheetdata;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class SyncFbMarketplaceSheet extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:sync-fb-marketplace-sheet';

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
        $url = 'https://script.google.com/macros/s/AKfycbzMKE_Ri8XbcMUDv_ad0FzcCbjew04hU0u25AuHOrt7llkMoNIxRf0Arou7p4Bs354B/exec';

        $this->info('Starting Facebook Marketplace sheet sync...');

        try {
            $response = Http::get($url);

            if ($response->failed()) {
                $this->error('Failed to fetch data from Google Sheet.');
                return;
            }

            $payload = $response->json();

            //  Auto-detect structure (either {data: [...]} or [...])
            $data = isset($payload['data']) ? $payload['data'] : $payload;

            if (!is_array($data) || empty($data)) {
                $this->warn('No valid data found.');
                return;
            }

            $savedCount = 0;

            foreach ($data as $item) {
                $sku = $item['sku'] ?? null;
                if (empty($sku)) {
                    $this->warn("Skipped a row with missing SKU.");
                    continue;
                }

                FbMarketplaceSheetdata::updateOrCreate(
                    ['sku' => trim($sku)],
                    [
                        'l30'  => $item['l30'] ?? null,
                        'l60'  => $item['l60'] ?? null,
                        'price'  => isset($item['price']) && $item['price'] !== '' ? $item['price'] : null,
                        'views'  => isset($item['views']) && $item['views'] !== '' ? (int)$item['views'] : null,
                    ]
                );

                $savedCount++;
            }

            $this->info("Facebook Marketplace Sheet data synced successfully! Total records saved/updated: {$savedCount}");
        } catch (\Exception $e) {
            Log::error('Error syncing Facebook Marketplace sheet: ' . $e->getMessage());
            $this->error('Error: ' . $e->getMessage());
        }
    }
}
