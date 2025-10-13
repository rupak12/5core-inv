<?php

namespace App\Console\Commands;

use App\Models\TopDawgSheetdata;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class TopDawgShopSheet extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:top-dawg-shop-sheet';

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
        $url = 'https://script.google.com/macros/s/AKfycbwww2dFy8TQ0l9omjMyQrfBlHhyvdFDz8v7unLbVhk2XSqCyRIcyN_0XFiUvsdzvtPcWA/exec';

        $this->info('Starting Top Dawg Sheet sync...');

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

                $price = $item['price'] ?? null;
                if (!is_numeric($price)) {
                    $price = null; // ignore invalid like #REF!, text, or blanks
                }

                $views = $item['views'] ?? null;
                if (!is_numeric($views)) {
                    $views = null;
                }

                TopDawgSheetdata::updateOrCreate(
                    ['sku' => trim($sku)],
                    [
                        'l30'  => $item['l30'] ?? null,
                        'l60'  => $item['l60'] ?? null,
                        'price'  => $price,
                        'views'  => $views,
                    ]
                );

                $savedCount++;
            }

            $this->info("Top Dawg Sheet data synced successfully! Total records saved/updated: {$savedCount}");
        } catch (\Exception $e) {
            Log::error('Error syncing Top Dawg Sheet: ' . $e->getMessage());
            $this->error('Error: ' . $e->getMessage());
        }
    }
}
