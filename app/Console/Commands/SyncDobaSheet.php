<?php

namespace App\Console\Commands;

use App\Models\DobaSheetdata;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;

class SyncDobaSheet extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:sync-doba-sheet';

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
        $this->info('Fetching DOBA Sheet data from Google Sheets...');

        // Same URL as in your controller
        $url = 'https://script.google.com/macros/s/AKfycbwSe2tvYfvb5_0uWK6YWumP7x6lpW90jtkL2DYEMhjMH6uNzJB27qjwEYdVe4QK3vHeIg/exec';

        try {
            $response = Http::timeout(120)->get($url);

            if (!$response->successful()) {
                $this->error('Failed to fetch data. Status: ' . $response->status());
                return;
            }

            $data = $response->json();

            if (!is_array($data) || count($data) === 0) {
                $this->warn('No data received from Google Sheet.');
                return;
            }

            $count = 0;

            foreach ($data as $row) {
                // Adjust column names to your actual sheet keys
                $sku = $row['(Child) sku'] ?? null;

                if (empty($sku)) {
                    $this->warn('Skipped row with missing SKU.');
                    continue;
                }

                $l30          = isset($row['T L30']) && $row['T L30'] !== '' ? (int)$row['T L30'] : null;
                $l60          = isset($row['D L60']) && $row['D L60'] !== '' ? (int)$row['D L60'] : null;
                $views        = isset($row['Click 90']) && $row['Click 90'] !== '' ? (int)$row['Click 90'] : null;
                $price        = isset($row['Doba AD']) && $row['Doba AD'] !== '' ? (float)$row['Doba AD'] : null;
                $pickup_price = isset($row['PICK UP PRICE ']) && $row['PICK UP PRICE '] !== '' ? (float)$row['PICK UP PRICE '] : null;
                $item_id      = isset($row['Item Id']) && $row['Item Id'] !== '' ? trim($row['Item Id']) : null;


                DobaSheetdata::updateOrCreate(
                    ['sku' => $sku],
                    [
                        'l30'          => $l30,
                        'l60'          => $l60,
                        'price'        => $price,
                        'views'        => $views,
                        'pickup_price' => $pickup_price,
                        'item_id'      => $item_id,
                    ]
                );

                $count++;
            }

            $this->info("DOBA Sheet data synced successfully! Total records saved/updated: {$count}");
        } catch (\Exception $e) {
            $this->error('Exception: ' . $e->getMessage());
        }
    }
}
