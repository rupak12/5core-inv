<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use App\Models\ReverbProduct;
use Carbon\Carbon;

class FetchReverbData extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'reverb:fetch';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Fetch Reverb listing data and store in database daily';

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $this->info('Fetching Reverb Listings...');
        $listings = $this->fetchAllListings();

        $today = Carbon::today();
        // $l30Start = $today->copy()->subDays(30);
        // $l30End   = $today->copy()->subDay();
        // $l60Start = $today->copy()->subDays(60);
        // $l60End   = $l30Start->copy()->subDay();

        $l30End   = $today->copy()->subDay();     // 2025-09-25
        $l30Start = $l30End->copy()->subDays(29); // 2025-08-27

        $l60End   = $l30Start->copy()->subDay();  // 2025-08-26
        $l60Start = $l60End->copy()->subDays(59); // 2025-07-29


        // $rL30 = $this->getOrderQuantities(30);
        // $rL60 = $this->getOrderQuantities(60);

        $rL30 = $this->getOrderQuantities($l30Start, $l30End);
        $rL60 = $this->getOrderQuantities($l60Start, $l60End);

        foreach ($listings as $item) {
            $sku = $item['sku'] ?? null;

            if (!$sku) {
                $this->warn("Skipping missing SKU or ID");
                continue;
            }

            $r30 = $rL30[$sku]['qty'] ?? 0;
            $r60 = $rL60[$sku]['qty'] ?? 0;
            $status = $rL30[$sku]['status'] ?? null;

            // $this->line("Listing SKU: $sku | R_L30: $r30 | R_L60: $r60");

            // Store record
            ReverbProduct::updateOrCreate(
            ['sku' => $sku], // Match on SKU
            [
                'sku' => $sku,
                'r_l30' => $r30,
                'r_l60' => $r60,
                'status' => $status,
                'price' => $item['price']['amount'] ?? null,
                'views' => $item['stats']['views'] ?? null,
            ]);
        }

        $this->info('Reverb data stored successfully.');
    }

    protected function fetchAllListings(): array
    {
        $listings = [];
        $url = 'https://api.reverb.com/api/my/listings';

        do {
            $response = Http::withHeaders([
                'Authorization' => 'Bearer ' . config('services.reverb.token'),
                'Accept' => 'application/hal+json',
                'Accept-Version' => '3.0',
            ])->get($url);

            if ($response->failed()) {
                $this->error('Failed to fetch listings.');
                break;
            }

            $data = $response->json();
            $listings = array_merge($listings, $data['listings'] ?? []);
            $url = $data['_links']['next']['href'] ?? null;

        } while ($url);

        $this->info('Fetched total listings: ' . count($listings));
        return $listings;
    }

    protected function getOrderQuantities(Carbon $startDate, Carbon $endDate): array
    {
        $this->info("Fetching orders from {$startDate->toDateString()} to {$endDate->toDateString()}...");

        
        $url = "https://api.reverb.com/api/my/orders/selling/all?updated_start_date={$startDate->toIso8601String()}&updated_end_date={$endDate->toIso8601String()}";
        $quantityMap = [];

        do {
            $response = Http::withHeaders([
                'Authorization' => 'Bearer ' . config('services.reverb.token'),
                'Accept' => 'application/hal+json',
                'Accept-Version' => '3.0',
            ])->get($url);

            if ($response->failed()) {
                $this->error("Failed to fetch orders");
                return [];
            }

            $orders = $response->json()['orders'] ?? [];
            foreach ($orders as $order) {
                $sku = $order['sku'] ?? null;
                $qty = $order['quantity'] ?? 0;
                $status = $order['status'] ?? ($order['status'] ?? null); 

            if ($sku) {
                if (!isset($quantityMap[$sku])) {
                        $quantityMap[$sku] = ['qty' => 0, 'status' => $status];
                }

                    $quantityMap[$sku]['qty'] += $qty;
                    $quantityMap[$sku]['status'] = $status;
                }
                    // $this->line("Order SKU: $sku, Qty: $qty");
                    // $quantityMap[$sku] = ($quantityMap[$sku] ?? 0) + $qty;
                
            }

            $url = $response->json()['_links']['next']['href'] ?? null;
        } while ($url);

        $this->info("Orders processed from {$startDate->toDateString()} to {$endDate->toDateString()}.");
        return $quantityMap;
    }

    
}
