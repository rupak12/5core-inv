<?php

namespace App\Console\Commands;

use App\Models\BestbuyUsaProduct;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Cache;
use App\Models\MacyProduct;
use App\Models\TiendamiaProduct;
use Carbon\Carbon;
use Illuminate\Support\Facades\Log;

class FetchMacyProducts extends Command
{
    /**  
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:fetch-macy-products';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Fetch and store Macy products data';

    /**
     * Execute the console command.
     */
    // public function handle()
    // {
    //     $token = $this->getAccessToken();
    //     if (!$token) return;

    //     // Step 1: Mass-fetch all orders once
    //     $skuSales = $this->getSalesTotals($token); // ['sku' => ['m_l30' => 12, 'm_l60' => 4]]

    //     // Step 2: Paginate through products
    //     $pageToken = null;
    //     $page = 1;

    //     do {
    //         $this->info("Fetching product page $page...");
    //         $url = 'https://miraklconnect.com/api/products?limit=1000';
    //         if ($pageToken) {
    //             $url .= '&page_token=' . urlencode($pageToken);
    //         }

    //         $response = Http::withToken($token)->get($url);
    //         if (!$response->successful()) {
    //             $this->error('Product fetch failed: ' . $response->body());
    //             return;
    //         }

    //         $json = $response->json();
    //         $products = $json['data'] ?? [];
    //         $pageToken = $json['next_page_token'] ?? null;

    //         foreach ($products as $product) {
    //             $sku = $product['id'] ?? null;
    //             $price = $product['discount_prices'][0]['price']['amount'] ?? null;

    //             if (!$sku || $price === null) continue;

    //             $m_l30 = $skuSales[$sku]['m_l30'] ?? 0;
    //             $m_l60 = $skuSales[$sku]['m_l60'] ?? 0;

    //             MacyProduct::updateOrCreate(
    //                 ['sku' => $sku],
    //                 [
    //                     'price' => $price,
    //                     'm_l30' => $m_l30,
    //                     'm_l60' => $m_l60,
    //                 ]
    //             );
    //         }

    //         $page++;
    //     } while ($pageToken);

    //     $this->info("All Macy products stored successfully");
    // }

    // private function getAccessToken()
    // {
    //     return Cache::remember('macy_access_token', 3500, function () {
    //         $response = Http::asForm()->post('https://auth.mirakl.net/oauth/token', [
    //             'grant_type' => 'client_credentials',
    //             'client_id' => config('services.macy.client_id'),
    //             'client_secret' => config('services.macy.client_secret'),
    //         ]);

    //         return $response->successful()
    //             ? $response->json()['access_token']
    //             : null;
    //     });
    // }

    // private function getSalesTotals(string $token): array
    // {
    //     $this->info("Fetching all orders in last 60 days...");

    //     $orders = [];
    //     $pageToken = null;
    //     $startDate = now()->subDays(60)->toIso8601String(); // ISO format for query param

    //     do {
    //         $url = 'https://miraklconnect.com/api/v2/orders?fulfillment_type=FULFILLED_BY_SELLER&limit=100';
    //         $url .= '&updated_from=' . urlencode($startDate);
    //         if ($pageToken) {
    //             $url .= '&page_token=' . urlencode($pageToken);
    //         }

    //         $response = Http::withToken($token)->get($url);
    //         if (!$response->successful()) {
    //             $this->error("Order fetch failed: " . $response->body());
    //             break;
    //         }

    //         $json = $response->json();
    //         $orders = array_merge($orders, $json['data'] ?? []);
    //         $pageToken = $json['next_page_token'] ?? null;
    //     } while ($pageToken);

    //     $this->info("Orders fetched: " . count($orders));

    //     // Define date ranges
    //     $now = now();
    //     $startL30 = $now->copy()->subDays(30);
    //     $endL30 = $now->copy()->subDay();

    //     $startL60 = $now->copy()->subDays(60);
    //     $endL60 = $now->copy()->subDays(31);

    //     // Initialize sku map
    //     $sales = [];

    //     foreach ($orders as $order) {
    //         $created = Carbon::parse($order['created_at']);

    //         foreach ($order['order_lines'] ?? [] as $line) {
    //             $sku = $line['product']['id'] ?? null;
    //             $qty = $line['quantity'] ?? 0;

    //             if (!$sku) continue;

    //             if (!isset($sales[$sku])) {
    //                 $sales[$sku] = ['m_l30' => 0, 'm_l60' => 0];
    //             }

    //             if ($created->between($startL60, $endL60)) {
    //                 $sales[$sku]['m_l60'] += $qty;
    //             } elseif ($created->between($startL30, $endL30)) {
    //                 $sales[$sku]['m_l30'] += $qty;
    //             }
    //         }
    //     }

    //     return $sales;
    // }


    // private function getSalesTotals(string $token): array
    // {
    //     $this->info("Fetching all Macy orders in last 60 days...");

    //     $sales = [];
    //     $pageToken = null;
    //     $startDate = now()->subDays(60)->startOfDay()->toIso8601String();

    //     // Define L30 and L60 ranges once
    //     $now = now();
    //     $startL30 = $now->copy()->subDays(30)->startOfDay();
    //     $endL30 = $now->copy()->endOfDay();

    //     $startL60 = $now->copy()->subDays(60)->startOfDay();
    //     $endL60 = $now->copy()->subDays(31)->endOfDay();

    //     do {
    //         $url = 'https://miraklconnect.com/api/v2/orders?fulfillment_type=FULFILLED_BY_SELLER&limit=100&created_from=' . urlencode($startDate);
    //         if ($pageToken) {
    //             $url .= '&page_token=' . urlencode($pageToken);
    //         }

    //         $response = Http::withToken($token)->get($url);
    //         if (!$response->successful()) {
    //             $this->error("Order fetch failed: " . $response->body());
    //             break;
    //         }

    //         $json = $response->json();
    //         $orders = $json['data'] ?? [];
    //         $pageToken = $json['next_page_token'] ?? null;

    //         // $channelNames = array_map(function($orders) {
    //         //     return $orders['origin']['channel_name'];
    //         // }, $orders);
    //         // if($channelNames == "Macy's, Inc."){
    //         //     Log::info('No orders found in this page.');
    //         // } else {
    //         //     Log::info('Channel names in this page: ' . implode(', ', $channelNames));
    //         // }

    //         foreach ($orders as $order) {
    //             $created = Carbon::parse($order['created_at']);

    //             foreach ($order['order_lines'] ?? [] as $line) {
    //                 $sku = $line['product']['id'] ?? null;
    //                 $qty = $line['quantity'] ?? 0;

    //                 if (!$sku) continue;

    //                 if (!isset($sales[$sku])) {
    //                     $sales[$sku] = ['m_l30' => 0, 'm_l60' => 0];
    //                 }

    //                 if ($created->between($startL30, $endL30, true)) {
    //                     $sales[$sku]['m_l30'] += $qty;
    //                 } elseif ($created->between($startL60, $endL60, true)) {
    //                     $sales[$sku]['m_l60'] += $qty;
    //                 }
    //             }
    //         }

    //         $this->info("Processed " . count($orders) . " orders in this page...");

    //     } while ($pageToken);

    //     return $sales;
    // }


    // private function getSalesTotals(string $token): array
    // {
    //     $this->info("Fetching Macy orders in last 60 days...");

    //     $pageToken = null;
    //     $sales = [];

    //     $now = now('America/New_York');
    //     $startDate = $now->copy()->subDays(60)->startOfDay()->toIso8601String();

    //     $startL30 = $now->copy()->subDays(29)->startOfDay();
    //     $endL30   = $now->copy()->endOfDay();
    //     $startL60 = $now->copy()->subDays(59)->startOfDay();
    //     $endL60   = $now->copy()->subDays(30)->endOfDay();

    //     do {
    //         $url = 'https://miraklconnect.com/api/v2/orders'
    //             . '?fulfillment_type=FULFILLED_BY_SELLER'
    //             . '&limit=100'
    //             . '&created_from=' . urlencode($startDate);

    //         if ($pageToken) {
    //             $url .= '&page_token=' . urlencode($pageToken);
    //         }

    //         $response = Http::withToken($token)->get($url);

    //         if (!$response->successful()) {
    //             $this->error("Order fetch failed: " . $response->body());
    //             break;
    //         }

    //         $json = $response->json();
    //         $pageOrders = $json['data'] ?? [];
    //         $pageToken = $json['next_page_token'] ?? null;

    //         // Filter only Macy's orders
    //         $macysOrders = array_filter($pageOrders, function($order) {
    //             return isset($order['origin']['channel_name']) && $order['origin']['channel_name'] === "Macy's, Inc.";
    //         });
    //         dd($macysOrders);

    //         foreach ($macysOrders as $order) {
    //             $created = Carbon::parse($order['created_at'], 'America/New_York');

    //             foreach ($order['order_lines'] ?? [] as $line) {
    //                 $sku = $line['product']['id'] ?? null;
    //                 $qty = $line['quantity'] ?? 0;

    //                 if (!$sku) continue;

    //                 if (!isset($sales[$sku])) {
    //                     $sales[$sku] = ['m_l30' => 0, 'm_l60' => 0];
    //                 }

    //                 if ($created->between($startL30, $endL30)) {
    //                     $sales[$sku]['m_l30'] += $qty;
    //                 } elseif ($created->between($startL60, $endL60)) {
    //                     $sales[$sku]['m_l60'] += $qty;
    //                 }
    //             }
    //         }

    //         Log::info("Processed page with " . count($macysOrders) . " Macy's orders.");
    //     } while ($pageToken);

    //     $this->info("Total Macy's SKUs: " . count($sales));

    //     return $sales;
    // }


    public function handle()
    {
        $token = $this->getAccessToken();
        if (!$token) return;

        $skuSales = $this->getSalesTotals($token); // ['sku' => ['m_l30'=>x,'m_l60'=>y]]

        $pageToken = null;
        $page = 1;
        $allProducts = [];

         do {
            $this->info("Fetching product page $page...");

            $url = 'https://miraklconnect.com/api/products?limit=1000';
            if ($pageToken) {
                $url .= '&page_token=' . urlencode($pageToken);
            }

            $response = Http::withToken($token)->get($url);
            if (!$response->successful()) {
                $this->error('Product fetch failed: ' . $response->body());
                return;
            }

            $json = $response->json();
            $products = $json['data'] ?? [];
            $pageToken = $json['next_page_token'] ?? null;
            $allProducts = array_merge($allProducts, $products);

            foreach ($products as $product) {
                $sku = $product['id'] ?? null;
                $price = $product['discount_prices'][0]['price']['amount'] ?? null;
                if (!$sku || $price === null) continue;

                $originalSku = $sku;
                $sku = strtolower($sku);

                // Loop through all channels that reported this SKU
                foreach ($skuSales as $channel => $skuMap) {
                    if (!isset($skuMap[$sku])) continue;

                    $l30 = $skuMap[$sku]['l30'];
                    $l60 = $skuMap[$sku]['l60'];

                    switch ($channel) {
                        case "Macy's, Inc.":
                            MacyProduct::updateOrCreate(
                                ['sku' => $originalSku],
                                ['price' => $price, 'm_l30' => $l30, 'm_l60' => $l60]
                            );
                            break;

                        case "Tiendamia":
                            TiendamiaProduct::updateOrCreate(
                                ['sku' => $originalSku],
                                ['price' => $price, 'm_l30' => $l30, 'm_l60' => $l60]
                            );
                            break;

                        case "Best Buy USA":
                            // Removed to store all products separately
                            break;

                        default:
                            Log::warning("Unknown channel: {$channel} for SKU {$originalSku}");
                            break;
                    }

                    Log::info("Stored {$channel} | SKU: {$originalSku}, Price: {$price}, L30: {$l30}, L60: {$l60}");
                }
            }

            $page++;
        } while ($pageToken);

        // Store all products in Best Buy table
        foreach ($allProducts as $product) {
            $sku = $product['id'] ?? null;
            $price = $product['discount_prices'][0]['price']['amount'] ?? null;
            if (!$sku || $price === null) continue;

            $originalSku = $sku;
            $sku = strtolower($sku);

            if ($originalSku === 'CDKC13 1pc') {
                Log::info("Storing CDKC13 1pc in Best Buy");
            }

            $l30 = $skuSales['Best Buy USA'][$sku]['l30'] ?? 0;
            $l60 = $skuSales['Best Buy USA'][$sku]['l60'] ?? 0;

            BestbuyUsaProduct::updateOrCreate(
                ['sku' => $originalSku],
                ['price' => $price, 'm_l30' => $l30, 'm_l60' => $l60]
            );

            Log::info("Stored Best Buy | SKU: {$originalSku}, Price: {$price}, L30: {$l30}, L60: {$l60}");
        }

        $this->info("All Macy, Tiendamia, BestbuyUSA products stored successfully.");
    }

    private function getAccessToken()
    {
        // return Cache::remember('macy_access_token', 3500, function () {
            $response = Http::asForm()->post('https://auth.mirakl.net/oauth/token', [
                'grant_type' => 'client_credentials',
                'client_id' => config('services.macy.client_id'),
                'client_secret' => config('services.macy.client_secret'),
            ]);

            return $response->successful() ? $response->json()['access_token'] : null;
        // });
    }

    private function getSalesTotals(string $token): array
    {
        $this->info("Fetching Macy, Tiendamia, BestbuyUSA orders in last 60 days...");

        $pageToken = null;
        $sales = [];

        $now = now('America/New_York');
        $startDate = $now->copy()->subDays(60)->startOfDay()->toIso8601String();

        $startL30 = $now->copy()->subDays(29)->startOfDay();
        $endL30   = $now->copy()->endOfDay();
        $startL60 = $now->copy()->subDays(59)->startOfDay();
        $endL60   = $now->copy()->subDays(30)->endOfDay();

        do {
            $url = 'https://miraklconnect.com/api/v2/orders'
                . '?fulfillment_type=FULFILLED_BY_SELLER'
                . '&limit=100'
                . '&created_from=' . urlencode($startDate);

            if ($pageToken) {
                $url .= '&page_token=' . urlencode($pageToken);
            }

            $response = Http::withToken($token)->get($url);
            if (!$response->successful()) {
                $this->error("Order fetch failed: " . $response->body());
                break;
            }

            $json = $response->json();
            $orders = $json['data'] ?? [];
            $pageToken = $json['next_page_token'] ?? null;

            foreach ($orders as $order) {
                $channel = $order['origin']['channel_name'] ?? 'UNKNOWN';
                $created = Carbon::parse($order['created_at'], 'America/New_York');

                foreach ($order['order_lines'] ?? [] as $line) {
                    $sku = $line['product']['id'] ?? null;
                    $qty = $line['quantity'] ?? 0;
                    if (!$sku) continue;

                    $sku = strtolower($sku);

                    if (str_contains($sku, 'cdkc13') && $channel === "Best Buy USA") {
                        Log::info("Found SKU containing cdkc13 in Best Buy order: {$sku}, qty {$qty}, created_at {$order['created_at']}");
                    }

                    if (!isset($sales[$channel][$sku])) {
                        $sales[$channel][$sku] = ['l30' => 0, 'l60' => 0];
                    }

                    if ($created->between($startL30, $endL30)) {
                        $sales[$channel][$sku]['l30'] += $qty;
                    } elseif ($created->between($startL60, $endL60)) {
                        $sales[$channel][$sku]['l60'] += $qty;
                    }
                }
            }
        } while ($pageToken);


        $this->info("Total Macy's SKUs: " . count($sales));
        // foreach ($sales as $channel => $skuMap) {
        //     $this->info("Channel {$channel} has " . count($skuMap) . " SKUs with orders.");
        // }

        if (isset($sales['Best Buy USA'])) {
            foreach ($sales['Best Buy USA'] as $sku => $data) {
                Log::info("Best Buy SKU: {$sku}, L30: {$data['l30']}, L60: {$data['l60']}");
            }
        }

        return $sales;
    }



}
