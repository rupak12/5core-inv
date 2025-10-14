<?php

namespace App\Console\Commands;

use App\Models\WaifairProductSheet;
use Illuminate\Console\Command;
use App\Http\Controllers\ApiController;
use App\Models\WayfairDataView;
use App\Models\WayfairProduct;
use Carbon\Carbon;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class SyncWayfairSheet extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'sync:wayfair-sheet';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Sync Wayfair Product Sheet';

    /**
     * Execute the console command.
     */
  

    protected $apiUrl;
    protected $apiKey;
    protected $password;

    public function __construct()
    {
        parent::__construct();
        $this->apiUrl    = "https://" . env('SHOPIFY_5CORE_DOMAIN') . "/admin/api/2024-10";
        $this->apiKey    = env('SHOPIFY_5CORE_API_KEY');
        $this->password  = env('SHOPIFY_5CORE_PASSWORD');
    }

    public function handle()
    {
        $this->info("Fetching Wayfair data from Shopify…");

        $now = Carbon::now();

        $l30Start = $now->copy()->subDays(30);
        $l30End   = $now->copy()->subDay(); // yesterday

        // L60 = 30 days before that
        $l60Start = $now->copy()->subDays(60);
        $l60End   = $now->copy()->subDays(31);

        $this->info("L30 window: {$l30Start->toDateString()} → {$l30End->toDateString()}");
        $this->info("L60 window: {$l60Start->toDateString()} → {$l60End->toDateString()}");

        $endpoint = "{$this->apiUrl}/orders.json?status=any&limit=250"
            . "&created_at_min={$l60Start->toIso8601String()}"
            . "&created_at_max={$l30End->toIso8601String()}";

        $groupedData = [];

        while ($endpoint) {
            $response = Http::withBasicAuth($this->apiKey, $this->password)->get($endpoint);

            if (!$response->successful()) {
                $this->error("Failed to fetch Shopify orders");
                $this->error("Status: " . $response->status());
                $this->error("Body: " . $response->body());
                return; 
            }

            $orders = $response->json('orders') ?? [];
            $this->info("Fetched " . count($orders) . " orders…");

            // filter Wayfair orders
            $wayfairOrders = collect($orders)->filter(function ($order) {
                $tags = strtolower($order['tags'] ?? '');
                if (str_contains($tags, 'wayfair')) {
                    return true;
                }

                if (!empty($order['note_attributes'])) {
                    foreach ($order['note_attributes'] as $attr) {
                        if (
                            strtolower($attr['name'] ?? '') === 'channel' &&
                            strtolower($attr['value'] ?? '') === 'wayfair'
                        ) {
                            return true;
                        }
                    }
                }

                $source = strtolower($order['source_name'] ?? '');
                return str_contains($source, 'wayfair');
            });

            // group SKUs
            foreach ($wayfairOrders as $order) {
                $date = Carbon::parse($order['created_at']);

                foreach ($order['line_items'] as $item) {
                    $sku = $item['sku'];
                    $qty = (int) $item['quantity'];
                    $price = $item['price'];

                    if (!isset($groupedData[$sku])) {
                        $groupedData[$sku] = [
                            'l30' => 0,
                            'l60' => 0,
                            'price' => $price,
                        ];
                    }

                    if ($date->between($l30Start, $l30End)) {
                        $groupedData[$sku]['l30'] += $qty;
                    }

                    if ($date->between($l60Start, $l60End)) {
                        $groupedData[$sku]['l60'] += $qty;
                    }

                    // update latest price
                    $groupedData[$sku]['price'] = $price;
                }
            }

            // get next page from Link header
            $linkHeader = $response->header('Link');
            if ($linkHeader && preg_match('/<([^>]+)>; rel="next"/', $linkHeader, $matches)) {
                $endpoint = $matches[1];
            } else {
                $endpoint = null;
            }
        }

        $this->info("Updating database…");

        foreach ($groupedData as $sku => $values) {
            WaifairProductSheet::updateOrCreate(
                ['sku' => $sku],
                [
                    'l30'   => $values['l30'],
                    'l60'   => $values['l60'],
                    'price' => $values['price'],
                ]
            );
        }

        $this->info("Wayfair L30/L60/Price updated successfully (" . count($groupedData) . " SKUs)");
    }





    // protected $authUrl = 'https://sso.auth.wayfair.com/oauth/token';
    // protected $graphqlUrl = 'https://api.wayfair.com/v1/graphql';
    // protected $clientId;
    // protected $clientSecret;
    // protected $audience;
    // protected $grantType = 'client_credentials';

    // public function __construct()
    // {
    //     parent::__construct();
    //     $this->clientId     = config('services.wayfair.client_id');
    //     $this->clientSecret = config('services.wayfair.client_secret');
    //     $this->audience     = config('services.wayfair.audience');
    // }

    // public function handle()
    // {
    //     $this->info("Fetching Wayfair access token...");
    //     $token = $this->getAccessToken();
    //     if (!$token) {
    //         $this->error("Failed to retrieve access token.");
    //         return;
    //     }

    //     $this->info("Fetching Wayfair purchase orders...");
    //     $orders = $this->fetchDropshipOrders($token);

    //     if (empty($orders)) {
    //         $this->warn("No orders found.");
    //         return;
    //     }

    //     $now = Carbon::now();
    //     // $l30Start = $now->copy()->subMonthNoOverflow()->startOfMonth();
    //     // $l30End   = $now->copy()->subMonthNoOverflow()->endOfMonth();
    //     // $l60Start = $now->copy()->subMonthsNoOverflow(2)->startOfMonth();
    //     // $l60End   = $now->copy()->subMonthsNoOverflow(2)->endOfMonth();

    //     $l30Start = $now->copy()->subDays(30)->startOfDay();
    //     $l30End   = $now->copy()->endOfDay();

    //     $l60Start = $now->copy()->subDays(60)->startOfDay();
    //     $l60End   = $now->copy()->subDays(31)->endOfDay();

        
    //     $skuData = [];

    //     foreach ($orders as $order) {
    //         $orderDate = isset($order['poDate']) ? Carbon::parse($order['poDate']) : null;
    //         if (!$orderDate) continue;

    //         foreach ($order['products'] ?? [] as $product) {
    //             $sku = $product['partNumber'] ?? null;
    //             if (!$sku) continue;

    //             if (!isset($skuData[$sku])) {
    //                 $skuData[$sku] = [
    //                     'l30' => 0,
    //                     'l60' => 0,
    //                     'price' => $product['price'] ?? 0,
    //                 ];
    //             }

    //             // Count L30
    //             if ($orderDate->between($l30Start, $l30End)) {
    //                 $skuData[$sku]['l30'] += $product['quantity'] ?? 0;
    //             }

    //             // Count L60
    //             if ($orderDate->between($l60Start, $l60End)) {
    //                 $skuData[$sku]['l60'] += $product['quantity'] ?? 0;
    //             }

    //             // Always update price with latest
    //             $skuData[$sku]['price'] = $product['price'] ?? $skuData[$sku]['price'];
    //         }
    //     }

    //     foreach ($skuData as $sku => $data) {
    //         WaifairProductSheet::updateOrCreate(
    //             ['sku' => $sku],
    //             [
    //                 'l30' => $data['l30'],
    //                 'l60' => $data['l60'],
    //                 'price' => $data['price'],
    //                 'updated_at' => now(),
    //             ]
    //         );

    //         $this->info("SKU {$sku} | L30={$data['l30']} | L60={$data['l60']} | Price={$data['price']}");
    //     }

    //     $this->info("Completed Wayfair Data Sheet sync.");
    // }

    // private function getAccessToken()
    // {
    //     $response = Http::withHeaders(['Content-Type' => 'application/json'])
    //         ->post($this->authUrl, [
    //             'grant_type'    => $this->grantType,
    //             'audience'      => $this->audience,
    //             'client_id'     => $this->clientId,
    //             'client_secret' => $this->clientSecret,
    //         ]);

    //     return $response->successful() ? $response->json()['access_token'] ?? null : null;
    // }

    // private function fetchDropshipOrders($token)
    // {
    //     $query = <<<'GRAPHQL'
    //     query getDropshipPurchaseOrders {
    //       getDropshipPurchaseOrders(
    //         limit: 100
    //         sortOrder: DESC
    //       ) {
    //         poNumber
    //         poDate
    //         products {
    //           partNumber
    //           name
    //           quantity
    //           price
    //         }
    //       }
    //     }
    //     GRAPHQL;

    //     $response = Http::withToken($token)->post($this->graphqlUrl, ['query' => $query]);
    //     Log::info('Wayfair Dropship API Response', $response->json());

    //     return $response->successful() ? $response->json()['data']['getDropshipPurchaseOrders'] ?? [] : [];
    // }
}
