<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Carbon\Carbon;
use App\Models\WayfairProduct;

class FetchWayfairData extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'app:fetch--data';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Fetch and store Wayfair data into the database';


    protected $authUrl = 'https://sso.auth..com/oauth/token';
    protected $graphqlUrl = 'https://api..com/v1/graphql';

    protected $clientId;
    protected $clientSecret;
    protected $audience;
    protected $accessToken;
    protected $grantType = 'client_credentials';

    public function __construct()
    {
        parent::__construct();

        $this->clientId = config('services..client_id');
        $this->clientSecret = config('services..client_secret');
        $this->audience = config('services..audience');
    }

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $this->info("Fetching Wayfair access token...");

        $token = $this->getAccessToken();
        if (!$token) {
            $this->error("Failed to retrieve access token.");
            return;
        }

        $this->info("Truncating old Wayfair data...");
        WayfairProduct::truncate(); 

        $this->info("Access token received. Fetching all purchase orders...");

        $limit = 100;
        $offset = 0;
        $totalSaved = 0;

        while (true) {
            $purchaseOrders = $this->fetchPurchaseOrders($token, $limit, $offset);

            if (empty($purchaseOrders)) {
                break;
            }

            foreach ($purchaseOrders as $po) {
                $poJson = json_encode($po);
                $products = $po['products'] ?? [];

                foreach ($products as $product) {
                    $sku = $product['partNumber'] ?? null;
                    if (!$sku) continue;

                    // Save to DB
                    WayfairProduct::create([
                        'sku' => $sku,
                        'purchase_order_data' => $poJson
                    ]);

                    $totalSaved++;
                    $this->info("Saved SKU: $sku | PO: " . ($po['poNumber'] ?? 'N/A'));
                }
            }

            // Next page
            $offset += $limit;
        }

        $this->info("All done. Total records saved: $totalSaved");
    }


    private function getAccessToken()
    {
        $response = Http::asForm()->post($this->authUrl, [
            'grant_type' => $this->grantType,
            'audience' => $this->audience,
            'client_id' => $this->clientId,
            'client_secret' => $this->clientSecret,
        ]);

        return $response->successful() ? ($response->json()['access_token'] ?? null) : null;
    }

    private function fetchPurchaseOrders($token, $limit, $offset)
    {
        $query = <<<'GRAPHQL'
        query GetPurchaseOrders($limit: Int!, $offset: Int!) {
            purchaseOrders(
                filters: [{ field: open, equals: "true" }],
                limit: $limit,
                offset: $offset
            ) {
                poNumber
                poDate
                estimatedShipDate
                customerName
                customerAddress1
                customerAddress2
                customerCity
                customerState
                customerPostalCode
                shippingInfo {
                    shipSpeed
                    carrierCode
                }
                packingSlipUrl
                warehouse {
                    id
                    name
                    address {
                        name
                        address1
                        address2
                        address3
                        city
                        state
                        country
                        postalCode
                    }
                }
                products {
                    partNumber
                    quantity
                    price
                    event {
                        id
                        type
                        name
                        startDate
                        endDate
                    }
                }
                shipTo {
                    name
                    address1
                    address2
                    address3
                    city
                    state
                    country
                    postalCode
                    phoneNumber
                }
            }
        }
        GRAPHQL;

        $response = Http::withToken($token)->post($this->graphqlUrl, [
            'query' => $query,
            'variables' => [
                'limit' => $limit,
                'offset' => $offset,
            ]
        ]);

        return $response->successful() ? ($response->json()['data']['purchaseOrders'] ?? []) : [];
    }


}
