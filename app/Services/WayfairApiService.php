<?php

namespace App\Services;

use Illuminate\Support\Facades\Http;

class WayfairApiService
{
    protected $token;
 protected $authUrl = 'https://sso.auth.wayfair.com/oauth/token';
    protected $graphqlUrl = 'https://api.wayfair.com/v1/graphql';
     protected $clientId;
    protected $clientSecret;
    protected $audience;
    protected $accessToken;
    protected $grantType = 'client_credentials';

    public function __construct()
    {
        $this->authenticate();
  

        $this->clientId = config('services.wayfair.client_id');
        $this->clientSecret = config('services.wayfair.client_secret');
        $this->audience = config('services.wayfair.audience');
    }

    /**
     * Authenticate with Wayfair and get access token
     */
    protected function authenticate()
    {
        $response = Http::withoutVerifying()->asForm()->post('https://sso.auth.wayfair.com/oauth/token', [
            'grant_type'    => 'client_credentials',
            'client_id'     => env('WAYFAIR_CLIENT_ID'),
            'client_secret' => env('WAYFAIR_CLIENT_SECRET'),
        ]);

        if ($response->failed()) {
            throw new \Exception('Failed to authenticate with Wayfair API: ' . $response->body());
        }

        return $response->json('access_token');
    }

    public function updatePrice(string $sku, float $price)
    {
        // Build XML for pricing feed
        $xml = <<<XML
<?xml version="1.0" encoding="UTF-8"?>
<PriceFeed xmlns="http://api.wayfair.com/v1/pricefeed.xsd">
    <Price>
        <SupplierPartNumber>{$sku}</SupplierPartNumber>
        <PriceAmount>{$price}</PriceAmount>
        <CurrencyCode>USD</CurrencyCode>
    </Price>
</PriceFeed>
XML;

        $response = Http::withToken($this->authenticate())
            ->attach('file', $xml, 'price_feed.xml')
            ->post('https://api.wayfair.com/v1/feeds/pricing');

        return $response->json();
    }



     private function getAccessToken()
    {
        $response = Http::withoutVerifying()->asForm()->post($this->authUrl, [
            'grant_type' => $this->grantType,
            'audience' => $this->audience,
            'client_id' => $this->clientId,
            'client_secret' => $this->clientSecret,
        ]);

        return $response->successful() ? ($response->json()['access_token'] ?? null) : null;
    }


      public function getInventory()
{
    $accessToken = $this->getAccessToken();
    if (!$accessToken) {
        throw new \Exception('Missing Wayfair access token');
    }

    $allProducts = [];
    $page = 1;
    $perPage = 50;

    do {
        $response = Http::withoutVerifying()->withToken($accessToken)
            ->get('https://api.wayfair.com/v1/catalog/products', [
                'page' => $page,
                'per_page' => $perPage
            ]);

        if ($response->failed()) {
            $errorBody = $response->json();
            $errorMessage = $errorBody['message'] ?? json_encode($errorBody);
            throw new \Exception('Wayfair API Error: ' . $errorMessage);
        }

        $data = $response->json();
        dd($data);

        if (!empty($data['products'])) {
            $allProducts = array_merge($allProducts, $data['products']);
        }

        $page++;
        $totalPages = $data['total_pages'] ?? 1;
    } while ($page <= $totalPages);

    return $allProducts;
}

    
   public function getInventory1()
{
    $limit = 100;
    $offset = 0;
    $inventoryUrl = 'https://api.wayfair.com/v1/graphql';

    $query = <<<'GRAPHQL'
    query {
        inventory {
            supplierPartNumber
            quantityOnHand
            quantityBackordered
            quantityOnOrder
            itemNextAvailabilityDate
            discontinued
        }
    }
    GRAPHQL;

    $response = Http::withoutVerifying()->withToken($this->getAccessToken())->post($inventoryUrl, [
        'query' => $query,
        'variables' => [
            'limit' => $limit,
            'offset' => $offset,
        ]
    ]);

    dd($response->json());
    if (!$response->successful()) {
        throw new \Exception("Wayfair API Error: " . $response->body());
    }

    $inventoryItems = $response->json()['data']['inventory'] ?? [];

    // Extract SKU and quantity
    $result = array_map(function ($item) {
        return [
            'sku' => $item['supplierPartNumber'] ?? null,
            'quantity' => $item['quantityOnHand'] ?? 0,
        ];
    }, $inventoryItems);

    return $result;
}
}
