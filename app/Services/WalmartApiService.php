<?php

namespace App\Services;

use Exception;
use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Str;
use App\Models\ProductStockMapping;

class WalmartApiService
{
    protected $clientId;
    protected $clientSecret;
    protected $baseUrl;
    protected $marketplaceId;
    protected $token;

    public function __construct()
    {
        $this->clientId       = env('WALMART_CLIENT_ID');
        $this->clientSecret   = env('WALMART_CLIENT_SECRET');
        $this->baseUrl        = env('WALMART_API_ENDPOINT', 'https://marketplace.walmartapis.com');
        $this->marketplaceId  = env('WALMART_MARKETPLACE_ID', 'WMTMP');
        $this->token          = $this->getAccessToken();
    }


    public function getAccessToken()
    {
        $clientId     = env('WALMART_CLIENT_ID');
        $clientSecret = env('WALMART_CLIENT_SECRET');

        $authorization = base64_encode("{$clientId}:{$clientSecret}");

        $response = Http::withoutVerifying()->asForm()->withHeaders([
            'Authorization'          => "Basic {$authorization}",
            'WM_QOS.CORRELATION_ID'  => uniqid(),
            'WM_SVC.NAME'            => 'Walmart Marketplace',
            'accept'                 => 'application/json',
        ])->post('https://marketplace.walmartapis.com/v3/token', [
            'grant_type' => 'client_credentials',
        ]);

        if ($response->successful()) {
            return $response->json()['access_token'] ?? null;
        }

        // dd($response->json());

        return null;
    }

  public function getAccessTokenV1(): ?string
{
    $clientId     = env('WALMART_CLIENT_ID');
    $clientSecret = env('WALMART_CLIENT_SECRET');

    if (!$clientId || !$clientSecret) {
        Log::error('Walmart credentials missing.');
        return null;
    }

    $authorization = base64_encode("{$clientId}:{$clientSecret}");

    $response = Http::withoutVerifying()->asForm()->withHeaders([
        'Authorization'         => "Basic {$authorization}",
        'WM_QOS.CORRELATION_ID' => "123",
        'WM_SVC.NAME'           => 'Walmart Marketplace',
        'Accept'                => 'application/json',
        'Content-Type'          => 'application/x-www-form-urlencoded',
    ])->post('https://marketplace.walmartapis.com/v3/token', [
        'grant_type' => 'client_credentials',
    ]);

    if ($response->successful()) {
        return $response->json()['access_token'] ?? null;
    }

    Log::error('Failed to fetch Walmart access token', [
        'status' => $response->status(),
        'body' => $response->json(),
    ]);

    return null;
}



    public function updatePrice(string $sku, float $price): array
    {
        $accessToken = $this->getAccessToken();

        $payload = [
            'sku' => $sku,
            'pricing' => [
                [
                    'currentPriceType' => 'BASE',
                    'currentPrice' => [
                        'currency' => 'USD',
                        'amount' => number_format($price, 2, '.', '')
                    ]
                ]
            ]
        ];

        $endpoint = $this->baseUrl . "/v3/price";

        $response = Http::withHeaders([
            'WM_QOS.CORRELATION_ID' => uniqid(),
            'WM_SEC.ACCESS_TOKEN' => $accessToken,
            'WM_SVC.NAME' => 'Walmart Marketplace',
            'Accept' => 'application/json',
            'Content-Type' => 'application/json',
        ])->put($endpoint, $payload);

        if ($response->failed()) {
            throw new Exception('Failed to update Walmart price: ' . $response->body());
        }
        Log::info('Walmart Price Update Response: ', $response->json());
        return $response->json();
    }

 public function getinventory(): array
{
    $accessToken = $this->getAccessToken();
    if (!$accessToken) {
        throw new \Exception('Failed to retrieve Walmart access token.');
    }

    $endpoint = $this->baseUrl . '/v3/inventories';
    $limit = 50;
    $cursor = null;
    $collected = [];

    do {
        $query = ['limit' => $limit];
        if ($cursor) {
            $query['nextCursor'] = $cursor;
        }

        $headers = [
            'WM_SEC.ACCESS_TOKEN'   => $accessToken,
            'WM_QOS.CORRELATION_ID' => (string) Str::uuid(),
            'WM_SVC.NAME'           => 'Walmart Marketplace',
            'Accept'                => 'application/json',
        ];

        $request = Http::withHeaders($headers);
        if (env('FILESYSTEM_DRIVER') === 'local') {
            $request = $request->withoutVerifying();
        }

        $response = $request->get($endpoint, $query);

        if ($response->failed()) {
            Log::error('Walmart inventory fetch failed', [
                'status' => $response->status(),
                'body' => $response->body(),
            ]);
            throw new \Exception('Failed to fetch Walmart inventory');
        }

        $json = $response->json();
        $items = $json['elements'] ?? [];
        $collected = array_merge($collected, $items);
        $cursor = $json['meta']['nextCursor'] ?? null;

    } while ($cursor);

    // Process the collected inventory data
    $collected=$collected['inventories'];
    // dd($collected);
    foreach ($collected as $item) {
        $sku = $item['sku'] ?? null;
        $quantity = 0;
        
        // Extract quantity from the first node's available to sell amount
        if (isset($item['nodes'][0]['availToSellQty']['amount'])) {
            $quantity = (int) $item['nodes'][0]['availToSellQty']['amount'];
        } elseif (isset($item['nodes'][0]['inputQty']['amount'])) {
            // Fallback to inputQty if availToSellQty is not available
            $quantity = (int) $item['nodes'][0]['inputQty']['amount'];
        }
         if (!$sku) {
                Log::warning('Missing SKU in parsed Amazon data', $item);
                continue;
            }
        // Only process if we have a valid SKU
        if ($sku !== null) {
            ProductStockMapping::updateOrCreate(
                ['sku' => $sku],
                ['inventory_walmart' => $quantity]
            );
        }
    }

    return $collected;
}


}
