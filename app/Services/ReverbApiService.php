<?php

namespace App\Services;

use GuzzleHttp\Client;
use GuzzleHttp\Psr7\Request;
use Aws\Signature\SignatureV4;
use Aws\Credentials\Credentials;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use App\Models\ProductStockMapping;

class ReverbApiService
{
    protected $clientId;
    protected $clientSecret;
    protected $refreshToken;
    protected $region;
    protected $marketplaceId;
    protected $awsAccessKey;
    protected $awsSecretKey;
    protected $endpoint;

    public function __construct()
    {
        $this->clientId = env('SPAPI_CLIENT_ID');
        $this->clientSecret = env('SPAPI_CLIENT_SECRET');
        $this->refreshToken = env('SPAPI_REFRESH_TOKEN');
        $this->region = env('SPAPI_REGION', 'us-east-1');
        $this->marketplaceId = env('SPAPI_MARKETPLACE_ID');
        $this->awsAccessKey = env('AWS_ACCESS_KEY_ID');
        $this->awsSecretKey = env('AWS_SECRET_ACCESS_KEY');
        $this->endpoint = 'https://sellingpartnerapi-na.amazon.com';
    }
    

 public function getInventory()
{
    $inventory = [];
    $url = 'https://api.reverb.com/api/my/listings'; // Start URL

    try {
        while ($url) {
            $response = Http::withoutVerifying()->withHeaders([
                'Authorization' => 'Bearer ' . config('services.reverb.token'),
                'Accept' => 'application/json',
                'Accept-Version' => '3.0',
            ])->get($url);

            if ($response->failed()) {
                Log::error('Failed to fetch inventory page.', [
                    'url' => $url,
                    'status' => $response->status(),
                    'body' => $response->body()
                ]);
                return []; // or break; depending on whether you want partial data
            }

            $data = $response->json();

            // Process listings
            if (isset($data['listings']) && is_array($data['listings'])) {
                foreach ($data['listings'] as $item) {
                    if (isset($item['sku'], $item['inventory'])) {
                        $inventory[] = [
                            'sku' => $item['sku'],
                            'quantity' => $item['inventory'],
                        ];
                    }
                }
            }

            // Check for next page
            if (isset($data['_links']['next']['href'])) {
                $url = $data['_links']['next']['href'];
                // Clean URL: Reverb sometimes adds trailing spaces in href
                $url = trim($url);
            } else {
                $url = null; // No more pages
            }
        }
       
        foreach ($inventory as $sku => $data) {
            $sku = $data['sku'] ?? null;
                    $quantity = $data['quantity'];
             if (!$sku) {
                Log::warning('Missing SKU in parsed Amazon data', $item);
                continue;
            }

            ProductStockMapping::updateOrCreate(
                ['sku' => $sku],
                ['inventory_reverb'=>$quantity,]
            );
        }
        return $inventory;

    } catch (\Throwable $e) {
        Log::error('Exception during paginated inventory fetch: ' . $e->getMessage(), [
            'trace' => $e->getTraceAsString()
        ]);
        return [];
    }
}

}
