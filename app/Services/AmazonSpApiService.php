<?php

namespace App\Services;

use GuzzleHttp\Client;
use GuzzleHttp\Psr7\Request;
use Aws\Signature\SignatureV4;
use Aws\Credentials\Credentials;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class AmazonSpApiService
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
    public function getAccessToken()
    {
        $client = new Client();
        $response = $client->post('https://api.amazon.com/auth/o2/token', [
            'form_params' => [
                'grant_type' => 'refresh_token',
                'refresh_token' => $this->refreshToken,
                'client_id' => $this->clientId,
                'client_secret' => $this->clientSecret,
            ]
        ]);

        $data = json_decode($response->getBody(), true);
        return $data['access_token'];
    }

    private function getAccessTokenV1()
    {
        $res = Http::withoutVerifying()->asForm()->post('https://api.amazon.com/auth/o2/token', [
            'grant_type' => 'refresh_token',
            'refresh_token' => env('SPAPI_REFRESH_TOKEN'),
            'client_id' => env('SPAPI_CLIENT_ID'),
            'client_secret' => env('SPAPI_CLIENT_SECRET'),
        ]);

        return $res['access_token'] ?? null;
    }

    public function updateAmazonPriceUS($sku, $price)
    {
        $sellerId = env('AMAZON_SELLER_ID');
        $accessToken = $this->getAccessToken();

        $productType = $this->getAmazonProductType($sku);

        $endpoint = "https://sellingpartnerapi-na.amazon.com/listings/2021-08-01/items/{$sellerId}/" . rawurlencode($sku) . "?marketplaceIds=ATVPDKIKX0DER";

        $body = [
            "productType" => $productType,
            "patches" => [[
                "op" => "replace",
                "path" => "/attributes/purchasable_offer",
                "value" => [[
                    "marketplaceId" => "ATVPDKIKX0DER",
                    "currency" => "USD",
                    "our_price" => [
                        [
                            "schedule" => [
                                [
                                    "value_with_tax" => (float) $price
                                ]
                            ]
                        ]
                    ]
                ]]
            ]]
        ];

        $response = Http::withToken($accessToken)
            ->withHeaders([
                'x-amz-access-token' => $accessToken,
                'content-type' => 'application/json',
                'accept' => 'application/json',
            ])
            ->patch($endpoint, $body);

        // Log::info("Amazon Price Update Request", [
        //     "sku" => $sku,
        //     "price" => $price,
        //     "endpoint" => $endpoint,
        //     "body" => $body
        // ]);

        if ($response->failed()) {
            Log::error("Amazon Price Update Failed", [
                "sku" => $sku,
                "status" => $response->status(),
                "response" => $response->json()
            ]);
        } else {
            Log::info("Amazon Price Update Success", $response->json());
        }

        return $response->json();
    }

    public function getAmazonProductType($sku)
    {
        $sellerId = env('AMAZON_SELLER_ID');
        $accessToken = $this->getAccessToken();

        $url = "https://sellingpartnerapi-na.amazon.com/listings/2021-08-01/items/{$sellerId}/" . rawurlencode($sku) . "?marketplaceIds=ATVPDKIKX0DER";

        $response = Http::withHeaders([
            'x-amz-access-token' => $accessToken,
            'Content-Type' => 'application/json',
        ])->get($url);

        if ($response->failed()) {
            return null;
        }

        $data = $response->json();
        return $data['summaries'][0]['productType'] ?? null;
    }


  public function getAmazonInventory()
{
    try {
        
        if (env('FILESYSTEM_DRIVER') === 'local') {$accessToken = $this->getAccessTokenV1();}
        else{$accessToken = $this->getAccessToken();}

        Log::info('Access Token: ', [$accessToken]);

        $marketplaceId = env('SPAPI_MARKETPLACE_ID');

         $request =Http::withHeaders([
            'x-amz-access-token' => $accessToken,
         ]);
         
         if (env('FILESYSTEM_DRIVER') === 'local') {$request = $request->withoutVerifying();}

        // Step 1: Request the report
        $response = $request->post('https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports', [
            'reportType' => 'GET_MERCHANT_LISTINGS_ALL_DATA',
            'marketplaceIds' => [$marketplaceId],
        ]);

        Log::info('Report Request Response: ' . $response->body());

        $reportId = $response['reportId'] ?? null;
        if (!$reportId) {
            Log::error('Failed to request report. Response: ' . $response->body());
            return response()->json(['error' => 'Failed to request report.'], 500);
        }

        // Step 2: Wait for report generation
        $processingStatus = 'IN_PROGRESS';
        while ($processingStatus !== 'DONE') {
            sleep(15);
            $request2=Http::withHeaders(['x-amz-access-token' => $accessToken,]);
            if (env('FILESYSTEM_DRIVER') === 'local') {$request2 = $request2->withoutVerifying();}

            $status = $request2->get("https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports/{$reportId}");

            $processingStatus = $status['processingStatus'] ?? 'UNKNOWN';
            Log::info("Waiting... Status: {$processingStatus}");
        }

        $documentId = $status['reportDocumentId'] ?? null;
        if (!$documentId) {
            Log::error('Document ID not found in report status.');
            return response()->json(['error' => 'Document ID not found.'], 500);
        }

        // Step 3: Get report document details
        $request3=Http::withHeaders(['x-amz-access-token' => $accessToken,]);
        if (env('FILESYSTEM_DRIVER') === 'local') {$request3 = $request3->withoutVerifying();}

        $doc = $request3->get("https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/documents/{$documentId}");

        $url = $doc['url'] ?? null;
        if (!$url) {
            Log::error('Document URL not found.');
            return response()->json(['error' => 'Document URL not found.'], 500);
        }

        // Step 4: Download and parse the data
        $csv = file_get_contents($url);
        $lines = explode("\n", $csv);
        $headers = explode("\t", array_shift($lines));

        $parsedData = [];
        foreach ($lines as $line) {
            $row = str_getcsv($line, "\t");
            if (count($row) < count($headers)) continue;

            $data = array_combine($headers, $row);
            if (($data['fulfillment-channel'] ?? '') !== 'DEFAULT') continue;

            $asin = $data['asin1'] ?? null;
            $title = $data['item-name'] ?? null;
            $sku = isset($data['seller-sku']) ? preg_replace('/[^\x20-\x7E]/', '', trim($data['seller-sku'])) : null;
            $price = isset($data['price']) && is_numeric($data['price']) ? $data['price'] : null;
            $quantity = isset($data['quantity']) && is_numeric($data['quantity']) ? $data['quantity'] : null;

            $parsedData[] = [
                'asin' => $asin,
                'product_title' => $title,
                'sku' => $sku,
                'price' => $price,
                'quantity'=>$quantity
            ];
        }

        Log::info('Amazon Inventory fetched successfully.', ['count' => count($parsedData)]);
        return $parsedData;
        // return response()->json(['success' => true, 'data' => $parsedData]);
    } catch (\Exception $e) {
        Log::error('Error in getAmazonInventory: ' . $e->getMessage(), [
            'trace' => $e->getTraceAsString(),
        ]);
        return response()->json(['error' => $e->getMessage()], 500);
    }
}

}
