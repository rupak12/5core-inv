<?php

namespace App\Services;

use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use SimpleXMLElement;
use ZipArchive;

class EbayApiService
{

    protected $appId;
    protected $certId;
    protected $devId;
    protected $userToken;
    protected $endpoint;
    protected $siteId;
    protected $compatLevel;

    public function __construct()
    {
        $this->appId       = env('EBAY_APP_ID');
        $this->certId      = env('EBAY_CERT_ID');
        $this->devId       = env('EBAY_DEV_ID');
        $this->endpoint    = env('EBAY_TRADING_API_ENDPOINT', 'https://api.ebay.com/ws/api.dll');
        $this->siteId      = env('EBAY_SITE_ID', 0); // US = 0
        $this->compatLevel = env('EBAY_COMPAT_LEVEL', '1189');
    }
    public function generateBearerToken()
    {
        // 1. If cached token exists, return it immediately
        if (Cache::has('ebay_bearer')) {
            echo "\nBearer Token in Cache";

            return Cache::get('ebay_bearer');
        }
       
        echo "Generating New Ebay Token";


        // 2. Otherwise, request new token from eBay
        $clientId     = env('EBAY_APP_ID');
        $clientSecret = env('EBAY_CERT_ID');
        $refreshToken = env('EBAY_REFRESH_TOKEN');

        $response = Http::withoutVerifying()->asForm()
            ->withBasicAuth($clientId, $clientSecret)
            ->post('https://api.ebay.com/identity/v1/oauth2/token', [
                'grant_type'    => 'refresh_token',
                'refresh_token' => $refreshToken,
                'scope'         => 'https://api.ebay.com/oauth/api_scope'.'https://api.ebay.com/oauth/api_scope/sell.inventory',
            ]);

        if ($response->failed()) {
            throw new \Exception('Failed to get eBay token: ' . $response->body());
        }

        $data        = $response->json();
        $accessToken = $data['access_token'];
        $expiresIn   = $data['expires_in'] ?? 3600; // seconds, defaults to 1h

        // 3. Store token in cache for slightly less than expiry time
        Cache::put('ebay_bearer', $accessToken, now()->addSeconds($expiresIn - 60));

        return $accessToken;
    }


    public function reviseFixedPriceItem($itemId, $price, $quantity = null, $sku = null, $variationSpecifics = null, $variationSpecificsSet = null)
    {
                // Build XML body
        $xml = new SimpleXMLElement('<?xml version="1.0" encoding="utf-8"?><ReviseFixedPriceItemRequest xmlns="urn:ebay:apis:eBLBaseComponents"/>');
        $credentials = $xml->addChild('RequesterCredentials');
        
        $authToken = $this->generateBearerToken();

        $credentials->addChild('eBayAuthToken', $authToken ?? '');


        $item = $xml->addChild('Item');
        $item->addChild('ItemID', $itemId);

        // Update price
        $item->addChild('StartPrice', $price);

        // Optionally update quantity
        if ($quantity !== null) {
            $item->addChild('Quantity', $quantity);
        }

        // If variation exists, use variation structure
        if ($variationSpecifics && $variationSpecificsSet) {
            $variations = $item->addChild('Variations');
            $variation = $variations->addChild('Variation');

            if ($sku) {
                $variation->addChild('SKU', $sku);
            }

            $variation->addChild('StartPrice', $price);
            if ($quantity !== null) {
                $variation->addChild('Quantity', $quantity);
            }

            // VariationSpecifics
            $vs = $variation->addChild('VariationSpecifics');
            foreach ($variationSpecifics as $name => $value) {
                $nvl = $vs->addChild('NameValueList');
                $nvl->addChild('Name', $name);
                $nvl->addChild('Value', $value);
            }

            // VariationSpecificsSet
            $vss = $item->addChild('VariationSpecificsSet');
            foreach ($variationSpecificsSet as $name => $values) {
                $nvl = $vss->addChild('NameValueList');
                $nvl->addChild('Name', $name);
                foreach ($values as $val) {
                    $nvl->addChild('Value', $val);
                }
            }
        }

        $xmlBody = $xml->asXML();

        // Prepare headers
        $headers = [
            'X-EBAY-API-COMPATIBILITY-LEVEL' => $this->compatLevel,
            'X-EBAY-API-DEV-NAME'            => $this->devId,
            'X-EBAY-API-APP-NAME'            => $this->appId,
            'X-EBAY-API-CERT-NAME'           => $this->certId,
            'X-EBAY-API-CALL-NAME'           => 'ReviseFixedPriceItem',
            'X-EBAY-API-SITEID'              => $this->siteId,
            'Content-Type'                   => 'text/xml',
        ];

        // Send API request
        $response = Http::withHeaders($headers)
            ->withBody($xmlBody, 'text/xml')
            ->post($this->endpoint);

        $body = $response->body();

        // Parse XML response
        libxml_use_internal_errors(true);
        $xmlResp = simplexml_load_string($body);
        
        if ($xmlResp === false) {
            return [
                'success' => false,
                'message' => 'Invalid XML response',
                'raw' => $body,
            ];
        }

        $responseArray = json_decode(json_encode($xmlResp), true);
        $ack = $responseArray['Ack'] ?? 'Failure';

        if ($ack === 'Success' || $ack === 'Warning') {
            return [
                'success' => true,
                'message' => 'Item updated successfully.',
                'data' => $responseArray,
            ];
        } else {
            return [
                'success' => false,
                'errors' => $responseArray['Errors'] ?? 'Unknown error',
                'data' => $responseArray,
            ];
        }
    }

    private function generateEbayToken(): ?string
    {
       
       $clientId = env('EBAY_APP_ID');
        $clientSecret = env('EBAY_CERT_ID');
        $refreshToken = env('EBAY_REFRESH_TOKEN');
        $credentials = base64_encode("{$clientId}:{$clientSecret}");

        $payload = [
            'grant_type' => 'refresh_token',
            'refresh_token' => $refreshToken,
            'scope' => implode(' ', [
                'https://api.ebay.com/oauth/api_scope/sell.inventory',
                'https://api.ebay.com/oauth/api_scope/sell.account',
            ]),
        ];

        $response = Http::withoutVerifying()
            ->asForm()
            ->withHeaders([
                'Authorization' => "Basic {$credentials}",
                'Content-Type' => 'application/json',
            ])
            ->post('https://api.ebay.com/identity/v1/oauth2/token', $payload);

        if ($response->failed()) {
            Log::error('eBay Access Token Error', ['response' => $response->json()]);
            throw new \RuntimeException('Unable to retrieve eBay access token.');
        }

        return $response->json('access_token');
    }
    
// ==========================================================================
 /**
     * Check API rate limits
     */
    public function getRateLimitForAPI(String $name, String $context)
    {
        $bearerToken = $this->generateEbayToken();

        $response = Http::withHeaders([
            'Authorization' => "Bearer {$bearerToken}"
        ])
            ->get('https://api.ebay.com/developer/analytics/v1_beta/rate_limit', [
                'api_name' => $name,
                'api_context' => $context,
            ]);

        return $response->json();
    }
    public function getEbayInventory(){
        $token = $this->generateEbayToken();
         if (!$token) {
            Log::error('Failed to generate token.');
            return;
        }
        $listingData = $this->fetchAndParseReport('LMS_ACTIVE_INVENTORY_REPORT', null, $token);
        return $listingData;
        $itemIdToSku = [];
        // foreach ($listingData as $row) {
        //     if (!empty($row['item_id']) && !empty($row['sku'])) {
        //         $itemIdToSku[$row['item_id']] = $row['sku'];
        //     }
        // }

    }

    public function fetchAndParseReport($reportType, $range, $token): array{
 
        Log::info("===============================================");
        Log::info("Start Processing: $reportType");
        $apiUrl = 'https://api.ebay.com/sell/feed/v1/inventory_task';
        $payload = [
            'feedType' => $reportType,
            'format' => 'TSV_GZIP',
            'schemaVersion' => '1.0',
        ];
        // Log::info('Request Payload:', [$payload]);
        $request= Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
            'Content-Type' => 'application/json',
        ]);
         if (env('FILESYSTEM_DRIVER') === 'local') {
        $request = $request->withoutVerifying();
        }
        $response =$request->post($apiUrl, $payload);
        Log::info('==============================');
        Log::info('Response:', [$response]);
        Log::info('==============================');
        $location = $response->header('Location');
        Log::info('location', [$location]);
        // dd($response);
        if (!$location) {
            Log::error("No 'Location' header returned. Can't extract task ID.");
            Log::error("Missing Location header", ['headers' => $response->headers()]);
            return [];
        }

        Log::info("=======// Step 2: Extract the task ID from URL");
        // Step 2: Extract the task ID from URL
        $taskId = basename($location); 
        Log::info("Task/Report ID: $taskId");
        $status = null;
        $downloadUrl = null;
        do {
            sleep(10);
            $request2=Http::withHeaders([
                'Authorization' => 'Bearer ' . $token,
            ]);
            if (env('FILESYSTEM_DRIVER') === 'local') {
                $request2 = $request2->withoutVerifying();
            }   
            $statusResponse = $request2->get("https://api.ebay.com/sell/feed/v1/inventory_task/{$taskId}");        
            $status = $statusResponse['status'] ?? 'PENDING';
            Log::info("Status: $status");
        
        } while (!in_array($status, ['COMPLETED', 'COMPLETED_WITH_ERROR', 'FAILED']));
        if ($status === 'FAILED') {
            Log::error("Inventory report task failed.");
            return [];
        }
        $data = $this->downloadAndParseEbayReport($taskId, $token);
        Log::info($data);
        return $data;
    }

    public function downloadAndParseEbayReport(string $taskId, string $token): array
    {
        Log::info('####');
        Log::info('downloadAndParseEbayReport');
        $baseTaskUrl = "https://api.ebay.com/sell/feed/v1/task/{$taskId}/download_result_file";
        $filePath = storage_path("app/inventory_{$taskId}");
        $zipPath = $filePath . ".zip";
        $xmlPath = $filePath . ".xml";

        Log::info("Downloading report from: $baseTaskUrl");
        try {
            $request3=Http::withHeaders([
                'Authorization' => 'Bearer ' . $token,
            ]);
            if (env('FILESYSTEM_DRIVER') === 'local') {
                $request3 = $request3->withoutVerifying();
            }   
            $response = $request3->get($baseTaskUrl);
            $content = $response->body();
            $magic = substr($content, 0, 2);
            if ($magic === "PK") {
                file_put_contents($zipPath, $content);
                $zip = new ZipArchive;
                if ($zip->open($zipPath) === TRUE) {
                    $zip->extractTo(storage_path('app/'));
                    $zip->close();
                    $extractedFiles = glob(storage_path('app/*.xml'));
                    if (empty($extractedFiles)) {
                        Log::error("No XML file found in zip.");
                        return [];
                    }
                    $xmlPath = $extractedFiles[0];
                    $xml = simplexml_load_file($xmlPath);

                    if (!$xml) {
                        Log::error("Failed to parse XML.");
                        return [];
                    }
                    Log::info("Root Element: " . $xml->getName());
                    Log::info("XML Preview", json_decode(json_encode($xml), true));

                    $data = [];
                    foreach ($xml->ActiveInventoryReport->SKUDetails as $item) {
                        $itemId = (string) $item->ItemID ?? null;
                        if (!$itemId) continue;
                        $data[] = [
                            'item_id' => $itemId,
                            'sku' => $item->SKU ?? '',
                            'price' => (float) ($item->Price ?? 0),
                        ];
                        // Handle variations if any
                        if (!empty($item->Variations->Variation)) {
                            foreach ($item->Variations->Variation as $variation) {
                                $itemId = (string) $item->ItemID ?? null;
                                $data[] = [
                                    'item_id' => $itemId,
                                    'sku' => $variation->SKU ?? '',
                                    'price' => (float) ($variation->Price ?? 0),
                                ];
                            }
                        }
                    }
                    @unlink($zipPath);
                    @unlink($xmlPath);
                    Log::info("Parsed " . count($data) . " XML items.");
                    Log::info('Sample parsed items:', array_slice($data, 0, 5));
                    return $data;
                }
                else {
                    logger()->error("Failed to open ZIP file.");
                    return [];
                }
            }
            // If not ZIP, check for GZ
            if (substr($content, 0, 2) === "\x1f\x8b") {
                $gzPath = $filePath . ".tsv.gz";
                $tsvPath = $filePath . ".tsv";
                file_put_contents($gzPath, $content);
                $gz = gzopen($gzPath, 'rb');
                $tsv = fopen($tsvPath, 'wb');
                while (!gzeof($gz)) {
                    fwrite($tsv, gzread($gz, 4096));
                }
                fclose($tsv);
                gzclose($gz);

                $lines = file($tsvPath, FILE_SKIP_EMPTY_LINES);
                if (!$lines || count($lines) < 2) return [];

                $rows = array_map('str_getcsv', $lines, array_fill(0, count($lines), "\t"));
                $headers = array_shift($rows);
                $data = [];

                foreach ($rows as $row) {
                    if (count($headers) !== count($row)) continue;
                    $item = array_combine($headers, $row);
                    $itemId = $item['itemId'] ?? null;
                    if (!$itemId) continue;

                    $data[$itemId] = [
                        'price' => $item['price'] ?? null,
                        'sku' => $item['sku'] ?? null,
                    ];
                }

                @unlink($gzPath);
                @unlink($tsvPath);
                Log::info("Parsed " . count($data) . " TSV items.");
                return $data;
            }

            // Unknown content
            Log::error("Unknown file type", [
                'first_bytes' => bin2hex(substr($content, 0, 4)),
                'taskId' => $taskId,
            ]);
            return [];

        } catch (\Throwable $e) {
            Log::error("Exception: " . $e->getMessage());
            return [];
        }
    }

    // =====================================================


    public function getEbayInventory3()
{
    $token = $this->generateEbayToken();
    if (!$token) {
        Log::error('Failed to generate eBay token.');
        return [];
    }

    // ✅ Correct feed type (NO "LMS_" prefix)
    $reportType = 'LMS_ACTIVE_INVENTORY_REPORT';

    Log::info("Start Processing: $reportType");

    // ✅ Fixed URL: no trailing spaces
    $apiUrl = 'https://api.ebay.com/sell/feed/v1/inventory_task';

    // ✅ Correct schema version (v3.0 as of 2024)
    $payload = [
        'feedType' => $reportType,
        'format' => 'TSV_GZIP', // You can also use 'XML' if preferred
        'schemaVersion' => '1.0'
    ];

    // Log::info('Request Payload:', [$payload]);

    $request = Http::withHeaders([
        'Authorization' => 'Bearer ' . $token,
        'Content-Type' => 'application/json',
    ]);

    if (env('FILESYSTEM_DRIVER') === 'local') {
        $request = $request->withoutVerifying();
    }

    $response = $request->post($apiUrl, $payload);
     if (!$response->successful()) {
        Log::error('Failed to create inventory task', [
            'status' => $response->status(),
            'body' => $response->body()
        ]);
        return [];
    }
    dd($response);

    if (!$response->successful()) {
        Log::error('Failed to create inventory task', [
            'status' => $response->status(),
            'body' => $response->body()
        ]);
        return [];
    }

    $location = $response->header('Location');
    Log::info('Location header:', [$location]);

    if (!$location) {
        Log::error("No 'Location' header returned. Can't extract task ID.");
        logger()->error("Missing Location header", ['headers' => $response->headers()]);
        return [];
    }

    // ✅ Extract task ID correctly
    $taskId = basename($location);
    Log::info("Task ID: $taskId");

    // Poll until task is complete
    $status = 'PENDING';
    $maxAttempts = 30; // ~5 minutes max
    $attempts = 0;

    do {
        sleep(10);
        $attempts++;

        $statusRequest = Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
        ]);

        if (env('FILESYSTEM_DRIVER') === 'local') {
            $statusRequest = $statusRequest->withoutVerifying();
        }

        // ✅ Fixed URL: no extra spaces
        $statusResponse = $statusRequest->get("https://api.ebay.com/sell/feed/v1/inventory_task/{$taskId}");

        if (!$statusResponse->successful()) {
            Log::warning('Failed to get task status', [
                'status' => $statusResponse->status(),
                'body' => $statusResponse->body()
            ]);
            continue;
        }

        $status = $statusResponse->json('status', 'PENDING');
        Log::info("Task status: $status (attempt $attempts)");

        if ($attempts >= $maxAttempts) {
            Log::error("Max polling attempts reached for task $taskId");
            return [];
        }

    } while (!in_array($status, ['COMPLETED', 'COMPLETED_WITH_ERROR', 'FAILED']));

    if ($status === 'FAILED') {
        Log::error("Inventory report task failed.", ['taskId' => $taskId]);
        return [];
    }

    Log::info('Downloading and parsing eBay report');

    // ✅ CORRECT download URL: must use /inventory_task/, not /task/
    $downloadUrl = "https://api.ebay.com/sell/feed/v1/inventory_task/{$taskId}/download_result_file";
    $filePath = storage_path("app/inventory_{$taskId}");

    try {
        $downloadRequest = Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
        ]);

        if (env('FILESYSTEM_DRIVER') === 'local') {
            $downloadRequest = $downloadRequest->withoutVerifying();
        }

        $downloadResponse = $downloadRequest->get($downloadUrl);

        if (!$downloadResponse->successful()) {
            Log::error('Failed to download report', [
                'status' => $downloadResponse->status(),
                'body' => $downloadResponse->body()
            ]);
            return [];
        }

        $content = $downloadResponse->body();
        $magic = substr($content, 0, 2);

        // Handle ZIP (XML) format
        if ($magic === "PK") {
            $zipPath = $filePath . ".zip";
            file_put_contents($zipPath, $content);

            $zip = new \ZipArchive;
            if ($zip->open($zipPath) === true) {
                $zip->extractTo(storage_path('app/'));
                $zip->close();

                $extractedFiles = glob(storage_path('app/*.xml'));
                if (empty($extractedFiles)) {
                    Log::error("No XML file found in ZIP.");
                    @unlink($zipPath);
                    return [];
                }

                $xmlPath = $extractedFiles[0];
                $xml = simplexml_load_file($xmlPath);
                if (!$xml) {
                    Log::error("Failed to parse XML.");
                    @unlink($zipPath);
                    @unlink($xmlPath);
                    return [];
                }

                // ✅ Handle XML namespace (critical for v3.0)
                $xml->registerXPathNamespace('ns', 'http://www.ebay.com/marketplace/sell/v1/services');
                $inventoryItems = $xml->xpath('//ns:ActiveInventory');

                $data = [];
                foreach ($inventoryItems as $item) {
                    $itemId = (string)($item->ItemID ?? '');
                    if (empty($itemId)) continue;

                    $data[] = [
                        'item_id' => $itemId,
                        'sku' => (string)($item->SKU ?? ''),
                        'price' => (float)($item->Price ?? 0),
                    ];
                }

                @unlink($zipPath);
                @unlink($xmlPath);
                Log::info("Parsed " . count($data) . " XML items.");
                return $data;
            } else {
                Log::error("Failed to open ZIP file.");
                @unlink($zipPath);
                return [];
            }
        }

        // Handle GZIP (TSV) format
        if (substr($content, 0, 2) === "\x1f\x8b") {
            $gzPath = $filePath . ".tsv.gz";
            $tsvPath = $filePath . ".tsv";
            file_put_contents($gzPath, $content);

            $gz = gzopen($gzPath, 'rb');
            $tsv = fopen($tsvPath, 'wb');
            while (!gzeof($gz)) {
                fwrite($tsv, gzread($gz, 4096));
            }
            fclose($tsv);
            gzclose($gz);

            $lines = @file($tsvPath, FILE_SKIP_EMPTY_LINES | FILE_IGNORE_NEW_LINES);
            if (!$lines || count($lines) < 2) {
                Log::error("TSV file is empty or invalid.");
                @unlink($gzPath);
                @unlink($tsvPath);
                return [];
            }

            $rows = array_map('str_getcsv', $lines, array_fill(0, count($lines), "\t"));
            $headers = array_shift($rows);
            $data = [];

            foreach ($rows as $row) {
                if (count($row) !== count($headers)) continue;
                $item = array_combine($headers, $row);
                $itemId = $item['item_id'] ?? null; // ✅ correct column name
                if (!$itemId) continue;

                $data[] = [
                    'item_id' => $itemId,
                    'sku' => $item['sku'] ?? '',
                    'price' => isset($item['price']) ? (float)$item['price'] : 0,
                ];
            }

            @unlink($gzPath);
            @unlink($tsvPath);
            Log::info("Parsed " . count($data) . " TSV items.");
            return $data;
        }

        // Unknown format
        Log::error("Unknown report file format", [
            'first_bytes_hex' => bin2hex(substr($content, 0, 8)),
            'taskId' => $taskId,
        ]);
        return [];

    } catch (\Throwable $e) {
        Log::error("Exception during report download/parsing: " . $e->getMessage(), [
            'trace' => $e->getTraceAsString()
        ]);
        return [];
    }
}

    public function getEbayInventory1(){
         $token = $this->generateEbayToken();
        if (!$token) { $this->error('Failed to generate token.'); return; }
        $reportType='LMS_ACTIVE_INVENTORY_REPORT';

        // $listingData = $this->fetchAndParseReport('LMS_ACTIVE_INVENTORY_REPORT', null, $token);
        Log::info("Start Processing: $reportType");

        $apiUrl = 'https://api.ebay.com/sell/feed/v1/inventory_task';

        $payload = ['feedType' => $reportType,'format' => 'TSV_GZIP','schemaVersion' => '1.0'];
         Log::info('Request Payload:', [$payload]);

        $request=Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
            'Content-Type' => 'application/json',
        ]);
         if (env('FILESYSTEM_DRIVER') === 'local') {$request = $request->withoutVerifying();}
        $response = $request->post($apiUrl, $payload);

        $location = $response->header('Location');
         Log::info('location', [$location]);

        if (!$location) {
            Log::error("No 'Location' header returned. Can't extract task ID.");
            logger()->error("Missing Location header", ['headers' => $response->headers()]);
            return [];
        }

        // Step 2: Extract the task ID from URL
        $taskId = basename($location); 
         Log::info("Task ID: $taskId");

         Log::info("Task/Report ID: $taskId");

        $status = null;
        $downloadUrl = null;


          do {
            sleep(10);
            $request2=Http::withHeaders([
                'Authorization' => 'Bearer ' . $token,
            ]);
            if (env('FILESYSTEM_DRIVER') === 'local') {$request2 = $request2->withoutVerifying();}
            $statusResponse = $request2->get("https://api.ebay.com/sell/feed/v1/inventory_task/{$taskId}");
        
            $status = $statusResponse['status'] ?? 'PENDING';
             Log::info("Status: $status");
        
        } while (!in_array($status, ['COMPLETED', 'COMPLETED_WITH_ERROR', 'FAILED']));
        
        if ($status === 'FAILED') {
             Log::error("Inventory report task failed.");
            return [];
        }


        info('downloadAndParseEbayReport');
        $baseTaskUrl = "https://api.ebay.com/sell/feed/v1/task/{$taskId}/download_result_file";
        $filePath = storage_path("app/inventory_{$taskId}");
        $zipPath = $filePath . ".zip";
        $xmlPath = $filePath . ".xml";

         Log::info("Downloading report from: $baseTaskUrl");

        try {
            $request3=Http::withHeaders(['Authorization' => 'Bearer ' . $token,]);

            if (env('FILESYSTEM_DRIVER') === 'local') {$request3 = $request3->withoutVerifying();}

            $response = $request3->get($baseTaskUrl);

            $content = $response->body();
            $magic = substr($content, 0, 2);

            // ZIP file: starts with "PK"
            if ($magic === "PK") {
                file_put_contents($zipPath, $content);

                $zip = new ZipArchive;
                if ($zip->open($zipPath) === TRUE) {
                    $zip->extractTo(storage_path('app/'));
                    $zip->close();

                    // Find extracted XML file
                    $extractedFiles = glob(storage_path('app/*.xml'));
                    if (empty($extractedFiles)) {
                        logger()->error("No XML file found in zip.");
                        return [];
                    }

                    $xmlPath = $extractedFiles[0];
                    $xml = simplexml_load_file($xmlPath);
                    if (!$xml) {
                        logger()->error("Failed to parse XML.");
                        return [];
                    }

                    logger()->info("Root Element: " . $xml->getName());
                    logger()->info("XML Preview", json_decode(json_encode($xml), true));

                    // Example conversion (customize based on XML structure)
                    $data = [];
                    foreach ($xml->ActiveInventoryReport->SKUDetails as $item) {
                        $itemId = (string) $item->ItemID ?? null;
                        if (!$itemId) continue;
                    
                        $data[] = [
                            'item_id' => $itemId,
                            'sku' => $item->SKU ?? '',
                            'price' => (float) ($item->Price ?? 0),
                        ];
                    
                        // Handle variations if any
                        if (!empty($item->Variations->Variation)) {
                            foreach ($item->Variations->Variation as $variation) {
                                $itemId = (string) $item->ItemID ?? null;
                                $data[] = [
                                    'item_id' => $itemId,
                                    'sku' => $variation->SKU ?? '',
                                    'price' => (float) ($variation->Price ?? 0),
                                ];
                            }
                        }
                    }

                    @unlink($zipPath);
                    @unlink($xmlPath);
                    
                     Log::info("Parsed " . count($data) . " XML items.");
                    logger()->info('Sample parsed items:', array_slice($data, 0, 5));
                    
                    return $data;
                } else {
                     Log::error("Failed to open ZIP file.");
                    return [];
                }
            }

            // If not ZIP, check for GZ
            if (substr($content, 0, 2) === "\x1f\x8b") {
                $gzPath = $filePath . ".tsv.gz";
                $tsvPath = $filePath . ".tsv";
                file_put_contents($gzPath, $content);

                $gz = gzopen($gzPath, 'rb');
                $tsv = fopen($tsvPath, 'wb');
                while (!gzeof($gz)) {
                    fwrite($tsv, gzread($gz, 4096));
                }
                fclose($tsv);
                gzclose($gz);

                $lines = file($tsvPath, FILE_SKIP_EMPTY_LINES);
                if (!$lines || count($lines) < 2) return [];

                $rows = array_map('str_getcsv', $lines, array_fill(0, count($lines), "\t"));
                $headers = array_shift($rows);
                $data = [];

                foreach ($rows as $row) {
                    if (count($headers) !== count($row)) continue;
                    $item = array_combine($headers, $row);
                    $itemId = $item['itemId'] ?? null;
                    if (!$itemId) continue;

                    $data[$itemId] = [
                        'price' => $item['price'] ?? null,
                        'sku' => $item['sku'] ?? null,
                    ];
                }

                @unlink($gzPath);
                @unlink($tsvPath);
                 Log::info("Parsed " . count($data) . " TSV items.");
                return $data;
            }

            // Unknown content
             Log::error("Unknown file type", [
                'first_bytes' => bin2hex(substr($content, 0, 4)),
                'taskId' => $taskId,
            ]);
            return [];

        } catch (\Throwable $e) {
             Log::error("Exception: " . $e->getMessage());
            return [];
        }

    }
}
