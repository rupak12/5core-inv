<?php

namespace App\Services;

use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use SimpleXMLElement;
use ZipArchive;
use App\Models\ProductStockMapping;

class Ebay3ApiService
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
        $this->appId       = env('EBAY_3_APP_ID');
        $this->certId      = env('EBAY_3_CERT_ID');
        $this->devId       = env('EBAY_DEV_ID');
        $this->endpoint    = env('EBAY_TRADING_API_ENDPOINT', 'https://api.ebay.com/ws/api.dll');
        $this->siteId      = env('EBAY_SITE_ID', 0); // US = 0
        $this->compatLevel = env('EBAY_COMPAT_LEVEL', '1189');
    }
   
    private function generateEbayToken(): ?string
    {
       
       $clientId = env('EBAY_3_APP_ID');
        $clientSecret = env('EBAY_3_CERT_ID');
        $refreshToken = env('EBAY_3_REFRESH_TOKEN');
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
        foreach ($listingData as $sku => $data) {
        $sku = $data['sku'] ?? null;
        $quantity = $data['quantity'];
        
            ProductStockMapping::updateOrCreate(
                ['sku' => $sku],
                ['inventory_ebay1'=>$quantity,]
            );
        }
        return $listingData;
        $itemIdToSku = [];       
    }

    public function fetchAndParseReport($reportType, $range, $token): array
{
    Log::info("Start Processing: $reportType");
    
    $apiUrl = 'https://api.ebay.com/sell/feed/v1/inventory_task';
    $payload = [
        'feedType' => $reportType,
        'format' => 'TSV_GZIP',
        'schemaVersion' => '1.0',
    ];

    try {
        // Create HTTP client with common settings
        $request = Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
            'Content-Type' => 'application/json',
        ])->timeout(60); // Add timeout

        // Disable SSL verification if needed (consider security implications)
        if (env('APP_ENV') === 'local' || env('APP_DEBUG') === true) {
            $request = $request->withoutVerifying();
        }

        Log::info('Sending request to eBay API');
        $response = $request->post($apiUrl, $payload);
        
        // Check if request was successful
        if (!$response->successful()) {
            Log::error("API request failed", [
                'status' => $response->status(),
                'body' => $response->body(),
                'headers' => $response->headers()
            ]);
            return [];
        }

        $location = $response->header('Location');
        Log::info('Location header', ['location' => $location]);

        if (!$location) {
            Log::error("No 'Location' header returned");
            Log::error("Response headers", ['headers' => $response->headers()]);
            return [];
        }

        // Extract task ID from URL
        $taskId = basename($location); 
        Log::info("Task ID: $taskId");

        $status = null;
        $maxAttempts = 30; // 5 minutes max waiting (30 * 10 seconds)
        $attempts = 0;

        do {
            sleep(10);
            $attempts++;
            
            Log::info("Checking task status (attempt $attempts)");
            
            $statusRequest = Http::withHeaders([
                'Authorization' => 'Bearer ' . $token,
            ])->timeout(30);

            if (env('APP_ENV') === 'local' || env('APP_DEBUG') === true) {
                $statusRequest = $statusRequest->withoutVerifying();
            }

            $statusResponse = $statusRequest->get("https://api.ebay.com/sell/feed/v1/inventory_task/{$taskId}");
            
            if (!$statusResponse->successful()) {
                Log::error("Status check failed", [
                    'status' => $statusResponse->status(),
                    'body' => $statusResponse->body()
                ]);
                continue; // Continue waiting despite temporary failures
            }

            $responseData = $statusResponse->json();
            $status = $responseData['status'] ?? 'PENDING';
            Log::info("Task Status: $status");

            // Break if max attempts reached to prevent infinite loop
            if ($attempts >= $maxAttempts) {
                Log::error("Max attempts reached. Task did not complete in time.");
                return [];
            }
        
        } while (!in_array($status, ['COMPLETED', 'COMPLETED_WITH_ERROR', 'FAILED']));

        if ($status === 'FAILED') {
            Log::error("Inventory report task failed for task ID: $taskId");
            return [];
        }

        Log::info("Task completed with status: $status");
        $data = $this->downloadAndParseEbayReport($taskId, $token);
        
        return $data;

    } catch (\Exception $e) {
        Log::error("Exception in fetchAndParseReport: " . $e->getMessage());
        return [];
    }
}

public function downloadAndParseEbayReport(string $taskId, string $token): array
{  $data = [];
    Log::info("Downloading report for task: $taskId");
    
    $baseTaskUrl = "https://api.ebay.com/sell/feed/v1/task/{$taskId}/download_result_file";
    $filePath = storage_path("app/inventory_{$taskId}");
    $zipPath = $filePath . ".zip";
    $xmlPath = $filePath . ".xml";

    try {
        $request = Http::withHeaders([
            'Authorization' => 'Bearer ' . $token,
        ])->timeout(120); // Longer timeout for file download

        if (env('APP_ENV') === 'local' || env('APP_DEBUG') === true) {
            $request = $request->withoutVerifying();
        }

        Log::info("Downloading report from: $baseTaskUrl");
        $response = $request->get($baseTaskUrl);
        
        if (!$response->successful()) {
            Log::error("Download failed", [
                'status' => $response->status(),
                'body' => $response->body()
            ]);
            return [];
        }

        $content = $response->body();
        
        if (empty($content)) {
            Log::error("Empty response content");
            return [];
        }

        $magic = substr($content, 0, 2);
        Log::info("File type detection - Magic bytes: " . bin2hex($magic));

        // ZIP file: starts with "PK"
        if ($magic === "PK") {
            Log::info("Processing ZIP file");
            file_put_contents($zipPath, $content);

            $zip = new ZipArchive;
            if ($zip->open($zipPath) === TRUE) {
                $zip->extractTo(storage_path('app/'));
                $zip->close();

                // Find extracted XML file
                $extractedFiles = glob(storage_path('app/*.xml'));
                if (empty($extractedFiles)) {
                    Log::error("No XML file found in zip.");
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

                Log::info("Root Element: " . $xml->getName());
                Log::info("XML structure preview", json_decode(json_encode($xml), true));

              
                // Handle different XML structures
                if (isset($xml->ActiveInventoryReport->SKUDetails)) {
                    foreach ($xml->ActiveInventoryReport->SKUDetails as $item) {
                        $itemId = (string) ($item->ItemID ?? null);
                        if (!$itemId) continue;
                        
                        $data[] = [                            
                            'sku' => (string) ($item->SKU ?? ''),
                            'quantity' => (string) ($item->Quantity ?? ''),                            
                        ];

                        // Handle variations if any
                        if (!empty($item->Variations->Variation)) {
                            foreach ($item->Variations->Variation as $variation) {
                                $variationItemId = (string) ($variation->ItemID ?? $itemId);
                                if (!$variationItemId) continue;
                                
                                $data[] = [                                    
                                    'sku' => (string) ($variation->SKU ?? ''),
                                    'quantity' => (float) ($variation->Quantity ?? 0),
                                ];
                            }
                        }
                    }
                } else {
                    Log::warning("Unexpected XML structure. Trying alternative parsing.");
                    // Alternative parsing for different XML structures
                    foreach ($xml->children() as $child) {
                        if ($child->getName() === 'item' || isset($child->ItemID)) {
                            $itemId = (string) ($child->ItemID ?? null);
                            if (!$itemId) continue;
                            
                            $data[] = [
                                'item_id' => $itemId,
                                'sku' => (string) ($child->SKU ?? ''),
                                'price' => (float) ($child->Price ?? 0),
                            ];
                        }
                    }
                }

                @unlink($zipPath);
                @unlink($xmlPath);
                
                Log::info("Successfully parsed " . count($data) . " items from XML");
                Log::info('Sample parsed items:', array_slice($data, 0, 3));
                return $data;
            } else {
                Log::error("Failed to open ZIP file.");
                @unlink($zipPath);
                return [];
            }
        }

        // If not ZIP, check for GZ (GZIP compressed TSV)
        if (substr($content, 0, 2) === "\x1f\x8b") {
            Log::info("Processing GZIP compressed TSV file");
            $gzPath = $filePath . ".tsv.gz";
            $tsvPath = $filePath . ".tsv";
            
            file_put_contents($gzPath, $content);

            $gz = gzopen($gzPath, 'rb');
            if (!$gz) {
                Log::error("Failed to open GZ file");
                @unlink($gzPath);
                return [];
            }

            $tsv = fopen($tsvPath, 'wb');
            if (!$tsv) {
                Log::error("Failed to create TSV file");
                gzclose($gz);
                @unlink($gzPath);
                return [];
            }

            while (!gzeof($gz)) {
                fwrite($tsv, gzread($gz, 4096));
            }
            fclose($tsv);
            gzclose($gz);

            $lines = file($tsvPath, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
            if (!$lines || count($lines) < 2) {
                Log::error("No data found in TSV file");
                @unlink($gzPath);
                @unlink($tsvPath);
                return [];
            }

            $rows = array_map(function($line) {
                return str_getcsv($line, "\t");
            }, $lines);
            
            $headers = array_shift($rows);
            $data = [];

            Log::info("TSV Headers: " . implode(', ', $headers));

            foreach ($rows as $index => $row) {
                if (count($headers) !== count($row)) {
                    Log::warning("Skipping row $index - column count mismatch");
                    continue;
                }
                
                try {
                    $item = array_combine($headers, $row);
                    // $itemId = $item['itemId'] ?? $item['item_id'] ?? null;
                    
                    if (!$itemId) {
                        Log::warning("Skipping row $index - no item ID found");
                        continue;
                    }

                    $data[] = [
                        'sku' => $item['sku'] ?? $item['SKU'] ?? null,
                        'quantity' => isset($item['Quantity']) ? (float) $item['Quantity'] : null,
                    ];
                } catch (\Exception $e) {
                    Log::warning("Error processing row $index: " . $e->getMessage());
                    continue;
                }
            }

            @unlink($gzPath);
            @unlink($tsvPath);
            
            Log::info("Successfully parsed " . count($data) . " items from TSV");
            Log::info('Sample parsed items:', array_slice($data, 0, 3));
            return $data;
        }

        // Unknown content type
        Log::error("Unknown file type", [
            'first_bytes' => bin2hex(substr($content, 0, 4)),
            'taskId' => $taskId,
            'content_length' => strlen($content)
        ]);
        
        // Log first 200 chars for debugging
        Log::debug("Content preview: " . substr($content, 0, 200));
        return [];

    } catch (\Throwable $e) {
        Log::error("Exception in downloadAndParseEbayReport: " . $e->getMessage());
        Log::error("Stack trace: " . $e->getTraceAsString());
        
        // Clean up any temporary files
        $tempFiles = [
            $zipPath ?? null,
            $xmlPath ?? null,
            $gzPath ?? null,
            $tsvPath ?? null
        ];
        
        foreach ($tempFiles as $tempFile) {
            if ($tempFile && file_exists($tempFile)) {
                @unlink($tempFile);
            }
        }
        
        return [];
    }
}
   
}
