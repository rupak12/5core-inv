<?php

namespace App\Services;

use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use SimpleXMLElement;

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
    // public function generateBearerToken()
    // {
    //     // 1. If cached token exists, return it immediately
    //     if (Cache::has('ebay_bearer')) {
    //         echo "\nBearer Token in Cache";

    //         return Cache::get('ebay_bearer');
    //     }
       
    //     echo "Generating New Ebay Token";


    //     // 2. Otherwise, request new token from eBay
    //     $clientId     = env('EBAY_APP_ID');
    //     $clientSecret = env('EBAY_CERT_ID');
    //     $refreshToken = env('EBAY_REFRESH_TOKEN');

    //     $response = Http::asForm()
    //         ->withBasicAuth($clientId, $clientSecret)
    //         ->post('https://api.ebay.com/identity/v1/oauth2/token', [
    //             'grant_type'    => 'refresh_token',
    //             'refresh_token' => $refreshToken,
    //             'scope'         => 'https://api.ebay.com/oauth/api_scope/sell.analytics.readonly https://api.ebay.com/oauth/api_scope/sell.inventory',
    //         ]);

    //     if ($response->failed()) {
    //         throw new \Exception('Failed to get eBay token: ' . $response->body());
    //     }

    //     $data        = $response->json();
    //     $accessToken = $data['access_token'];
    //     $expiresIn   = $data['expires_in'] ?? 3600; // seconds, defaults to 1h

    //     // 3. Store token in cache for slightly less than expiry time
    //     Cache::put('ebay_bearer', $accessToken, now()->addSeconds($expiresIn - 60));

    //     return $accessToken;
    // }
    public function generateBearerToken()
    {

        $clientId     = env('EBAY_APP_ID');
        $clientSecret = env('EBAY_CERT_ID');
        $refreshToken = env('EBAY_REFRESH_TOKEN');

        $response = Http::asForm()
            ->withBasicAuth($clientId, $clientSecret)
            ->post('https://api.ebay.com/identity/v1/oauth2/token', [
                'grant_type'    => 'refresh_token',
                'refresh_token' => $refreshToken,
                'scope'         => 'https://api.ebay.com/oauth/api_scope/sell.analytics.readonly https://api.ebay.com/oauth/api_scope/sell.inventory',
            ]);

        if ($response->failed()) {
            throw new \Exception('Failed to get eBay token: ' . $response->body());
        }

        $data        = $response->json();
        $accessToken = $data['access_token'] ?? null;
        $expiresIn   = $data['expires_in'] ?? 3600; 

        if (!$accessToken) {
            throw new \Exception('No access token returned from eBay.');
        }

        
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
    public function getValidTrackingRate()
{
    $accessToken = $this->generateBearerToken();
    $url = "https://api.ebay.com/sell/analytics/v1/seller_standards_profile";

    $response = Http::withToken($accessToken)
        ->withHeaders([
            'Content-Type' => 'application/json',
        ])
        ->get($url);

    if ($response->failed()) {
        return [
            'success' => false,
            'message' => 'Failed to fetch seller standards: ' . $response->body(),
        ];
    }

    $data = $response->json();

    // Get the first profile
    $profile = $data['standardsProfiles'][1] ?? null;

    if (!$profile || empty($profile['metrics'])) {
        return [
            'success' => false,
            'message' => 'Standards profile or metrics not found',
            'data' => $data,
        ];
    }
    $vtrMetric = null;
    foreach ($profile['metrics'] as $metric) {
        if (($metric['metricKey'] ?? null) === 'VALID_TRACKING_UPLOADED_WITHIN_HANDLING_RATE') {
            $vtrMetric = $metric;
            break;
        }
    }

    if (!$vtrMetric) {
        return [
            'success' => false,
            'message' => 'Valid Tracking Rate metric not found',
            'data' => $data,
        ];
    }

    return [
        'success' => true,
        'Channels' => 'Ebay1',
        'vtr' => $vtrMetric['value']['value'] ?? null,
        'numerator' => $vtrMetric['value']['numerator'] ?? null,
        'denominator' => $vtrMetric['value']['denominator'] ?? null,
        'thresholdLower' => $vtrMetric['thresholdLowerBound']['value'] ?? null,
    ];
 }
}
