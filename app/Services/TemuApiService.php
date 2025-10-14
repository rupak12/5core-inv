<?php

namespace App\Services;

use GuzzleHttp\Client;
use GuzzleHttp\Psr7\Request;
use Aws\Signature\SignatureV4;
use Aws\Credentials\Credentials;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use App\Models\ProductStockMapping;

class TemuApiService
{
    protected $clientId;
    protected $clientSecret;
    protected $refreshToken;
    protected $region;
    protected $marketplaceId;
    protected $awsAccessKey;
    protected $awsSecretKey;
    protected $endpoint;


 private function generateSignValue($requestBody)
    {
        // Environment/config variables
        $appKey = env('TEMU_APP_KEY');
        $appSecret = env('TEMU_SECRET_KEY');
        $accessToken = env('TEMU_ACCESS_TOKEN');
        $timestamp = time();
        
        // Top-level params
        $params = [
            'access_token' => $accessToken,
            'app_key' => $appKey,
            'timestamp' => $timestamp,
            'data_type' => 'JSON',
        ];

        // Flatten and sort for signing
        $signParams = array_merge($params, $requestBody);
        ksort($signParams);
        
        $temp = '';
        foreach ($signParams as $key => $value) {
            if (is_array($value) || is_object($value)) {
                $value = json_encode($value, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
            }
            $temp .= $key . $value;
        }

        $signStr = $appSecret . $temp . $appSecret;
        $sign = strtoupper(md5($signStr));
        $params['sign'] = $sign;

        
        // Log the request
        Log::info("Generated Sign: $sign");
        return array_merge($params, $requestBody);
    }

    public function getInventory(){
        $allitems=[];
        $pageNumber = 1;
        do {
                $requestBody = [
                    // "type" => "bg.local.goods.list.query",
                    "type" => "bg.local.goods.list.query",
                    "goodsSearchType"=>1,
                    "pageSize" => 100,
                    "pageNumber" => $pageNumber,

                ];

                   $signedRequest = $this->generateSignValue($requestBody);
                   $request= Http::withHeaders([
                    'Content-Type' => 'application/json'
                   ]);
                    if (env('FILESYSTEM_DRIVER') === 'local') {$request = $request->withoutVerifying();}

                    $response =$request->post('https://openapi-b-us.temu.com/openapi/router', $signedRequest);
    
                if ($response->failed()) {
                    $this->error("Request failed: " . $response->body());
                    break;
                }
    
                 $data = $response->json();
        if (!($data['success'] ?? false)) {
            $this->error("Temu Error: " . ($data['errorMsg'] ?? 'Unknown'));
            break;
        }

        $items = $data['result']['goodsList'] ?? [];
        if (empty($items)) break;

        foreach ($items as $item) {
            $skuId = $item['outGoodsSn'] ?? null;
            $qty = $item['quantity'] ?? 0;

            if (!$skuId) continue;

            $allitems[]=[
                'sku'=>$skuId,
                'quantity'=>$qty
            ];
           
        }

        $pageNumber++;
    } while (true);

    foreach ($allitems as $sku => $data) {
      $sku = $data['sku'] ?? null;
        $quantity = $data['quantity'];        
            ProductStockMapping::updateOrCreate(
                ['sku' => $sku],
                ['inventory_temu'=>$quantity,]
            );
        }
    return $allitems;
    } 

}
