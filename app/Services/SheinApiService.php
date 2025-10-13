<?php

namespace App\Services;

use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use SimpleXMLElement;
use ZipArchive;
use Illuminate\Support\Str;

class SheinApiService
{

      protected $appId;
    protected $appSecret;
    protected $baseUrl = 'https://openapi.sheincorp.com'; // or sandbox: openapi-test01.sheincorp.cn

    public function __construct()
    {
        $this->appId     = env('SHEIN_APP_ID');
        $this->appSecret = env('SHEIN_APP_SECRET');
    }

  function generateSheinSignature($path, $timestamp, $randomKey)
    {
        $openKeyId = env('SHEIN_OPEN_KEY_ID');
        $secretKey = env('SHEIN_SECRET_KEY');

        $value = $openKeyId . "&" . $timestamp . "&" . $path;

        $key = $secretKey . $randomKey;

        $hmacResult = hash_hmac('sha256', $value, $key, false); // false means return hexadecimal

        $base64Signature = base64_encode($hmacResult);

        $finalSignature = $randomKey . $base64Signature;

        return $finalSignature;
    }

     public function listAllProducts()
    {
        $endpoint  = "/open-api/openapi-business-backend/product/query";
        $pageSize  = 400;
        $allProducts = [];

        // Loop max 1000 pages (safe upper bound)
        for ($pageNum = 1; $pageNum <= 1000; $pageNum++) {

            $timestamp = round(microtime(true) * 1000);
            $random    = Str::random(5);
            $signature = $this->generateSheinSignature($endpoint, $timestamp, $random);

            $url = $this->baseUrl . $endpoint;

            $payload = [
                "pageNum"         => $pageNum,
                "pageSize"        => $pageSize,
                "insertTimeEnd"   => "",
                "insertTimeStart" => "",
                "updateTimeEnd"   => "",
                "updateTimeStart" => "",
            ];

            $response = Http::withoutVerifying()->withHeaders([
                "Language"       => "en-us",
                "x-lt-openKeyId" => env('SHEIN_OPEN_KEY_ID'),
                "x-lt-timestamp" => $timestamp,
                "x-lt-signature" => $signature,
                "Content-Type"   => "application/json",
            ])->post($url, $payload);

            if (!$response->successful()) {
                throw new \Exception("Shein API Error: " . $response->body());
            }

            $data = $response->json();
            $products = $data["info"]["data"] ?? [];
            // dd($products);

            // If no products returned → stop looping
            if (empty($products)) {
                break;
            }

            $allProducts = array_merge($allProducts, $products);
        }
        
        $spuNames = array_map(function ($item) {
    return $item['skuCodeList'][0] ?? null;
}, $allProducts);

$spuNames = array_filter($spuNames); // remove nulls if any

    $result=$this->getStock($spuNames);
    // dd($result);
    return $result;
        // return $spuNames;
    }


public function getStock(array $skuCodes)
{
    $endpoint = "/open-api/openapi-business-backend/product/full-detail";
    $chunkSize = 100;
    $allStock = [];

    // Split SKU codes into chunks of 100
    $chunks = array_chunk($skuCodes, $chunkSize);

    foreach ($chunks as $chunk) {
        $timestamp = round(microtime(true) * 1000);
        $random = Str::random(5);
        $signature = $this->generateSheinSignature($endpoint, $timestamp, $random);
        $url = $this->baseUrl . $endpoint;

        $payload = [
            "skuCodes" => $chunk
        ];

        $response = Http::withoutVerifying()->withHeaders([
            "Language" => "en-us",
            "x-lt-openKeyId" => env('SHEIN_OPEN_KEY_ID'),
            "x-lt-timestamp" => $timestamp,
            "x-lt-signature" => $signature,
            "Content-Type" => "application/json",
        ])->post($url, $payload);

        if (!$response->successful()) {
            throw new \Exception("Shein API Error: " . $response->body());
        }

        $data = $response->json();

        if (isset($data["info"]) && is_array($data["info"])) {
            foreach ($data["info"] as $item) {
                $skuCode = $item['sellerSku'] ?? null;
                $quantity = $item['goodsInventory']['inventoryQuantity'] ?? null;

                if ($skuCode !== null && $quantity !== null) {
                    $allStock[] = [
                        'sku' => $skuCode,
                        'quantity' => (int) $quantity,
                    ];
                }
            }
        }
    }

    return $allStock;
}

    public function getStock1(array $spus)
{
    $endpoint = "/open-api/stock/stock-query";
    $chunkSize = 10;
    $allStock = [];

    // Split SPUs into chunks of 100
    $chunks = array_chunk($spus, $chunkSize);

    foreach ($chunks as $chunk) {
        $timestamp = round(microtime(true) * 1000);
        $random = Str::random(5);
        $signature = $this->generateSheinSignature($endpoint, $timestamp, $random);
        $url = $this->baseUrl . $endpoint;

        $payload = [
            "languageList" => ["en"],
            "skuCodeList" => [],       // must be empty
            "skcNameList" => [],       // must be empty
            "spuNameList" => $chunk,   // only this populated
            "warehouseType" => "3",    // required: 1, 2, or 3
        ];

        $response = Http::withoutVerifying()->withHeaders([
            "Language" => "en-us",
            "x-lt-openKeyId" => env('SHEIN_OPEN_KEY_ID'),
            "x-lt-timestamp" => $timestamp,
            "x-lt-signature" => $signature,
            "Content-Type" => "application/json",
        ])->post($url, $payload);

        if (!$response->successful()) {
            throw new \Exception("Shein API Error: " . $response->body());
        }

        $data = $response->json();
        // dd($data['info']);
        if (isset($data["info"]["data"]) && is_array($data["info"]["data"])) {
            $allStock = array_merge($allStock, $data["info"]["data"]);
        }
    }

    return $allStock;
}

}
