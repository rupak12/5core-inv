<?php

namespace App\Services;

use App\Models\ProductStockMapping;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class ShopifyApiService
{
    public function __construct()
    {
        $this->shopifyApiKey = config('services.shopify.api_key');
        $this->shopifyPassword = config('services.shopify.password');
        $this->shopifyStoreUrl = str_replace(['https://', 'http://'], '', config('services.shopify.store_url'));
        $this->shopifyStoreUrlName = env('SHOPIFY_STORE');
        $this->shopifyAccessToken = env('SHOPIFY_PASSWORD');
    }

    public function getInventory()
    {
        $inventoryData = [];
        $parentVariants = [];
        $pageInfo = null;
        $hasMore = true;
        $pageCount = 0;
        $totalProducts = 0;
        $totalVariants = 0;

        while ($hasMore) {
            $pageCount++;
            $queryParams = [
                'limit' => 250,
                'fields' => 'id,title,variants,image,images',
            ];
            if ($pageInfo) {
                $queryParams['page_info'] = $pageInfo;
            }

            $request = Http::withHeaders([
                'X-Shopify-Access-Token' => $this->shopifyAccessToken,
                'Content-Type' => 'application/json',
            ]);

            if (env('FILESYSTEM_DRIVER') === 'local') {
                $request = $request->withoutVerifying();
            }

            $response = $request->timeout(120)->retry(3, 500)->get(
                "https://{$this->shopifyStoreUrl}/admin/api/2025-01/products.json",
                $queryParams
            );

            if (! $response->successful()) {
                Log::error("Failed to fetch products (Page {$pageCount}): ".$response->body());
                break;
            }

            $products = $response->json()['products'] ?? [];
            $productCount = count($products);
            $totalProducts += $productCount;

            foreach ($products as $product) {
                foreach ($product['variants'] as $variant) {
                    $totalVariants++;

                    $sku = $variant['sku'] ?? '';
                    $isParent = count($product['variants']) > 1 ? 1 : 0;
                    $imageUrl = $this->sanitizeImageUrl(
                        $product['image']['src'] ?? (! empty($product['images']) ? $product['images'][0]['src'] : null),
                        $sku
                    );

                    if (! empty($sku)) {
                        $inventoryData[$sku] = [
                            'variant_id' => $variant['id'],
                            'product_id' => $product['id'],
                            'inventory' => $variant['inventory_quantity'] ?? 0,
                            'product_title' => $product['title'] ?? '',
                            'sku' => $sku,
                            'variant_title' => $variant['title'] ?? '',
                            'inventory_item_id' => $variant['inventory_item_id'],
                            'on_hand' => $variant['old_inventory_quantity'] ?? 0,
                            'available_to_sell' => $variant['inventory_quantity'] ?? 0,
                            'price' => $variant['price'],
                            'image_src' => $imageUrl,
                            'is_parent' => $isParent,
                        ];

                        if ($isParent) {
                            $parentVariants[] = [
                                'sku' => $sku,
                                'variant_id' => $variant['id'],
                                'product_title' => $product['title'] ?? '',
                            ];
                            Log::info('Parent SKU detected', end($parentVariants));
                        }

                        if ($totalVariants <= 3 || $totalVariants % 500 === 0) {
                            Log::info('Variant preview', [
                                'product_title' => $product['title'] ?? '',
                                'sku' => $sku,
                                'image' => $imageUrl,
                            ]);
                        }
                    } else {
                        Log::warning('Variant without SKU', [
                            'product_id' => $product['id'],
                            'variant_id' => $variant['id'],
                            'on_hand' => $variant['old_inventory_quantity'] ?? 0,
                            'available_to_sell' => $variant['inventory_quantity'] ?? 0,
                            'image' => $imageUrl,
                        ]);
                    }
                }
            }

            $pageInfo = $this->getNextPageInfo($response);
            $hasMore = (bool) $pageInfo;

            if ($hasMore) {
                Log::info('Waiting 0.5s before next page...');
                usleep(500000);
            }
        }

        // ✅ Update all SKUs in DB
        foreach ($inventoryData as $sku => $data) {
            ProductStockMapping::updateOrCreate(
                ['sku' => $sku],
                [
                    'image' =>$data['image_src'],
                    'inventory_shopify' => $data['inventory'],
                    'inventory_amazon'=>'Not Listed',
                    'inventory_walmart'=>'Not Listed',
                    'inventory_reverb'=>'Not Listed',
                    'inventory_shein'=>'Not Listed',
                    'inventory_doba'=>'Not Listed',
                    'inventory_temu'=>'Not Listed',
                    'inventory_macy'=>'Not Listed',
                    'inventory_ebay1'=>'Not Listed',
                    'inventory_ebay2'=>'Not Listed',
                    'inventory_ebay3'=>'Not Listed',
                ]
            );
        }

        return $inventoryData;
    }

    protected function sanitizeImageUrl(?string $url,$sku): ?string
    {
                     
        if (empty($url)) {return null;}
      
        // Remove line breaks and spaces
        $cleanUrl = trim(preg_replace('/\s+/', '', $url));

        // Remove ?v= query string (Shopify versioning param)
        $cleanUrl = strtok($cleanUrl, '?');

        return $cleanUrl;
    }
      
    protected function getNextPageInfo($response): ?string
    {
        if ($response->hasHeader('Link') && str_contains($response->header('Link'), 'rel="next"')) {
            $links = explode(',', $response->header('Link'));
            foreach ($links as $link) {
                if (str_contains($link, 'rel="next"')) {
                    preg_match('/<(.*)>; rel="next"/', $link, $matches);
                    parse_str(parse_url($matches[1], PHP_URL_QUERY), $query);
                    return $query['page_info'] ?? null;
                }
            }
        }
        return null;
    }
}
