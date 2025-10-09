<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\ProductMaster;
use App\Models\Warehouse;
use Illuminate\Http\Request;
use App\Http\Controllers\ApiController;
use App\Models\AutoStockBalance;
use Carbon\Carbon;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

use App\Http\Controllers\ShopifyApiInventoryController;
use App\Models\ShopifySku;
use App\Models\Inventory;
use App\Models\ShopifyInventory;

use App\Models\AmazonDataView;
use App\Models\AmazonListingStatus;
use App\Services\AmazonSpApiService;
use App\Models\ProductStockMapping;

class StockMappingController extends Controller
{

    protected $shopifyDomain;
    protected $shopifyApiKey;
    protected $shopifyPassword;

    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
        $this->shopifyApiKey = config('services.shopify.api_key');
        $this->shopifyPassword = config('services.shopify.password');
        $this->shopifyStoreUrl = str_replace(['https://', 'http://'],'',config('services.shopify.store_url'));
        $this->shopifyStoreUrlName = env('SHOPIFY_STORE');
        $this->shopifyAccessToken = env('SHOPIFY_PASSWORD');
    }


    /**
     * Display a listing of the resource.
     */
    public function index()
    {
       return view('stock_mapping.view-stock-mapping');
    }

     
   public function getShopifyAmazonInventoryStock(Request $request)
{
    // Check if data is older than 1 day
    $latestRecord = ProductStockMapping::orderBy('updated_at', 'desc')->first();
    if ($latestRecord && $latestRecord->updated_at > now()->subDay()) {
        // Return cached data from DB
        $data = ProductStockMapping::all();
        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $data,
            'status' => 200
        ]);
    }
    
    $freshData=$this->fetchFreshData();   

    return response()->json([
        'message' => 'Data fetched successfully',
        'data' => $freshData,
        'status' => 200
    ]);
}

protected function fetchFreshData(){
    // Fetch fresh data from APIs
    $shopifyInventoryData = $this->getAllInventoryData();
    $amazonInventoryData = $this->getAllInventoryDataAmazon();

    // Index Amazon data by SKU
    $amazonIndex = [];
    foreach ($amazonInventoryData as $item) {
        if (!empty($item['sku'])) {
            $amazonIndex[$item['sku']] = $item;
        }
    }

    $mergedInventory = [];

    foreach ($shopifyInventoryData as $shopifyItem) {
        $sku = $shopifyItem['sku'] ?? null;
        if (!$sku) continue;

        $amazonItem = $amazonIndex[$sku] ?? null;
        $product_title = $shopifyItem['product_title'] ?? '';
        $inventoryShopify = $shopifyItem['inventory'] ?? 0;

        $inventoryAmazon = 'Not Listed';
        if ($amazonItem !== null && array_key_exists('quantity', $amazonItem)) {
            $qty = (int) $amazonItem['quantity'];
            $inventoryAmazon = ($qty === 0) ? 0 : $qty;
        }

        $mergedInventory[] = [
            'sku' => $sku,
            'product_title' => $product_title,
            'inventory_shopify' => $inventoryShopify,
            'inventory_amazon' => $inventoryAmazon,
        ];

        $insertData = [
            'sku' => $sku,
            'title' => $product_title,
            'inventory_shopify' => $inventoryShopify,
            'inventory_shopify_product' => json_encode($shopifyItem),
            'inventory_amazon' => $inventoryAmazon,
            'inventory_amazon_product' => json_encode($amazonItem),
        ];

        ProductStockMapping::updateOrCreate(
            ['sku' => $sku],
            $insertData
        );
    }
    return $mergedInventory;
}

 
    protected function getAllInventoryData(): array
    {
        $inventoryData = [];
        $pageInfo = null;
        $hasMore = true;
        $pageCount = 0;
        $totalProducts = 0;
        $totalVariants = 0;

        Log::info("Starting Shopify inventory fetch...");

        while ($hasMore) {
            $pageCount++;
            $queryParams = ['limit' => 250, 'fields' => 'id,title,variants,image,images'];
            if ($pageInfo) {$queryParams['page_info'] = $pageInfo;}
             $request = Http::withHeaders([
                'X-Shopify-Access-Token' => $this->shopifyAccessToken,
                'Content-Type' => 'application/json'
            ]);

            if (env('FILESYSTEM_DRIVER') === 'local') {
                $request = $request->withoutVerifying();
            }

            $response = $request
                ->timeout(120)
                ->retry(3, 500)
                ->get("https://{$this->shopifyStoreUrl}/admin/api/2025-01/products.json", $queryParams);


            if (!$response->successful()) {
                Log::error("Failed to fetch products (Page {$pageCount}): " . $response->body());
                break;
            }

            $products = $response->json()['products'] ?? [];
            $productCount = count($products);
            $totalProducts += $productCount;

            Log::info("Page {$pageCount} fetched successfully. Products: {$productCount}");

            foreach ($products as $product) {
                foreach ($product['variants'] as $variant) {
                    $totalVariants++;
                  
                    if (!empty($variant['sku'])) {
                        $inventoryData[$variant['sku']] = [
                            'variant_id'        => $variant['id'],
                            'inventory'         => $variant['inventory_quantity'] ?? 0,
                            'product_title'     => $product['title'] ?? '',
                            'sku'               => $variant['sku'] ?? '',
                            'variant_title'     => $variant['title'] ?? '',
                            'inventory_item_id' => $variant['inventory_item_id'],
                            'on_hand'           => $variant['old_inventory_quantity'] ?? 0,   // OnHand
                            'available_to_sell' => $variant['inventory_quantity'] ?? 0,       // AvailableToSell
                            'price'             => $variant['price'],                           
                        ];

                        // Log first 3 SKUs + images per page (to avoid huge logs)
                        if ($totalVariants <= 3 || $totalVariants % 500 === 0) {
                            Log::info("Variant preview", [
                                'product_title' => $product['title'] ?? '',
                                'sku'           => $variant['sku'],
                            ]);
                        }
                    } else {
                        Log::warning('Variant without SKU', [
                            'product_id' => $product['id'],
                            'variant_id' => $variant['id'],
                            'on_hand'    => $variant['old_inventory_quantity'] ?? 0,
                            'available_to_sell' => $variant['inventory_quantity'] ?? 0,
                        ]);
                    }
                }
            }

            // Pagination handling
            $pageInfo = $this->getNextPageInfo($response);
            $hasMore = (bool) $pageInfo;

            // Avoid rate limiting
            if ($hasMore) {
                Log::info(" Waiting 0.5s before next page...");
                usleep(500000); // 0.5s delay
            }
        }

        Log::info("Finished fetching Shopify inventory. Pages: {$pageCount}, Products: {$totalProducts}, Variants: {$totalVariants}");

        return $inventoryData;
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


    protected function getAllInventoryDataAmazon(){
        return $result = (new AmazonSpApiService())->getAmazonInventory();
    }
    
}

