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
use App\Models\ProductStockMapping;


use App\Services\ShopifyApiService;
use App\Services\AmazonSpApiService;
use App\Services\EbayApiService;
use App\Services\WalmartService;
use App\Services\ReverbApiService;
use App\Services\TemuApiService;
use App\Services\SheinApiService;
use App\Services\DobaApiService;
use App\Services\WayfairApiService;
use App\Services\MacysApiService;

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
  ini_set('max_execution_time', 300);
     
    // Check if data is older than 1 day
    $latestRecord = ProductStockMapping::orderBy('updated_at', 'desc')->first();
    if ($latestRecord) {
    // if ($latestRecord && $latestRecord->updated_at > now()->subDay()) {
        // Return cached data from DB
        $data = ProductStockMapping::all();
        $datainfo=$this->getDataInfo($data);
        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $data,
                'datainfo'=>$datainfo,
            'status' => 200
        ]);
    }
    
    $freshData=$this->fetchFreshData();   
    $datainfo=$this->getDataInfo($freshData);

    return response()->json([
        'message' => 'Data fetched successfully',
        'data' => $freshData,
        'datainfo'=>$datainfo,
        'status' => 200
    ]);
}

protected function fetchFreshData(){
    ini_set('max_execution_time', 600);
     
    //  $result = (new WayfairApiService())->getInventory();    
    //  dd($result);
    
    // $macyInventory = (new MacysApiService())->getInventory();
    // die();

    // Fetch fresh data from APIs
    $delete=ProductStockMapping::truncate();
    $shopifyInventoryData = (new ShopifyApiService())->getinventory();    
    die();
    $parentskuList=$this->filterParentSKU($shopifyInventoryData);
    $amazonInventoryData = $this->getAllInventoryDataAmazon();
    // $ebayInventoryData = $this->getAllInventoryDataebay();
    $walmartInventory=$this->WalmartInventoryData();
    $reverbInventory=$this->getReverbInventoryData();
    $sheinInventory = (new SheinApiService())->listAllProducts();
    $dobaInventory = (new DobaApiService())->getinventory();
    $temuInventory = (new TemuApiService())->getInventory();
    $macyInventory = (new MacysApiService())->getInventory();
    $ebay1Inventory = (new EbayApiService())->getEbayInventory();
    $ebay2Inventory = (new Ebay2ApiService())->getEbayInventory();
    $ebay3Inventory = (new Ebay3ApiService())->getEbayInventory();
    // Index Amazon data by SKU
    $amazonIndex = [];
    foreach ($amazonInventoryData as $item) {
        if (!empty($item['sku'])) {
            $amazonIndex[$item['sku']] = $item;
        }
    }

    $walmartIndex=[];
    $walmartInventory=$walmartInventory['inventories'][0];
    foreach($walmartInventory as $witem){
        if (!empty($witem['sku'])) {
            $walmartIndex[$witem['sku']] = $witem;
        }
    }

    $reverbIndex=[];
    foreach($reverbInventory as $revitem){
        if (!empty($revitem['sku'])) {
            $reverbIndex[$revitem['sku']] = $revitem;
        }
    }

    // sheinInventory
    $sheinIndex=[];
    foreach($sheinInventory as $sheinitem){
        if (!empty($sheinitem['sku'])) {
            $sheinIndex[$sheinitem['sku']] = $sheinitem;
        }
    }

    $dobaIndex=[];
    foreach($dobaInventory as $dobaitem){
        if (!empty($dobaitem['sku'])) {
            $dobaIndex[$dobaitem['sku']] = $dobaitem;
        }
    }


    $temuIndex=[];
    foreach($temuInventory as $temuitem){
        if (!empty($temuitem['sku'])) {
            $temuIndex[$temuitem['sku']] = $temuitem;
        }
    }

    $macyIndex=[];
    foreach($macyInventory as $macyitem){
        if (!empty($macyitem['sku'])) {
            $macyIndex[$macyitem['sku']] = $macyitem;
        }
    }

    $ebay1Index=[];
    foreach($ebay1Inventory as $ebay1item){
        if (!empty($ebay1item['sku'])) {
            $ebay1Index[$ebay1item['sku']] = $ebay1item;
        }
    }

    $ebay2Index=[];
    foreach($ebay2Inventory as $ebay2item){
        if (!empty($ebay2item['sku'])) {
            $ebay2Index[$ebay2item['sku']] = $ebay2item;
        }
    }

    $ebay3Index=[];
    foreach($ebay3Inventory as $ebay3item){
        if (!empty($ebay3item['sku'])) {
            $ebay3Index[$ebay3item['sku']] = $ebay3item;
        }
    }

    $mergedInventory = [];

    foreach ($shopifyInventoryData as $shopifyItem) {
        $sku = $shopifyItem['sku'] ?? null;
        $image=$shopifyItem['image_src'];
        if (!$sku) continue;

        $amazonItem = $amazonIndex[$sku] ?? null;
        $walmartItem = $walmartIndex[$sku] ?? null;
        $reverbItem = $reverbIndex[$sku] ?? null;
        $sheinItem = $sheinIndex[$sku] ?? null;
        $dobaItem = $dobaIndex[$sku] ?? null;
        $temuItem = $temuIndex[$sku] ?? null;
        $macyItem = $macyIndex[$sku] ?? null;
        $ebay1Item = $ebay1Index[$sku] ?? null;
        $ebay2Item = $ebay2Index[$sku] ?? null;
        $ebay3Item = $ebay3Index[$sku] ?? null;
        $product_title = $shopifyItem['product_title'] ?? '';
        $variant_id= $shopifyItem['variant_id'] ?? '';
        $parent=$parentskuList[$sku]??'';
        $is_parent=$parent!=''?1:0;
        $inventoryShopify = 'Not Listed';
        $product_id=$shopifyItem['product_id'];

if (!empty($shopifyItem) && array_key_exists('inventory', $shopifyItem)) {
    $qty = (int) $shopifyItem['inventory'];
    $inventoryShopify = $qty;
}

$inventoryAmazon = 'Not Listed';
if (!empty($amazonItem) && array_key_exists('quantity', $amazonItem)) {
    $qty = (int) $amazonItem['quantity'];
    $inventoryAmazon = $qty;
}

$inventoryWalmart = 'Not Listed';
if (!empty($walmartitem) && array_key_exists('quantity', $walmartitem)) {
    $qty = (int) $walmartitem['nodes'][0]['quantity']['availToSellQty']['amount'];
    $inventoryWalmart = $qty;
}

$inventoryReverb = 'Not Listed';
if (!empty($reverbItem) && array_key_exists('quantity', $reverbItem)) {
    $qty = (int) $reverbItem['quantity'];
    $inventoryReverb = $qty;
}

$inventoryShein = 'Not Listed';
if (!empty($sheinItem) && array_key_exists('quantity', $sheinItem)) {
    $qty = (int) $sheinItem['quantity'];
    $inventoryShein = $qty;
}

$inventoryDoba = 'Not Listed';
if (!empty($dobaItem) && array_key_exists('quantity', $dobaItem)) {
    $qty = (int) $dobaItem['quantity'];
    $inventoryDoba = $qty;
}

$inventoryTemu = 'Not Listed';
if (!empty($temuItem) && array_key_exists('quantity', $temuItem)) {
    $qty = (int) $temuItem['quantity'];
    $inventoryTemu = $qty;
}

$inventoryMacy = 'Not Listed';
if (!empty($macyItem) && array_key_exists('quantity', $macyItem)) {
    $qty = (int) $macyItem['quantity'];
    $inventoryMacy = $qty;
}

$inventoryEbay1 = 'Not Listed';
if (!empty($ebay1Item) && array_key_exists('quantity', $ebay1Item)) {
    $qty = (int) $ebay1Item['quantity'];
    $inventoryEbay1 = $qty;
}

$inventoryEbay2 = 'Not Listed';
if (!empty($ebay2Item) && array_key_exists('quantity', $ebay2Item)) {
    $qty = (int) $ebay2Item['quantity'];
    $inventoryEbay2 = $qty;
}

$inventoryEbay3 = 'Not Listed';
if (!empty($ebay3Item) && array_key_exists('quantity', $ebay3Item)) {
    $qty = (int) $ebay3Item['quantity'];
    $inventoryEbay3 = $qty;
}

    
        
        $mergedInventory[] = [
            'is_parent'=>$is_parent,
            'image'=>$image,
            'variant_id'=>$variant_id,
            'product_id'=>$product_id,
            'sku' => $sku,
            'product_title' => $product_title,
            'inventory_shopify' => $inventoryShopify,
            'inventory_amazon' => $inventoryAmazon,
            'inventory_walmart' => $inventoryWalmart,
            'inventory_reverb' => $inventoryReverb,
            'inventory_shein' => $inventoryShein,
            'inventory_doba' => $inventoryDoba,
            'inventory_temu' => $inventoryTemu,
            'inventory_macy' => $inventoryMacy,
            'inventory_ebay1' => $inventoryEbay1,
            'inventory_ebay2' => $inventoryEbay2,
            'inventory_ebay3' => $inventoryEbay3,
            'parent'=>$parent
        ];

        $insertData = [
            'is_parent'=>$is_parent,
            'variant_id'=>$variant_id,
            'product_id'=>$product_id,
            'sku' => $sku,
            'image'=>$image,
            'title' => $product_title,
            'inventory_shopify' => $inventoryShopify,
            'inventory_shopify_product' => json_encode($shopifyItem),
            'inventory_amazon' => $inventoryAmazon,
            'inventory_amazon_product' => json_encode($amazonItem),
            'inventory_walmart' => $inventoryWalmart,
            'inventory_reverb' => $inventoryReverb,
            'inventory_shein' => $inventoryShein,
            'inventory_doba' => $inventoryDoba,
            'inventory_temu' => $inventoryTemu,
            'inventory_macy' => $inventoryMacy,
            'inventory_ebay1' => $inventoryEbay1,
            'inventory_ebay2' => $inventoryEbay2,
            'inventory_ebay3' => $inventoryEbay3,
            'parent'=>$parent,
        ];

        ProductStockMapping::updateOrCreate(
            ['sku' => $sku],
            $insertData
        );
    }
    // dd($mergedInventory);
    return $mergedInventory;
}


protected function getDataInfo($data){
    $info = [
        'shopify' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'amazon' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
         'walmart' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'reverb' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'shein' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'doba' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'temu' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],

        'macy' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'ebay1' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'ebay2' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
        'ebay3' => [
            'listed' => 0,
            'notlisted' => 0,
            'matching' => 0,
            'mismatching' => 0,
        ],
    ];

    foreach ($data as $item) {
        $shopifyQty = $item['inventory_shopify'] ?? null;
        $amazonQty = $item['inventory_amazon'] ?? null;
        $walmartQty = $item['inventory_walmart'] ?? null;
        $reverbQty = $item['inventory_reverb'] ?? null;
        $sheinQty = $item['inventory_shein'] ?? null;
        $dobaQty = $item['inventory_doba'] ?? null;
        $temuQty = $item['inventory_temu'] ?? null;
        $macyQty = $item['inventory_macy'] ?? null;
        $ebay1Qty = $item['inventory_ebay1'] ?? null;
        $ebay2Qty = $item['inventory_ebay2'] ?? null;
        $ebay3Qty = $item['inventory_ebay3'] ?? null;

        $isShopifyListed = is_numeric($shopifyQty);
        $isAmazonListed = is_numeric($amazonQty);
        $isWalmartListed = is_numeric($walmartQty);
        $isReverbListed = is_numeric($reverbQty);
        $isSheinListed = is_numeric($sheinQty);
        $isDobaListed = is_numeric($dobaQty);
        $isTemuListed = is_numeric($temuQty);
        $isMacyListed = is_numeric($macyQty);
        $isEbay1Listed = is_numeric($ebay1Qty);
        $isEbay2Listed = is_numeric($ebay2Qty);
        $isEbay3Listed = is_numeric($ebay3Qty);

        // Channel-specific listing status
        $info['shopify'][$isShopifyListed ? 'listed' : 'notlisted']++;
        $info['amazon'][$isAmazonListed ? 'listed' : 'notlisted']++;
        $info['walmart'][$isWalmartListed ? 'listed' : 'notlisted']++;
        $info['reverb'][$isReverbListed ? 'listed' : 'notlisted']++;
        $info['shein'][$isSheinListed ? 'listed' : 'notlisted']++;
        $info['doba'][$isDobaListed ? 'listed' : 'notlisted']++;
        $info['temu'][$isTemuListed ? 'listed' : 'notlisted']++;
        $info['macy'][$isMacyListed ? 'listed' : 'notlisted']++;
        $info['ebay1'][$isEbay1Listed ? 'listed' : 'notlisted']++;
        $info['ebay2'][$isEbay2Listed ? 'listed' : 'notlisted']++;
        $info['ebay3'][$isEbay3Listed ? 'listed' : 'notlisted']++;
        

        // Channel-specific matching/mismatching
        if ($isShopifyListed && $isAmazonListed) {
            if ((int)$shopifyQty === (int)$amazonQty) {
                $info['shopify']['matching']++;
                $info['amazon']['matching']++;
                $info['walmart']['matching']++;
                $info['reverb']['matching']++;
                $info['shein']['matching']++;
                $info['doba']['matching']++;
                $info['temu']['matching']++;
                $info['macy']['matching']++;
                $info['ebay1']['matching']++;
                $info['ebay2']['matching']++;
                $info['ebay3']['matching']++;
            } else {
                $info['shopify']['mismatching']++;
                $info['amazon']['mismatching']++;
                $info['walmart']['mismatching']++;
                $info['reverb']['mismatching']++;
                $info['shein']['mismatching']++;
                $info['doba']['mismatching']++;
                $info['temu']['mismatching']++;
                $info['macy']['mismatching']++;
                $info['ebay1']['mismatching']++;
                $info['ebay2']['mismatching']++;
                $info['ebay3']['mismatching']++;
            }
        }
    }

    return $info;
}

    protected function getAllInventoryData(): array
{
    $inventoryData = [];
    $parentVariants = [];
    $pageInfo = null;
    $hasMore = true;
    $pageCount = 0;
    $totalProducts = 0;
    $totalVariants = 0;
    

    Log::info("Starting Shopify inventory fetch...");

    while ($hasMore) {
        $pageCount++;
        $queryParams = ['limit' => 250, 'fields' => 'id,title,variants,image,images'];
        if ($pageInfo) {
            $queryParams['page_info'] = $pageInfo;
        }

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


            // dd($product);
            foreach ($product['variants'] as $variant) {
                $totalVariants++;

                $sku = $variant['sku'] ?? '';
                // $isParent = stripos($sku, 'PARENT') !== false;
                 $isParent = count($product['variants']) > 1?1:0;
                $imageUrl = $this->sanitizeImageUrl(
                    $product['image']['src'] ?? (!empty($product['images']) ? $product['images'][0]['src'] : null),
                    $sku
                );

                if (!empty($sku)) {
                    $inventoryData[$sku] = [
                        'variant_id'        => $variant['id'],
                        'product_id'        => $product['id'],
                        'inventory'         => $variant['inventory_quantity'] ?? 0,
                        'product_title'     => $product['title'] ?? '',
                        'sku'               => $sku,
                        'variant_title'     => $variant['title'] ?? '',
                        'inventory_item_id' => $variant['inventory_item_id'],
                        'on_hand'           => $variant['old_inventory_quantity'] ?? 0,
                        'available_to_sell' => $variant['inventory_quantity'] ?? 0,
                        'price'             => $variant['price'],
                        'image_src'         => $imageUrl,
                        'is_parent'         => $isParent,
                        
                    ];

                    if ($isParent) {
                        $parentVariants[] = [
                            'sku' => $sku,
                            'variant_id' => $variant['id'],
                            'product_title' => $product['title'] ?? '',
                        ];
                        Log::info("Parent SKU detected", end($parentVariants));
                    }

                    if ($totalVariants <= 3 || $totalVariants % 500 === 0) {
                        Log::info("Variant preview", [
                            'product_title' => $product['title'] ?? '',
                            'sku'           => $sku,
                            'image'         => $imageUrl,
                        ]);
                    }
                } else {
                    Log::warning('Variant without SKU', [
                        'product_id' => $product['id'],
                        'variant_id' => $variant['id'],
                        'on_hand'    => $variant['old_inventory_quantity'] ?? 0,
                        'available_to_sell' => $variant['inventory_quantity'] ?? 0,
                        'image'      => $imageUrl,
                    ]);
                }
            }
        }

        $pageInfo = $this->getNextPageInfo($response);
        $hasMore = (bool) $pageInfo;

        if ($hasMore) {
            Log::info("Waiting 0.5s before next page...");
            usleep(500000);
        }
    }

    Log::info("Finished fetching Shopify inventory. Pages: {$pageCount}, Products: {$totalProducts}, Variants: {$totalVariants}");


    return $inventoryData;
}

protected function filterParentSKU(array $data): array
{
    // Extract SKUs from input array
    $filteredSkus = array_values(array_filter(array_map(function ($item) {
        return $item['sku'] ?? null;
    }, $data)));

    // Query ProductMaster for matching SKUs
    $parentRecords = ProductMaster::whereIn('sku', $filteredSkus)->get();

    // Return associative array: [sku => parent]
    return $parentRecords->pluck('parent', 'sku')->toArray();
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

    protected function getAllInventoryDataebay(){
        return $result = (new EbayApiService())->getEbayInventory();
    }
    
    public function WalmartInventoryData(){
        return $result = (new WalmartService())->getAllInventoryData();
    }

    public function getReverbInventoryData(){
        return $result = (new ReverbApiService())->getInventory();
    }

    

      protected function sanitizeImageUrl(?string $url,$sku): ?string
    {

                     
        if (empty($url)) {
            return null;
        }
      
        // Remove line breaks and spaces
        $cleanUrl = trim(preg_replace('/\s+/', '', $url));

        // Remove ?v= query string (Shopify versioning param)
        $cleanUrl = strtok($cleanUrl, '?');

        return $cleanUrl;
    }

    protected function updateNotRequired(Request $request)
    {
       $not_required = $request->input('notrequired');
           foreach ($not_required as $entry) {
        [$sku, $id] = explode('___', $entry);

        ProductStockMapping::where('sku', $sku)
            ->where('id', $id)
            ->update(['not_required' => 1]); // or true, or any value you need
    }
        return response()->json(['status' => 'success']);
    }

    public function refetchLiveData(){
        $freshData=$this->fetchFreshData();   
        if($freshData){
            return response()->json(['status' => 'success']);
        }
    }

    

    }

    
