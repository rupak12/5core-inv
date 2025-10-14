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
use App\Services\WalmartApiService;
use App\Services\ReverbApiService;
use App\Services\TemuApiService;
use App\Services\SheinApiService;
use App\Services\DobaApiService;
use App\Services\WayfairApiService;
use App\Services\MacysApiService;
use App\Services\BestbuyusaApiService;
class StockMappingController extends Controller
{

    protected $shopifyDomain;
    protected $shopifyApiKey;
    protected $shopifyPassword;


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
    ini_set('max_execution_time', 1000);
     
    //   $result = (new BestbuyusaApiService())->getChannels();
    //   dd($result);
    //  $result = (new WayfairApiService())->getInventory();    
    
    // $macyInventory = (new MacysApiService())->getInventory();
    // die();

    // Fetch fresh data from APIs
    $delete=ProductStockMapping::truncate();
    $shopifyInventoryData = (new ShopifyApiService())->getinventory();        
    $parentskuList=$this->filterParentSKU($shopifyInventoryData);
    $amazonInventoryData = (new AmazonSpApiService())->getinventory();
    $walmartInventory=(new WalmartApiService())->getinventory();
    $reverbInventory=(new ReverbApiService())->getInventory();
    $sheinInventory = (new SheinApiService())->listAllProducts();
    $dobaInventory = (new DobaApiService())->getinventory();
    $temuInventory = (new TemuApiService())->getInventory();
    $macyInventory = (new MacysApiService())->getInventory();
    $ebay1Inventory = (new EbayApiService())->getEbayInventory();
    $ebay2Inventory = (new Ebay2ApiService())->getEbayInventory();
    $ebay3Inventory = (new Ebay3ApiService())->getEbayInventory();
      $data = ProductStockMapping::all();
    return $data;
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


    

    protected function getAllInventoryDataebay(){
        return $result = (new EbayApiService())->getEbayInventory();
    }
    
    public function WalmartInventoryData(){
        return $result = (new WalmartService())->getAllInventoryData();
    }

    public function getReverbInventoryData(){
        return $result = (new ReverbApiService())->getInventory();
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

    
