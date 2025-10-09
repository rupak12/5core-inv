<?php

namespace App\Http\Controllers\InventoryManagement;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Http\Controllers\ApiController;
use App\Http\Controllers\ShopifyApiInventoryController;
use App\Models\ShopifySku;
use App\Models\ProductMaster;
use App\Models\Warehouse;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use App\Models\Inventory;
use Illuminate\Support\Facades\Auth;
use Carbon\Carbon;
use App\Models\ShopifyInventory;
use Illuminate\Support\Facades\DB;


class VerificationAdjustmentController extends Controller
{

    protected $shopifyDomain;
    protected $shopifyApiKey;
    protected $shopifyPassword;

    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
        $this->shopifyDomain = env('SHOPIFY_STORE_URL');
        $this->shopifyApiKey = env('SHOPIFY_API_KEY');
        $this->shopifyPassword = env('SHOPIFY_PASSWORD');
    }

    /**
     * Display a listing of the resource.
     */
    public function index()
    {
        return view('inventory-management.verification-adjustment');
    }

    /**
     * Show the form for creating a new resource.
     */
    public function create()
    {
        //
    }

    /**
     * Store a newly created resource in storage.
     */
    public function store(Request $request)
    {
        //
    }

    /**
     * Display the specified resource.
     */
    public function show(string $id)
    {
        //
    }

    /**
     * Show the form for editing the specified resource.
     */
    public function edit(string $id)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     */
    public function update(Request $request, string $id)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     */
    public function destroy(string $id)
    {
        //
    }

    // public function getViewVerificationAdjustmentData(Request $request)
    // {

    //     $shopifyInventoryController = new ShopifyApiInventoryController();
    //     $inventoryResponse = $shopifyInventoryController->fetchInventoryWithCommitment();
    //     Log::info('Fetched inventory response:', $inventoryResponse);
    //     // $inventoryArray = json_decode($inventoryResponse->getContent(), true)['data'] ?? [];

    //     // $inventoryData = collect($inventoryResponse)->mapWithKeys(function ($item) {
    //     //     return [strtoupper(trim($item['sku'])) => $item];
    //     // });
    //     $inventoryData = collect($inventoryResponse);
    //     // $inventoryData = json_decode($inventoryResponse->getContent(), true)['data'];
    //     // $inventoryData = $shopifyInventoryController->getInventoryArray();
    //     // $inventoryData = []; 
    //     Log::info('Shopify Inventory Data:', $inventoryData->toArray());


    //     // Fetch data from the Google Sheet using the ApiController method
    //     $response = $this->apiController->fetchDataFromProductMasterGoogleSheet();
        
    //     // Check if the response is successful
    //     if ($response->getStatusCode() === 200) { 
    //         $data = $response->getData(); // Get the JSON data from the response

    //         // Get all non-PARENT SKUs from the data to fetch from ShopifySku model
    //         // $skus = collect($data->data)
    //         //     ->filter(function ($item) {
    //         //         $childSku = $item->{'SKU'} ?? '';
    //         //         return !empty($childSku) && stripos($childSku, 'PARENT') === false;
    //         //     })
    //         //     ->map(function ($item) {
    //         //     return strtoupper(trim($item->{'SKU'})); //  Normalize SKUs
    //         //     })
    //         //     ->unique()->toArray();
    //         $skus = collect($data->data)
    //             ->map(function ($item) {
    //                 return strtoupper(trim($item->{'SKU'} ?? ''));
    //             })
    //             ->filter()
    //             ->unique()
    //             ->toArray();

    //         // Fetch Shopify inventory data for non-PARENT SKUs
    //         $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(function ($item) {
    //             return strtoupper(trim($item->sku)); //  Normalize DB SKUs
    //         });

    //         // Fetch saved Verified Stock data from DB
    //         //  $verifiedStockData = Inventory::get()->keyBy(function ($item) {
    //         //     return strtoupper(trim($item->sku)); // Normalize DB SKUs
    //         // });
    //         $verifiedStockData = Inventory::whereIn('sku', $skus)
    //         ->get()
    //         ->mapWithKeys(function ($item) {
    //             return [strtoupper(trim($item->sku)) => $item];
    //         });

    //         // Filter out rows where both Parent and (Child) sku are empty and process data
    //         $filteredData = array_filter($data->data, function ($item) {
    //             $parent = $item->Parent ?? '';
    //             $childSku = $item->{'SKU'} ?? '';

    //             // Keep the row if either Parent or (Child) sku is not empty
    //             return !(empty(trim($parent)) && empty(trim($childSku)));
    //         });

    //         // Process the data to include Shopify inventory values
    //         $mergedData = collect($filteredData)->map(function ($item) use ($shopifyData, $inventoryData, $verifiedStockData) {
    //             $childSku = $item->{'SKU'} ?? '';
    //             $normalizedSku = strtoupper(trim($childSku));

    //             $lp = isset($item->LP) && is_numeric($item->LP) ? floatval($item->LP) : 0;

    //             // Only update INV and L30 if this is not a PARENT SKU
    //             if (!empty($childSku) && stripos($childSku, 'PARENT') === false) {
    //                 if ($shopifyData->has($normalizedSku)) {
    //                     $item->INV = $shopifyData[$normalizedSku]->inv;
    //                     $item->L30 = $shopifyData[$normalizedSku]->quantity;
    //                 } else {
    //                     // Default to 0 if SKU not found in Shopify
    //                     $item->INV = 0;
    //                     $item->L30 = 0; 
    //                 }

    //                 if (isset($inventoryData[$normalizedSku])) {
    //                     $item->ON_HAND = $inventoryData[$normalizedSku]['on_hand'];
    //                     $item->COMMITTED = $inventoryData[$normalizedSku]['committed'];
    //                     $item->AVAILABLE_TO_SELL = $inventoryData[$normalizedSku]['available_to_sell'];

    //                     ShopifyInventory::updateOrCreate(
    //                         ['sku' => $normalizedSku],
    //                         [
    //                             'parent' => $item->Parent ?? null,
    //                             'on_hand' => $inventoryData[$normalizedSku]['on_hand'],
    //                             'committed' => $inventoryData[$normalizedSku]['committed'],
    //                             'available_to_sell' => $inventoryData[$normalizedSku]['available_to_sell'],
    //                         ]
    //                     );
                       
    //                 } else {
    //                     $item->ON_HAND = 'N/A';
    //                     $item->AVAILABLE_TO_SELL = 'N/A';
    //                     $item->COMMITTED  = 'N/A';
    //                 }

    //                 if ($verifiedStockData->has($normalizedSku)) {
    //                     $verifiedStockRow = $verifiedStockData[$normalizedSku];
    //                     $item->VERIFIED_STOCK = $verifiedStockRow->verified_stock ?? null;
    //                     $item->TO_ADJUST = $verifiedStockRow->to_adjust ?? null;
    //                     $item->REASON = $verifiedStockRow->reason ?? null;
    //                     $item->APPROVED = (bool) $verifiedStockRow->approved;
    //                     $item->APPROVED_BY = $verifiedStockRow->approved_by ?? null;

    //                 } else {
    //                     $item->VERIFIED_STOCK = null;
    //                     $item->TO_ADJUST = null;
    //                     $item->REASON = null;
    //                     $item->APPROVED = false;
    //                     $item->APPROVED_BY = null;
    //                 }

    //                 $adjustedQty = isset($item->TO_ADJUST) && is_numeric($item->TO_ADJUST) ? floatval($item->TO_ADJUST) : 0;
    //                 $item->LOSS_GAIN = round($adjustedQty * $lp, 2);
    //             }

    //             // For PARENT SKUs or when childSku is empty, keep original values

    //             return $item;
    //         });

    //         // Re-index the array after filtering
    //         // $processedData = array_values($processedData);
    //         $processedData = $mergedData->values();
    //         Log::info('Processed data count: ' . count($processedData));

    //         // Return the filtered data
    //         return response()->json([
    //             'message' => 'Data fetched successfully',
    //             'data' => $processedData,
    //             'status' => 200
    //         ]);
    //     } else {
    //         // Handle the error if the request failed
    //         return response()->json([
    //             'message' => 'Failed to fetch data from Google Sheet',
    //             'status' => $response->getStatusCode()
    //         ], $response->getStatusCode());
    //     }
    // }

    // public function getViewVerificationAdjustmentData(Request $request)       //13/09 working
    // {
    //     $shopifyInventoryController = new ShopifyApiInventoryController();
    //     $inventoryResponse = $shopifyInventoryController->fetchInventoryWithCommitment();
    //     Log::info('Fetched inventory response:', $inventoryResponse);

    //     $inventoryData = collect($inventoryResponse); 
    //     Log::info('Shopify Inventory Data:', $inventoryData->toArray());

    //     $response = $this->apiController->fetchDataFromProductMasterGoogleSheet();

    //     if ($response->getStatusCode() === 200) {
    //         $data = $response->getData();

    //         // use trimmed SKUs only (no strtoupper)
    //         $skus = collect($data->data)
    //             ->map(function ($item) {
    //                 return trim($item->{'SKU'} ?? '');
    //             })
    //             ->filter()
    //             ->unique()
    //             ->toArray();

    //         // use exact-case SKU keys
    //         $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(function ($item) {
    //             return trim($item->sku); 
    //         });

    //         $latestInventoryIds = Inventory::select(DB::raw('MAX(id) as latest_id'))
    //             ->whereIn('sku', $skus)
    //             ->groupBy('sku')
    //             ->pluck('latest_id');

    //         $latestInventoryData = Inventory::whereIn('id', $latestInventoryIds)->get();


    //         $verifiedStockData = $latestInventoryData
    //             ->filter(fn ($inv) => $inv->is_hide == 0)
    //             ->mapWithKeys(fn ($inv) => [trim($inv->sku) => $inv]);

    //         $hiddenSkuSet = $latestInventoryData
    //             ->filter(fn ($inv) => $inv->is_hide == 1)
    //             ->pluck('sku')
    //             ->map(fn ($sku) => trim($sku))
    //             ->toArray();

    //         // $verifiedStockData = Inventory::whereIn('sku', $skus)->get()
    //         //     ->mapWithKeys(function ($item) {
    //         //         return [trim($item->sku) => $item]; 
    //         //     });


    //         $filteredData = array_filter($data->data, function ($item) use ($hiddenSkuSet) {
    //             $sku = trim($item->SKU ?? '');
    //             return !(empty(trim($item->Parent ?? '')) && empty($sku)) && !in_array($sku, $hiddenSkuSet);
    //         });

    //         // $filteredData = array_filter($data->data, function ($item) {
    //         //     return !(empty(trim($item->Parent ?? '')) && empty(trim($item->{'SKU'} ?? '')));
    //         // });

    //         $mergedData = collect($filteredData)->map(function ($item) use ($shopifyData, $inventoryData, $verifiedStockData) {
    //             $childSku = trim($item->{'SKU'} ?? ''); 
    //             $lp = isset($item->LP) && is_numeric($item->LP) ? floatval($item->LP) : 0;

    //             if (!empty($childSku) && stripos($childSku, 'PARENT') === false) {
    //                 if ($shopifyData->has($childSku)) {
    //                     $item->INV = $shopifyData[$childSku]->inv;
    //                     $item->L30 = $shopifyData[$childSku]->quantity;
    //                     $item->IMAGE_URL = $shopifyData[$childSku]->image_url ?? null;
    //                 } else {
    //                     $item->INV = 0;
    //                     $item->L30 = 0;
    //                     $item->IMAGE_URL = null;
    //                 }

    //                 if (isset($inventoryData[$childSku])) {
    //                     $item->ON_HAND = $inventoryData[$childSku]['on_hand'];
    //                     $item->COMMITTED = $inventoryData[$childSku]['committed'];
    //                     $item->AVAILABLE_TO_SELL = $inventoryData[$childSku]['available_to_sell'];
    //                     $item->IMAGE_URL = $inventoryData[$childSku]['image_url'] ?? $item->IMAGE_URL;

    //                     ShopifyInventory::updateOrCreate(
    //                         ['sku' => $childSku],
    //                         [
    //                             'parent' => $item->Parent ?? null,
    //                             'on_hand' => $inventoryData[$childSku]['on_hand'],
    //                             'committed' => $inventoryData[$childSku]['committed'],
    //                             'available_to_sell' => $inventoryData[$childSku]['available_to_sell'],
    //                         ]
    //                     );
    //                 } else {
    //                     $item->ON_HAND = 'N/A';
    //                     $item->AVAILABLE_TO_SELL = 'N/A';
    //                     $item->COMMITTED = 'N/A';
    //                 }

    //                 if ($verifiedStockData->has($childSku)) {
    //                     $verifiedStockRow = $verifiedStockData[$childSku];
    //                     $item->VERIFIED_STOCK = $verifiedStockRow->verified_stock ?? null;
    //                     $item->TO_ADJUST = $verifiedStockRow->to_adjust ?? null;
    //                     $item->REASON = $verifiedStockRow->reason ?? null;
    //                     $item->REMARKS = $verifiedStockRow->REMARKS ?? null;
    //                     $item->APPROVED = (bool) $verifiedStockRow->approved;
    //                     $item->APPROVED_BY = $verifiedStockRow->approved_by ?? null;
    //                     $item->APPROVED_AT = $verifiedStockRow->approved_at ?? null;
    //                 } else {
    //                     $item->VERIFIED_STOCK = null;
    //                     $item->TO_ADJUST = null;
    //                     $item->REASON = null;
    //                     $item->REMARKS = null;
    //                     $item->APPROVED = false;
    //                     $item->APPROVED_BY = null;
    //                     $item->APPROVED_AT = null;
    //                 }

    //                 $adjustedQty = isset($item->TO_ADJUST) && is_numeric($item->TO_ADJUST) ? floatval($item->TO_ADJUST) : 0;
    //                 $item->LOSS_GAIN = round($adjustedQty * $lp, 2);
    //             }

    //             return $item;
    //         });

    //         $processedData = $mergedData->values();
    //         Log::info('Processed data count: ' . count($processedData));

    //         return response()->json([
    //             'message' => 'Data fetched successfully',
    //             'data' => $processedData,
    //             'status' => 200
    //         ]);
    //     } else {
    //         return response()->json([
    //             'message' => 'Failed to fetch data from Google Sheet',
    //             'status' => $response->getStatusCode()
    //         ], $response->getStatusCode());
    //     }
    // }



    // public function getViewVerificationAdjustmentData(Request $request)     
    // {
    //     $shopifyInventoryController = new ShopifyApiInventoryController();
    //     $inventoryResponse = $shopifyInventoryController->fetchInventoryWithCommitment();
    //     Log::info('Fetched inventory response:', $inventoryResponse);

    //     $inventoryData = collect($inventoryResponse); 
    //     Log::info('Shopify Inventory Data:', $inventoryData->toArray());

    //     // $response = $this->apiController->fetchDataFromProductMasterGoogleSheet();

    //     $productMasters = ProductMaster::orderBy('parent', 'asc')
    //         ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
    //         ->orderBy('sku', 'asc')
    //         ->get();

    //     if ($productMasters->isNotEmpty()) {
    //         $data = $productMasters;

    //         // use trimmed SKUs only (no strtoupper)
    //         $skus = $productMasters->pluck('sku')
    //             ->map(fn($s) => trim($s))
    //             ->filter()
    //             ->unique()
    //             ->toArray();

    //         // use exact-case SKU keys
    //         $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(function ($item) {
    //             return trim($item->sku); 
    //         });

    //         $latestInventoryIds = Inventory::select(DB::raw('MAX(id) as latest_id'))
    //             ->whereIn('sku', $skus)
    //             ->groupBy('sku')
    //             ->pluck('latest_id');

    //         $latestInventoryData = Inventory::whereIn('id', $latestInventoryIds)->get();


    //         $verifiedStockData = $latestInventoryData
    //             ->filter(fn ($inv) => $inv->is_hide == 0)
    //             ->mapWithKeys(fn ($inv) => [trim($inv->sku) => $inv]);

    //         $hiddenSkuSet = $latestInventoryData
    //             ->filter(fn ($inv) => $inv->is_hide == 1)
    //             ->pluck('sku')
    //             ->map(fn ($sku) => trim($sku))
    //             ->toArray();

    //         // $verifiedStockData = Inventory::whereIn('sku', $skus)->get()
    //         //     ->mapWithKeys(function ($item) {
    //         //         return [trim($item->sku) => $item]; 
    //         //     });


    //         // $filteredData = array_filter($data->data, function ($item) use ($hiddenSkuSet) {
    //         //     $sku = trim($item->SKU ?? '');
    //         //     return !(empty(trim($item->Parent ?? '')) && empty($sku)) && !in_array($sku, $hiddenSkuSet);
    //         // });
    //         $filteredData = $data->filter(function ($item) use ($hiddenSkuSet) {
    //             $sku = trim($item->sku ?? ''); // lowercase because it's DB column
    //             return !(empty(trim($item->parent ?? '')) && empty($sku)) && !in_array($sku, $hiddenSkuSet);
    //         });

    //         // $filteredData = array_filter($data->data, function ($item) {
    //         //     return !(empty(trim($item->Parent ?? '')) && empty(trim($item->{'SKU'} ?? '')));
    //         // });

    //         $mergedData = collect($filteredData)->map(function ($item) use ($shopifyData, $inventoryData, $verifiedStockData) {
    //             $childSku = trim($item->{'SKU'} ?? ''); 
    //             $lp = isset($item->LP) && is_numeric($item->LP) ? floatval($item->LP) : 0;

    //             if (!empty($childSku) && stripos($childSku, 'PARENT') === false) {
    //                 if ($shopifyData->has($childSku)) {
    //                     $item->INV = $shopifyData[$childSku]->inv;
    //                     $item->L30 = $shopifyData[$childSku]->quantity;
    //                     $item->IMAGE_URL = $shopifyData[$childSku]->image_url ?? null;
    //                 } else {
    //                     $item->INV = 0;
    //                     $item->L30 = 0;
    //                     $item->IMAGE_URL = null;
    //                 }

    //                 if (isset($inventoryData[$childSku])) {
    //                     $item->ON_HAND = $inventoryData[$childSku]['on_hand'];
    //                     $item->COMMITTED = $inventoryData[$childSku]['committed'];
    //                     $item->AVAILABLE_TO_SELL = $inventoryData[$childSku]['available_to_sell'];
    //                     $item->IMAGE_URL = $inventoryData[$childSku]['image_url'] ?? $item->IMAGE_URL;

    //                     ShopifyInventory::updateOrCreate(
    //                         ['sku' => $childSku],
    //                         [
    //                             'parent' => $item->Parent ?? null,
    //                             'on_hand' => $inventoryData[$childSku]['on_hand'],
    //                             'committed' => $inventoryData[$childSku]['committed'],
    //                             'available_to_sell' => $inventoryData[$childSku]['available_to_sell'],
    //                         ]
    //                     );
    //                 } else {
    //                     $item->ON_HAND = 'N/A';
    //                     $item->AVAILABLE_TO_SELL = 'N/A';
    //                     $item->COMMITTED = 'N/A';
    //                 }

    //                 if ($verifiedStockData->has($childSku)) {
    //                     $verifiedStockRow = $verifiedStockData[$childSku];
    //                     $item->VERIFIED_STOCK = $verifiedStockRow->verified_stock ?? null;
    //                     $item->TO_ADJUST = $verifiedStockRow->to_adjust ?? null;
    //                     $item->REASON = $verifiedStockRow->reason ?? null;
    //                     $item->REMARKS = $verifiedStockRow->REMARKS ?? null;
    //                     $item->APPROVED = (bool) $verifiedStockRow->approved;
    //                     $item->APPROVED_BY = $verifiedStockRow->approved_by ?? null;
    //                     $item->APPROVED_AT = $verifiedStockRow->approved_at ?? null;
    //                 } else {
    //                     $item->VERIFIED_STOCK = null;
    //                     $item->TO_ADJUST = null;
    //                     $item->REASON = null;
    //                     $item->REMARKS = null;
    //                     $item->APPROVED = false;
    //                     $item->APPROVED_BY = null;
    //                     $item->APPROVED_AT = null;
    //                 }

    //                 $adjustedQty = isset($item->TO_ADJUST) && is_numeric($item->TO_ADJUST) ? floatval($item->TO_ADJUST) : 0;
    //                 $item->LOSS_GAIN = round($adjustedQty * $lp, 2);
    //             }

    //             return $item;
    //         });

    //         $processedData = $mergedData->values();
    //         Log::info('Processed data count: ' . count($processedData));

    //         return response()->json([
    //             'message' => 'Data fetched successfully',
    //             'data' => $processedData,
    //             'status' => 200
    //         ]);
    //     } else {
    //         return response()->json([
    //             'message' => 'No product masters found',
    //             'status' => 404
    //         ], 404);
    //     }
    // }





    // public function getViewVerificationAdjustmentData(Request $request)  //current
    // {
    //     // Fetch inventory data from Shopify
    //     $shopifyInventoryController = new ShopifyApiInventoryController();
    //     $inventoryResponse = $shopifyInventoryController->fetchInventoryWithCommitment();
    //     Log::info('Fetched inventory response:', $inventoryResponse);

    //     $inventoryData = collect($inventoryResponse); 
    //     Log::info('Shopify Inventory Data:', $inventoryData->toArray());

    //     // Fetch data from the local product_master table
    //     $productMasterData = ProductMaster::all();

    //     if ($productMasterData) {
    //         $skus = $productMasterData
    //             ->pluck('sku')
    //             ->filter()
    //             ->unique()
    //             ->map(fn ($sku) => trim($sku))
    //             ->toArray();

    //         // Fetch Shopify SKU data
    //         $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(function ($item) {
    //             return trim($item->sku); 
    //         });

    //         // Get the latest inventory data for each SKU
    //         $latestInventoryIds = Inventory::select(DB::raw('MAX(id) as latest_id'))
    //             ->whereIn('sku', $skus)
    //             ->groupBy('sku')
    //             ->pluck('latest_id');

    //         $latestInventoryData = Inventory::whereIn('id', $latestInventoryIds)->get();

    //         // Separate verified and hidden stock data
    //         $verifiedStockData = $latestInventoryData
    //             ->filter(fn ($inv) => $inv->is_hide == 0)
    //             ->mapWithKeys(fn ($inv) => [trim($inv->sku) => $inv]);

    //         $hiddenSkuSet = $latestInventoryData
    //             ->filter(fn ($inv) => $inv->is_hide == 1)
    //             ->pluck('sku')
    //             ->map(fn ($sku) => trim($sku))
    //             ->toArray();

    //         // Filter out hidden SKUs from the main dataset
    //         $filteredData = $productMasterData->filter(function ($item) use ($hiddenSkuSet) {
    //             $sku = trim($item->sku ?? '');
    //             return !in_array($sku, $hiddenSkuSet);
    //         });

    //         // Merge all data sources
    //         $mergedData = $filteredData->map(function ($item) use ($shopifyData, $inventoryData, $verifiedStockData) {
    //             $childSku = trim($item->sku ?? ''); 
    //             // $item->IS_PARENT = (stripos($childSku, 'PARENT') === 0);
    //             $isParent = (stripos($childSku, 'PARENT') === 0);
    //             $item->IS_PARENT = $isParent;

    //             // Decode the JSON string in the 'Values' column
    //             // $values = json_decode($item->Values, true);
    //             $values = $item->values;
    //             $lp = $values['lp'] ?? 0;

    //             if (!$isParent) {
    //                 // Add Shopify data
    //                 if ($shopifyData->has($childSku)) {
    //                     $item->INV = $shopifyData[$childSku]->inv;
    //                     $item->L30 = $shopifyData[$childSku]->quantity;
    //                     $item->IMAGE_URL = $shopifyData[$childSku]->image_url ?? null;
    //                 } else {
    //                     $item->INV = 0;
    //                     $item->L30 = 0;
    //                     $item->IMAGE_URL = null;
    //                 }

    //                 // Add Shopify inventory data
    //                 if (isset($inventoryData[$childSku])) {
    //                     $item->ON_HAND = $inventoryData[$childSku]['on_hand'];
    //                     $item->COMMITTED = $inventoryData[$childSku]['committed'];
    //                     $item->AVAILABLE_TO_SELL = $inventoryData[$childSku]['available_to_sell'];
    //                     $item->IMAGE_URL = $inventoryData[$childSku]['image_url'] ?? $item->IMAGE_URL;

    //                     // Update or create ShopifyInventory record
    //                     ShopifyInventory::updateOrCreate(
    //                         ['sku' => $childSku],
    //                         [
    //                             'parent' => $item->parent ?? null,
    //                             'on_hand' => $inventoryData[$childSku]['on_hand'],
    //                             'committed' => $inventoryData[$childSku]['committed'],
    //                             'available_to_sell' => $inventoryData[$childSku]['available_to_sell'],
    //                         ]
    //                     );
    //                 } else {
    //                     $item->ON_HAND = 'N/A';
    //                     $item->AVAILABLE_TO_SELL = 'N/A';
    //                     $item->COMMITTED = 'N/A';
    //                 }

    //                 // Add verified stock data
    //                 if ($verifiedStockData->has($childSku)) {
    //                     $verifiedStockRow = $verifiedStockData[$childSku];
    //                     $item->VERIFIED_STOCK = $verifiedStockRow->verified_stock ?? null;
    //                     $item->TO_ADJUST = $verifiedStockRow->to_adjust ?? null;
    //                     $item->REASON = $verifiedStockRow->reason ?? null;
    //                     $item->REMARKS = $verifiedStockRow->REMARKS ?? null;
    //                     $item->APPROVED = (bool) $verifiedStockRow->approved;
    //                     $item->APPROVED_BY = $verifiedStockRow->approved_by ?? null;
    //                     $item->APPROVED_AT = $verifiedStockRow->approved_at ?? null;
    //                 } else {
    //                     $item->VERIFIED_STOCK = null;
    //                     $item->TO_ADJUST = null;
    //                     $item->REASON = null;
    //                     $item->REMARKS = null;
    //                     $item->APPROVED = false;
    //                     $item->APPROVED_BY = null;
    //                     $item->APPROVED_AT = null;
    //                 }

    //                 // Calculate loss/gain
    //                 $adjustedQty = isset($item->TO_ADJUST) && is_numeric($item->TO_ADJUST) ? floatval($item->TO_ADJUST) : 0;
    //                 $item->LOSS_GAIN = round($adjustedQty * $lp, 2);
    //             }

    //             return $item;
    //         });

    //         $processedData = $mergedData->values();
    //         Log::info('Processed data count: ' . count($processedData));

    //         return response()->json([
    //             'message' => 'Data fetched successfully',
    //             'data' => $processedData,
    //             'status' => 200
    //         ]);
    //     } else {
    //         return response()->json([
    //             'message' => 'Failed to fetch data from product_master table',
    //             'status' => 500
    //         ], 500);
    //     }
    // }


    public function getViewVerificationAdjustmentData(Request $request)
    {
        // Fetch inventory data from Shopify
        $shopifyInventoryController = new ShopifyApiInventoryController();
        $inventoryResponse = $shopifyInventoryController->fetchInventoryWithCommitment();
        // Log::info('Fetched inventory response:', $inventoryResponse);

        $inventoryData = collect($inventoryResponse);
        // Log::info('Shopify Inventory Data:', $inventoryData->toArray());

        // Fetch data from the local product_master table
        $productMasterData = ProductMaster::all();

        if ($productMasterData->isEmpty()) {
            return response()->json([
                'message' => 'Failed to fetch data from product_master table',
                'status' => 500
            ], 500);
        }

        // Normalize helper
        $normalizeSku = fn($sku) => strtoupper(trim(preg_replace('/\s+/', ' ', $sku)));

        $skus = $productMasterData
            ->pluck('sku')
            ->filter()
            ->unique()
            ->map(fn($sku) => $normalizeSku($sku))
            ->toArray();

        // Fetch Shopify SKU data (local DB)
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(function ($item) use ($normalizeSku) {
            return $normalizeSku($item->sku);
        });

        // Get the latest inventory data for each SKU
        $latestInventoryIds = Inventory::select(DB::raw('MAX(id) as latest_id'))
            ->whereIn('sku', $skus)
            ->groupBy('sku')
            ->pluck('latest_id');

        $latestInventoryData = Inventory::whereIn('id', $latestInventoryIds)->get();

        // Separate verified and hidden stock data
        $verifiedStockData = $latestInventoryData
            ->filter(fn($inv) => $inv->is_hide == 0)
            ->mapWithKeys(fn($inv) => [$normalizeSku($inv->sku) => $inv]);

        $hiddenSkuSet = $latestInventoryData
            ->filter(fn($inv) => $inv->is_hide == 1)
            ->pluck('sku')
            ->map(fn($sku) => $normalizeSku($sku))
            ->toArray();

        // Filter out hidden SKUs from the main dataset
        $filteredData = $productMasterData->filter(function ($item) use ($hiddenSkuSet, $normalizeSku) {
            $sku = $normalizeSku($item->sku ?? '');
            return !in_array($sku, $hiddenSkuSet);
        });

        // Merge all data sources
        $mergedData = $filteredData->map(function ($item) use ($shopifyData, $inventoryData, $verifiedStockData, $normalizeSku) {
            $childSku = $normalizeSku($item->sku ?? '');
            $isParent = (stripos($childSku, 'PARENT') === 0);
            $item->IS_PARENT = $isParent;

            // Decode JSON values column
            $values = $item->values;
            $lp = $values['lp'] ?? 0;

            if (!$isParent) {
                // Add Shopify SKU (local DB) data
                if ($shopifyData->has($childSku)) {
                    $item->INV = $shopifyData[$childSku]->inv;
                    $item->L30 = $shopifyData[$childSku]->quantity;
                    $item->IMAGE_URL = $shopifyData[$childSku]->image_url ?? null;
                } else {
                    $item->INV = 0;
                    $item->L30 = 0;
                    $item->IMAGE_URL = null;
                }

                // Add Shopify API inventory data
                if ($inventoryData->has($childSku)) {
                    $shopifyRow = $inventoryData->get($childSku);

                    $item->ON_HAND = $shopifyRow['on_hand'];
                    $item->COMMITTED = $shopifyRow['committed'];
                    $item->AVAILABLE_TO_SELL = $shopifyRow['available_to_sell'];
                    $item->IMAGE_URL = $shopifyRow['image_url'] ?? $item->IMAGE_URL;

                    // Ensure SKU is stored/updated in ShopifyInventory table
                    ShopifyInventory::updateOrCreate(
                        ['sku' => $childSku],
                        [
                            'parent' => $item->parent ?? null,
                            'on_hand' => $shopifyRow['on_hand'],
                            'committed' => $shopifyRow['committed'],
                            'available_to_sell' => $shopifyRow['available_to_sell'],
                        ]
                    );
                } else {
                    $item->ON_HAND = 'N/A';
                    $item->AVAILABLE_TO_SELL = 'N/A';
                    $item->COMMITTED = 'N/A';
                }

                // Add verified stock data
                if ($verifiedStockData->has($childSku)) {
                    $verifiedStockRow = $verifiedStockData[$childSku];
                    $item->VERIFIED_STOCK = $verifiedStockRow->verified_stock ?? null;
                    $item->TO_ADJUST = $verifiedStockRow->to_adjust ?? null;
                    $item->REASON = $verifiedStockRow->reason ?? null;
                    $item->REMARKS = $verifiedStockRow->REMARKS ?? null;
                    $item->APPROVED = (bool) $verifiedStockRow->approved;
                    $item->APPROVED_BY = $verifiedStockRow->approved_by ?? null;
                    $item->APPROVED_AT = $verifiedStockRow->approved_at ?? null;
                } else {
                    $item->VERIFIED_STOCK = null;
                    $item->TO_ADJUST = null;
                    $item->REASON = null;
                    $item->REMARKS = null;
                    $item->APPROVED = false;
                    $item->APPROVED_BY = null;
                    $item->APPROVED_AT = null;
                }

                // Calculate loss/gain
                $adjustedQty = isset($item->TO_ADJUST) && is_numeric($item->TO_ADJUST) ? floatval($item->TO_ADJUST) : 0;
                $item->LOSS_GAIN = round($adjustedQty * $lp, 2);
            }

            return $item;
        });

        $processedData = $mergedData->values();
        // Log::info('Processed data count: ' . count($processedData));

        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $processedData,
            'status' => 200
        ]);
    }


    // public function updateVerifiedStock(Request $request)   //current
    // {
    //     $validated = $request->validate([
    //         'sku' => 'nullable|string',
    //         'verified_stock' => 'required|numeric',
    //         'on_hand' => 'nullable|numeric',
    //         'reason' => 'required|string',
    //         'remarks' => 'nullable|string',
    //         'is_approved' => 'required|boolean',
    //     ]);

    //     $lp = 0;
    //     $sku = trim($validated['sku']);
    //     $product = ProductMaster::whereRaw('LOWER(sku) = ?', [strtolower($sku)])->first();

    //     if ($product) {
    //         // $values = json_decode($product->Values, true);
    //         $values = $product->Values; 
    //         if (isset($values['lp']) && is_numeric($values['lp'])) {
    //             $lp = floatval($values['lp']);
    //         }
    //     } 
    //     // $response = $this->apiController->fetchDataFromProductMasterGoogleSheet(); 
    //     // if ($response->getStatusCode() === 200) { 
    //     //     $sheetData = $response->getData()->data; 
    //     //     foreach ($sheetData as $row) { 
    //     //         if (isset($row->SKU) && strtoupper(trim($row->SKU)) === strtoupper(trim($validated['sku']))) { 
    //     //             $lp = isset($row->LP) && is_numeric($row->LP) ? floatval($row->LP) : 0; 
    //     //             break; 
    //     //         }
    //     //     }
    //     // }

    //     $toAdjust = $validated['verified_stock'] - ($validated['on_hand'] ?? 0);
    //     $lossGain = round($toAdjust * $lp, 2);


    //     // Save record in DB
    //     $record = new Inventory();
    //     $record->sku = $validated['sku'];
    //     $record->on_hand = $validated['on_hand'];
    //     $record->verified_stock = $validated['verified_stock'];
    //     $record->reason = $validated['reason'];
    //     $record->remarks = $validated['remarks'];
    //     $record->is_approved = $validated['is_approved'];
    //     $record->approved_by = $validated['is_approved'] ? Auth::user()->name : null;
    //     $record->approved_at = $validated['is_approved'] ? Carbon::now('America/New_York') : null;
    //     $record->to_adjust = $toAdjust;
    //     $record->loss_gain = $lossGain;
    //     $record->is_hide = 0;
    //     $record->save();

    //     if ($validated['is_approved']) {
    //         $sku = $validated['sku'];
    //         // $verifiedToAdd = $validated['verified_stock']; // This is the value to add

    //         // 1. Fetch all products (with pagination to ensure all SKUs are fetched)
    //         $inventoryItemId = null;
    //         $pageInfo = null;

    //         do {
    //             $queryParams = ['limit' => 250];
    //             if ($pageInfo) {
    //                 $queryParams['page_info'] = $pageInfo;
    //             }

    //             $response = Http::withBasicAuth($this->shopifyApiKey, $this->shopifyPassword)
    //                 ->get("https://{$this->shopifyDomain}/admin/api/2025-01/products.json", $queryParams);

    //             $products = $response->json('products');

    //             foreach ($products as $product) {
    //                 foreach ($product['variants'] as $variant) {
    //                     if ($variant['sku'] === $sku) {
    //                         $inventoryItemId = $variant['inventory_item_id'];
    //                         break 2;
    //                     }
    //                 }
    //             }

    //             // Handle pagination
    //             $linkHeader = $response->header('Link');
    //             $pageInfo = null;
    //             if ($linkHeader && preg_match('/<([^>]+page_info=([^&>]+)[^>]*)>; rel="next"/', $linkHeader, $matches)) {
    //                 $pageInfo = $matches[2];
    //             }

    //         } while (!$inventoryItemId && $pageInfo);

    //         if (!$inventoryItemId) {
    //             return response()->json(['success' => false, 'message' => 'Inventory item ID not found for SKU.']);
    //         }

    //         // 2. Get location ID and current available
    //         $invLevelResponse = Http::withBasicAuth($this->shopifyApiKey, $this->shopifyPassword)
    //             ->get("https://{$this->shopifyDomain}/admin/api/2025-01/inventory_levels.json", [
    //                 'inventory_item_ids' => $inventoryItemId
    //             ]);

    //         $levels = $invLevelResponse->json('inventory_levels');
    //         $locationId = $levels[0]['location_id'] ?? null;
    //         // $currentAvailable = $levels[0]['available'] ?? 0;

    //         if (!$locationId) {
    //             return response()->json(['success' => false, 'message' => 'Location ID not found for inventory item.']);
    //         }

    //         // 4. Send inventory adjustment to Shopify
    //         $adjustResponse = Http::withBasicAuth($this->shopifyApiKey, $this->shopifyPassword)
    //             ->post("https://{$this->shopifyDomain}/admin/api/2025-01/inventory_levels/adjust.json", [
    //                 'inventory_item_id' => $inventoryItemId,
    //                 'location_id' => $locationId,
    //                 'available_adjustment' => $toAdjust,
    //             ]);

    //         Log::info('Shopify Adjust Response:', $adjustResponse->json());

    //         if (!$adjustResponse->successful()) {
    //             return response()->json(['success' => false, 'message' => 'Failed to update Shopify inventory.']);
    //         }
    //     }

    //     return response()->json([
    //         'success' => true,
    //         'data' => [
    //             'sku' => $record->sku,
    //             'verified_stock' => $record->verified_stock,
    //             'reason' => $record->reason,
    //             'remarks' => $record->remarks,
    //             'is_approved' => $record->is_approved,
    //             'approved_by' => $record->approved_by,
    //             'approved_at' => optional($record->approved_at)->format('Y-m-d\TH:i:s.u\Z'),
    //             'created_at' => optional($record->created_at)->format('Y-m-d\TH:i:s.u\Z'),
    //             'updated_at' => optional($record->updated_at)->format('Y-m-d\TH:i:s.u\Z'),
    //             'to_adjust' => $record->to_adjust,
    //             'loss_gain' => $lossGain, // Only used in response, not stored
    //         ]
    //     ]);
    // }


    public function updateVerifiedStock(Request $request)
    {
        $validated = $request->validate([
            'sku' => 'nullable|string',
            'verified_stock' => 'required|numeric',
            'on_hand' => 'nullable|numeric',
            'reason' => 'required|string',
            'remarks' => 'nullable|string',
            'is_approved' => 'required|boolean',
        ]);

        $lp = 0;
        $sku = trim($validated['sku']);
        $product = ProductMaster::whereRaw('LOWER(sku) = ?', [strtolower($sku)])->first();

        if ($product) {
            $values = $product->Values; 
            if (isset($values['lp']) && is_numeric($values['lp'])) {
                $lp = floatval($values['lp']);
            }
        }

        $toAdjust = $validated['verified_stock'] - ($validated['on_hand'] ?? 0);
        $lossGain = round($toAdjust * $lp, 2);

        // Save record in DB
        $record = new Inventory();
        $record->sku = $sku;
        $record->on_hand = $validated['on_hand'];
        $record->verified_stock = $validated['verified_stock'];
        $record->reason = $validated['reason'];
        $record->remarks = $validated['remarks'];
        $record->is_approved = $validated['is_approved'];
        $record->approved_by = $validated['is_approved'] ? Auth::user()->name : null;
        $record->approved_at = $validated['is_approved'] ? Carbon::now('America/New_York') : null;
        $record->to_adjust = $toAdjust;
        $record->loss_gain = $lossGain;
        $record->is_hide = 0;
        $record->save();

        if ($validated['is_approved']) {
            $inventoryItemId = null;
            $pageInfo = null;

            // Normalize input SKU: replace all whitespace (normal + non-breaking) with single space
            $normalizedSku = strtoupper(preg_replace('/\s+/u', ' ', $sku));

            do {
                $queryParams = ['limit' => 250];
                if ($pageInfo) {
                    $queryParams['page_info'] = $pageInfo;
                }

                $response = Http::withBasicAuth($this->shopifyApiKey, $this->shopifyPassword)
                    ->get("https://{$this->shopifyDomain}/admin/api/2025-01/products.json", $queryParams);

                if (!$response->successful()) {
                    // Log::error('Failed to fetch products from Shopify', ['status' => $response->status()]);
                    break;
                }

                $products = $response->json('products');

                foreach ($products as $product) {
                    foreach ($product['variants'] as $variant) {
                        $variantSku = strtoupper(preg_replace('/\s+/u', ' ', trim($variant['sku'] ?? '')));
                        if ($variantSku === $normalizedSku) {
                            $inventoryItemId = $variant['inventory_item_id'];
                            break 2;
                        }
                    }
                }

                // Handle pagination
                $linkHeader = $response->header('Link');
                $pageInfo = null;
                if ($linkHeader && preg_match('/<([^>]+page_info=([^&>]+)[^>]*)>; rel="next"/', $linkHeader, $matches)) {
                    $pageInfo = $matches[2];
                }

            } while (!$inventoryItemId && $pageInfo);

            if (!$inventoryItemId) {
                // Log::error('Inventory item ID not found for SKU', ['sku' => $sku]);
                return response()->json(['success' => false, 'message' => 'Inventory item ID not found for SKU.']);
            }

            // Get location ID
            $invLevelResponse = Http::withBasicAuth($this->shopifyApiKey, $this->shopifyPassword)
                ->get("https://{$this->shopifyDomain}/admin/api/2025-01/inventory_levels.json", [
                    'inventory_item_ids' => $inventoryItemId
                ]);

            $levels = $invLevelResponse->json('inventory_levels');
            $locationId = $levels[0]['location_id'] ?? null;

            if (!$locationId) {
                return response()->json(['success' => false, 'message' => 'Location ID not found for inventory item.']);
            }

            // Adjust inventory in Shopify
            $adjustResponse = Http::withBasicAuth($this->shopifyApiKey, $this->shopifyPassword)
                ->post("https://{$this->shopifyDomain}/admin/api/2025-01/inventory_levels/adjust.json", [
                    'inventory_item_id' => $inventoryItemId,
                    'location_id' => $locationId,
                    'available_adjustment' => $toAdjust,
                ]);

            // Log::info('Shopify Adjust Response', $adjustResponse->json());

            if (!$adjustResponse->successful()) {
                return response()->json(['success' => false, 'message' => 'Failed to update Shopify inventory.']);
            }
        }

        return response()->json([
            'success' => true,
            'data' => [
                'sku' => $record->sku,
                'verified_stock' => $record->verified_stock,
                'reason' => $record->reason,
                'remarks' => $record->remarks,
                'is_approved' => $record->is_approved,
                'approved_by' => $record->approved_by,
                'approved_at' => optional($record->approved_at)->format('Y-m-d\TH:i:s.u\Z'),
                'created_at' => optional($record->created_at)->format('Y-m-d\TH:i:s.u\Z'),
                'updated_at' => optional($record->updated_at)->format('Y-m-d\TH:i:s.u\Z'),
                'to_adjust' => $record->to_adjust,
                'loss_gain' => $lossGain,
            ]
        ]);
    }



    public function getVerifiedStock()
    {
        $savedInventories = Inventory::all();


        // Format data to return in JSON with key 'data'
        $data = $savedInventories->map(function ($item) {

            return [
                'sku' => strtoupper(trim($item->sku)),
                'R&A' => (bool) $item->is_ra_checked,
                'verified_stock' => $item->verified_stock,
                'reason' => $item->reason,
                'is_approved' => (bool) $item->is_approved,
                'approved_by_ih' => (bool) $item->approved_by_ih,
                'approved_by' => $item->approved_by,
                'approved_at' =>  $item->approved_at,
            ];
        });

        return response()->json(['data' => $data]);
    }

    public function updateApprovedByIH(Request $request)
    {
        $request->validate([
            'sku' => 'required|string',
            'approved_by_ih' => 'required|boolean',
        ]);

        $inventory = Inventory::where('sku', $request->sku)->first();

        if (!$inventory) {
            return response()->json(['success' => false, 'message' => 'SKU not found.']);
        }

        $inventory->approved_by_ih = $request->approved_by_ih;
        $inventory->save();

        return response()->json(['success' => true]);
    }


    public function updateRAStatus(Request $request)
    {
        $validated = $request->validate([
            'sku' => 'required|string',
            'is_ra_checked' => 'required|boolean'
        ]);

        $inventory = Inventory::where('sku', $validated['sku'])->first();

        if ($inventory) {
            // SKU exists → Only update is_ra_checked
            $inventory->is_ra_checked = $validated['is_ra_checked'];
            $inventory->save();
        } else {
            //  SKU not found → Create new record
            $inventory = Inventory::create([
                'sku' => $validated['sku'],
                'is_ra_checked' => $validated['is_ra_checked'],
            ]);
        }

        return response()->json(['success' => true]);
    }

    public function getVerifiedStockActivityLog()
    {
        $activityLogs = Inventory::where('type', null)->where('is_approved', true)
            ->orderByDesc('created_at')
            ->get()
            ->map(function ($item) {
                return [
                    'sku' => $item->sku,
                    'verified_stock' => $item->verified_stock,
                    'to_adjust' => $item->to_adjust,
                    'loss_gain' => $item->loss_gain,
                    'reason' => $item->reason,
                    'remarks' => $item->remarks,
                    'approved_by' => $item->approved_by,
                    'approved_at' => Carbon::parse($item->created_at)->timezone('America/New_York')->format('d M Y, h:i A'),
                ];
            });
            
        return response()->json(['data' => $activityLogs]);
    }


    public function viewInventory()
    {
        return view('inventory-management.view-inventory');
    }

    public function getSkuWiseHistory(Request $request)
    {
        $sku = $request->input('sku');

        $query = Inventory::where('is_approved', true);

        if ($sku) {
            $query->where('sku', $sku);
        }

        $activityLogs = $query->orderByDesc('created_at')
            ->get()
            ->map(function ($item) {
                return [
                    'sku' => $item->sku,
                    'verified_stock' => $item->verified_stock,
                    'to_adjust' => $item->to_adjust,
                    'on_hand' => $item->on_hand,
                    'reason' => $item->reason,
                    'approved_by' => $item->approved_by,
                    'approved_at' => Carbon::parse($item->created_at)->timezone('America/New_York')->format('d M Y, h:i A'),
                ];
            });

        return response()->json(['data' => $activityLogs]);
    }


    public function toggleHide(Request $request)
    {
        $latestRecord = Inventory::where('sku', $request->sku)->latest()->first();

        if ($latestRecord) {
            $latestRecord->update(['is_hide' => 1]);
            return response()->json(['success' => true]);
        }

        return response()->json(['success' => false, 'message' => 'Record not found.']);
    }


    public function getHiddenRows()
    {
        $latestHiddenIds = Inventory::select(DB::raw('MAX(id) as latest_id'))
            ->where('is_hide', 1)
            ->groupBy('sku')
            ->pluck('latest_id');

            Inventory::where('is_hide', 1)
            ->whereNotIn('id', $latestHiddenIds)
            ->update(['is_hide' => 0]);

        $hiddenRecords = Inventory::whereIn('id', $latestHiddenIds)->get();

        $data = $hiddenRecords->map(function ($item) {
            return [
                'sku' => $item->sku,
                'verified_stock' => $item->verified_stock,
                'to_adjust' => $item->to_adjust,
                'loss_gain' => $item->loss_gain, // already stored in DB
                'reason' => $item->reason,
                'approved_by' => $item->approved_by,
                'approved_at' => $item->approved_at 
                    ? Carbon::parse($item->approved_at)->timezone('America/New_York')->format('Y-m-d H:i:s') 
                    : null,
                'remarks' => $item->remarks ?? '-',
            ];
        });

        return response()->json(['data' => $data]);
    }


    public function unhideMultipleRows(Request $request)
    {
        $skus = $request->skus ?? [];

        foreach ($skus as $sku) {
            $latest = Inventory::where('sku', $sku)->where('is_hide', 1)->latest()->first();
            if ($latest) {
                $latest->update(['is_hide' => 0]);
            }
        }

        return response()->json(['success' => true]);
    }

    
}
