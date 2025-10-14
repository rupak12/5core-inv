<?php

namespace App\Http\Controllers\MarketPlace;

use App\Http\Controllers\ApiController;
use App\Http\Controllers\Controller;
use App\Models\MarketplacePercentage;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\JungleScoutProductData;
use App\Models\WayfairDataView;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Cache;

class WayfairZeroController extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }

    public function wayfairZeroView(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        // Get percentage from cache or database
        $percentage = Cache::remember('wayfair_marketplace_percentage', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'Wayfair')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100; // Default to 100 if not set
        });

        return view('market-places.WayfairZeroview', [
            'mode' => $mode,
            'demo' => $demo,
            'wayfairPercentage' => $percentage

        ]);
    }
    public function getViewWayfairZeroData(Request $request)
    {
        // 1. Fetch all ProductMaster rows as the base
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        // 2. Fetch Google Sheet data
        $response = $this->apiController->fetchDataFromWayfairMasterGoogleSheet();

        if ($response->getStatusCode() === 200) {
            $sheetData = $response->getData();
            // Index sheet data by SKU for fast lookup
            $sheetSkuMap = [];
            foreach ($sheetData->data as $item) {
                $sku = $item->{'(Child) sku'} ?? '';
                if (!empty($sku)) {
                    $sheetSkuMap[$sku] = $item;
                }
            }

            // 3. Prepare SKU list for related models
            $skus = $productMasters->pluck('sku')->unique()->toArray();

            // 4. Fetch related data
            $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
            $nrValues = WayfairDataView::pluck('value', 'sku');
            $jungleScoutData = JungleScoutProductData::all()
                ->groupBy('parent')
                ->map(function ($group) {
                    $validPrices = $group->filter(function ($item) {
                        $price = $item->data['price'] ?? null;
                        return is_numeric($price) && $price > 0;
                    })->pluck('data.price');
                    return [
                        'scout_parent' => $group->first()->parent,
                        'min_price' => $validPrices->isNotEmpty() ? $validPrices->min() : null,
                        'product_count' => $group->count(),
                        'all_data' => $group->map(function ($item) {
                            $data = $item->data;
                            if (isset($data['price'])) {
                                $data['price'] = is_numeric($data['price']) ? (float) $data['price'] : null;
                            }
                            return $data;
                        })->toArray()
                    ];
                });

            // 5. Build the result using ProductMaster as the base
            $processedData = [];
            foreach ($productMasters as $pm) {
                $sku = $pm->sku;
                $parentSku = $pm->parent;

                // Start with ProductMaster base
                $item = [
                    'Parent' => $parentSku,
                    'sku' => $sku,
                ];

                // Merge Google Sheet data if exists
                if (isset($sheetSkuMap[$sku])) {
                    foreach ((array)$sheetSkuMap[$sku] as $key => $val) {
                        $item[$key] = $val;
                    }
                }

                // Add JungleScout data if parent ASIN matches
                if (!empty($parentSku) && $jungleScoutData->has($parentSku)) {
                    $scoutData = $jungleScoutData[$parentSku];
                    $item['scout_data'] = [
                        'scout_parent' => $scoutData['scout_parent'],
                        'min_price' => $scoutData['min_price'],
                        'product_count' => $scoutData['product_count'],
                        'all_data' => $scoutData['all_data']
                    ];
                }

                // Shopify data
                $item['INV'] = $shopifyData->has($sku) ? $shopifyData[$sku]->inv : 0;
                $item['L30'] = $shopifyData->has($sku) ? $shopifyData[$sku]->quantity : 0;

                // NR and A_Z fields
                $item['NR'] = 'REQ';
                $item['A_Z_Reason'] = null;
                $item['A_Z_ActionRequired'] = null;
                $item['A_Z_ActionTaken'] = null;
                if (isset($nrValues[$sku])) {
                    $val = $nrValues[$sku];
                    if (is_array($val)) {
                        $item['NR'] = $val['NR'] ?? 'REQ';
                        $item['A_Z_Reason'] = $val['A_Z_Reason'] ?? null;
                        $item['A_Z_ActionRequired'] = $val['A_Z_ActionRequired'] ?? null;
                        $item['A_Z_ActionTaken'] = $val['A_Z_ActionTaken'] ?? null;
                    } else {
                        $decoded = json_decode($val, true);
                        $item['NR'] = $decoded['NR'] ?? 'REQ';
                        $item['A_Z_Reason'] = $decoded['A_Z_Reason'] ?? null;
                        $item['A_Z_ActionRequired'] = $decoded['A_Z_ActionRequired'] ?? null;
                        $item['A_Z_ActionTaken'] = $decoded['A_Z_ActionTaken'] ?? null;
                    }
                }

                $processedData[] = (object)$item;
            }

            // 6. Filter: INV > 0 and SESS L30 == 0 (or your metric)
            $filteredResults = array_filter($processedData, function ($item) {
                $childSku = $item->{'(Child) sku'} ?? '';
                $inv = $item->INV ?? 0;
                $sessL30 = isset($item->{'Sess30'}) ? (int)$item->{'Sess30'} : 0;
                return
                    stripos($childSku, 'PARENT') === false &&
                    $inv > 0 &&
                    $sessL30 == 0;
            });

            // 7. Return the processed data with filters applied
            return response()->json([
                'message' => 'Data fetched successfully',
                'data' => array_values($filteredResults),
                'status' => 200
            ]);
        } else {
            return response()->json([
                'message' => 'Failed to fetch data from Google Sheet',
                'status' => $response->getStatusCode()
            ], $response->getStatusCode());
        }
    } 

    public function updateReasonAction(Request $request)
    {
        $sku = $request->input('sku');
        $reason = $request->input('reason');
        $actionRequired = $request->input('action_required');
        $actionTaken = $request->input('action_taken');

        if (!$sku) {
            return response()->json([
                'status' => 400,
                'message' => 'SKU is required.'
            ], 400);
        }

        $row = WayfairDataView::firstOrCreate(['sku' => $sku]);
        $value = $row->value ?? [];
        $value['A_Z_Reason'] = $reason;
        $value['A_Z_ActionRequired'] = $actionRequired;
        $value['A_Z_ActionTaken'] = $actionTaken;
        $row->value = $value;
        $row->save();

        return response()->json([
            'status' => 200,
            'message' => 'Reason and actions updated successfully.'
        ]);
    }

    public function getZeroViewCount()
    {
        // 1. Fetch all ProductMaster rows as the base
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        // 2. Fetch Google Sheet data
        $response = $this->apiController->fetchDataFromWayfairMasterGoogleSheet();

        $sheetSkuMap = [];
        if ($response->getStatusCode() === 200) {
            $sheetData = $response->getData();
            foreach ($sheetData->data as $item) {
                $sku = $item->{'(Child) sku'} ?? '';
                if (!empty($sku)) {
                    $sheetSkuMap[$sku] = $item;
                }
            }
        }

        // 3. Prepare SKU list for related models
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        // 4. Fetch related data
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        // 5. Count SKUs where INV > 0 and Sess30 == 0
        $zeroViewCount = 0;
        foreach ($productMasters as $pm) {
            $sku = $pm->sku;
            $inv = $shopifyData->has($sku) ? $shopifyData[$sku]->inv : 0;
            $sess30 = isset($sheetSkuMap[$sku]) ? (int)($sheetSkuMap[$sku]->{'Sess30'} ?? 0) : 0;
            if ($inv > 0 && $sess30 == 0) {
                $zeroViewCount++;
            }
        }

        return $zeroViewCount;
    }
}