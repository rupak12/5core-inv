<?php

namespace App\Http\Controllers\MarketPlace;

use App\Http\Controllers\ApiController;
use App\Http\Controllers\Controller;
use App\Models\MacyDataView;
use App\Models\MacyProduct;
use App\Models\MacysListingStatus;
use App\Models\MarketplacePercentage;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Cache;


class MacyZeroController extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }
    public function macyZeroView(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        // Get percentage from cache or database
        $percentage = Cache::remember('macys_marketplace_percentage', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'Macys')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100; // Default to 100 if not set
        });

        return view('market-places.macysZeroView', [
            'mode' => $mode,
            'demo' => $demo,
            'macysPercentage' => $percentage
        ]);
    }
    public function getViewMacyZeroData(Request $request)
    {
        // Fetch all ProductMaster records
        $productMasters = ProductMaster::all();
        $skus = $productMasters->pluck('sku')->toArray();

        // Fetch ShopifySku and MacyProduct records for those SKUs
        $shopifySkus = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $macyProducts = MacyProduct::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch MacyDataView NR values for those SKUs
        $macyDataViews = MacyDataView::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch data from the Google Sheet using the ApiController method
        $sheetResponse = $this->apiController->fetchMacyListingData();
        $sheetData = [];
        if ($sheetResponse->getStatusCode() === 200) {
            $sheetRaw = $sheetResponse->getData();
            $sheetData = $sheetRaw->data ?? [];
        }

        // Map ProductMaster SKUs to sheet data (by (Child) sku)
        $sheetSkuMap = collect($sheetData)
            ->filter(function ($item) {
                return !empty($item->{'(Child) sku'} ?? null);
            })
            ->keyBy(function ($item) {
                return $item->{'(Child) sku'};
            });

        // Attach matching ShopifySku, MacyProduct, Sheet data, and NR to each ProductMaster record
        $processedData = $productMasters->map(function ($product) use ($shopifySkus, $macyProducts, $sheetSkuMap, $macyDataViews) {
            $sku = $product->sku;
            $product->INV = $shopifySkus->has($sku) ? $shopifySkus[$sku]->inv : 0;
            $product->L30 = $shopifySkus->has($sku) ? $shopifySkus[$sku]->quantity : 0;
            $product->m_l30 = $macyProducts->has($sku) ? $macyProducts[$sku]->m_l30 : null;
            $product->m_l60 = $macyProducts->has($sku) ? $macyProducts[$sku]->m_l60 : null;
            $product->price = $macyProducts->has($sku) ? $macyProducts[$sku]->price : null;
            $product->sheet_data = $sheetSkuMap->has($sku) ? $sheetSkuMap[$sku] : null;

            // Fetch NR and A_Z fields from MacyDataView
            $nrValue = null;
            $a_z_reason = null;
            $a_z_action_required = null;
            $a_z_action_taken = null;
            if ($macyDataViews->has($sku)) {
                $value = $macyDataViews[$sku]->value;
                if (is_array($value)) {
                    $nrValue = $value['NR'] ?? 'REQ';
                    $a_z_reason = $value['A_Z_Reason'] ?? null;
                    $a_z_action_required = $value['A_Z_ActionRequired'] ?? null;
                    $a_z_action_taken = $value['A_Z_ActionTaken'] ?? null;
                } else {
                    $decoded = json_decode($value, true);
                    $nrValue = $decoded['NR'] ?? 'REQ';
                    $a_z_reason = $decoded['A_Z_Reason'] ?? null;
                    $a_z_action_required = $decoded['A_Z_ActionRequired'] ?? null;
                    $a_z_action_taken = $decoded['A_Z_ActionTaken'] ?? null;
                }
            }
            $product->NR = $nrValue;
            $product->A_Z_Reason = $a_z_reason;
            $product->A_Z_ActionRequired = $a_z_action_required;
            $product->A_Z_ActionTaken = $a_z_action_taken;
            return $product;
        })
        // --- Apply filter: only rows where INV > 0 and VIEWS == 0 ---
        ->filter(function ($product) {
            $inv = (int)($product->INV ?? 0);
            $views = 0;
            if ($product->sheet_data && isset($product->sheet_data->VIEWS)) {
                $views = (int)$product->sheet_data->VIEWS;
            }
            return $inv > 0 && $views == 0;
        })
        ->values();

        return response()->json([
            'message' => 'Data fetched successfully',
            'product_master_data' => $processedData,
            'sheet_data' => $sheetData,
            'status' => 200
        ]);
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

        $row = MacyDataView::firstOrCreate(['sku' => $sku]);
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
        // Fetch all ProductMaster records
        $productMasters = ProductMaster::all();
        $skus = $productMasters->pluck('sku')->toArray();

        // Fetch ShopifySku and MacyProduct records for those SKUs
        $shopifySkus = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $macyProducts = MacyProduct::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch data from the Google Sheet using the ApiController method
        $sheetResponse = $this->apiController->fetchMacyListingData();
        $sheetData = [];
        if ($sheetResponse->getStatusCode() === 200) {
            $sheetRaw = $sheetResponse->getData();
            $sheetData = $sheetRaw->data ?? [];
        }

        // Map ProductMaster SKUs to sheet data (by (Child) sku)
        $sheetSkuMap = collect($sheetData)
            ->filter(function ($item) {
                return !empty($item->{'(Child) sku'} ?? null);
            })
            ->keyBy(function ($item) {
                return $item->{'(Child) sku'};
            });

        // Count SKUs where INV > 0 and VIEWS == 0
        $zeroViewCount = $productMasters->filter(function ($product) use ($shopifySkus, $sheetSkuMap) {
            $sku = $product->sku;
            $inv = $shopifySkus->has($sku) ? $shopifySkus[$sku]->inv : 0;
            $views = 0;
            if ($sheetSkuMap->has($sku) && isset($sheetSkuMap[$sku]->VIEWS)) {
                $views = (int) $sheetSkuMap[$sku]->VIEWS;
            }
            return $inv > 0 && $views == 0;
        })->count();

        return $zeroViewCount;
    }

     public function getLivePendingAndZeroViewCounts()
    {
        $productMasters = ProductMaster::whereNull('deleted_at')->get();
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $macyDataViews = MacysListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');

        $macyMetrics = MacyProduct::whereIn('sku', $skus)->get()->keyBy('sku');


        $listedCount = 0;
        $zeroInvOfListed = 0;
        $liveCount = 0;
        $zeroViewCount = 0;

        foreach ($productMasters as $item) {
            $sku = trim($item->sku);
            $inv = $shopifyData[$sku]->inv ?? 0;
            $isParent = stripos($sku, 'PARENT') !== false;
            if ($isParent) continue;

            $status = $macyDataViews[$sku]->value ?? null;
            if (is_string($status)) {
                $status = json_decode($status, true);
            }
            $listed = $status['listed'] ?? (floatval($inv) > 0 ? 'Pending' : 'Listed');
            $live = $status['live'] ?? null;

            // Listed count (for live pending)
            if ($listed === 'Listed') {
                $listedCount++;
                if (floatval($inv) <= 0) {
                    $zeroInvOfListed++;
                }
            }

            // Live count
            if ($live === 'Live') {
                $liveCount++;
            }

            // Zero view: INV > 0, views == 0 (from ebay_metric table), not parent SKU (NR ignored)
            $views = $macyMetrics[$sku]->views_clicks ?? null;
            // if (floatval($inv) > 0 && $views !== null && intval($views) === 0) {
            //     $zeroViewCount++;
            // }
            if ($inv > 0) {
                if ($views === null) {
                    // Do nothing, ignore null
                } elseif (intval($views) === 0) {
                    $zeroViewCount++;
                }
            }
        }

        // live pending = listed - 0-inv of listed - live
        $livePending = $listedCount - $zeroInvOfListed - $liveCount;

        return [
            'live_pending' => $livePending,
            'zero_view' => $zeroViewCount,
        ];
    }
}