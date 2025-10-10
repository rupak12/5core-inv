<?php

namespace App\Http\Controllers\MarketPlace;

use App\Http\Controllers\Controller;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\TemuDataView;
use Illuminate\Http\Request;
use App\Http\Controllers\ApiController;
use App\Models\MarketplacePercentage;
use App\Models\TemuListingStatus;
use App\Models\TemuMetric;
use Illuminate\Support\Facades\Cache;


class TemuZeroController extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }
    public function temuZeroView(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        // Get percentage from cache or database
        $percentage = Cache::remember('temu_marketplace_percentage', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'Temu')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100;
        });

        return view('market-places.temuZeroView', [
            'mode' => $mode,
            'demo' => $demo,
            'percentage' => $percentage
        ]);
    }

    public function getViewTemuZeroData(Request $request)
    {
        // Get percentage from cache or database
        $percentage = Cache::remember('temu_marketplace_percentage', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'Temu')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100;
        });
        $percentageValue = $percentage / 100;

        // Fetch all product master records
        $productMasterRows = ProductMaster::all()->keyBy('sku');

        // Get all unique SKUs from product master
        $skus = $productMasterRows->pluck('sku')->toArray();

        // Fetch shopify data for these SKUs
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch NR values for these SKUs from TemuDataView
        $temuDataViews = TemuDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $extraValues = [];
        foreach ($temuDataViews as $sku => $dataView) {
            $value = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
            $extraValues[$sku] = [
                'NR' => $value['NR'] ?? 'REQ',
                'A_Z_Reason' => $value['A_Z_Reason'] ?? null,
                'A_Z_ActionRequired' => $value['A_Z_ActionRequired'] ?? null,
                'A_Z_ActionTaken' => $value['A_Z_ActionTaken'] ?? null,
            ];
        }

        // Process data from product master and shopify tables
        $processedData = [];
        $slNo = 1;

        foreach ($productMasterRows as $productMaster) {
            $sku = $productMaster->sku;
            $isParent = stripos($sku, 'PARENT') !== false;

            // Initialize the data structure
            $processedItem = [
                'SL No.' => $slNo++,
                'Parent' => $productMaster->parent ?? null,
                'Sku' => $sku,
                'R&A' => false, // Default value, can be updated as needed
                'is_parent' => $isParent,
                'raw_data' => [
                    'parent' => $productMaster->parent,
                    'sku' => $sku,
                    'Values' => $productMaster->Values
                ]
            ];

            // Add values from product_master
            $values = $productMaster->Values ?: [];
            $processedItem['LP'] = $values['lp'] ?? 0;
            $processedItem['Ship'] = $values['ship'] ?? 0;
            $processedItem['COGS'] = $values['cogs'] ?? 0;

            // Add data from shopify_skus if available
            if (isset($shopifyData[$sku])) {
                $shopifyItem = $shopifyData[$sku];
                $processedItem['INV'] = $shopifyItem->inv ?? 0;
                $processedItem['L30'] = $shopifyItem->quantity ?? 0;
            } else {
                $processedItem['INV'] = 0;
                $processedItem['L30'] = 0;
            }

            // Fetch NR and A_Z fields if available
            $processedItem['NR'] = $extraValues[$sku]['NR'] ?? 'REQ';
            $processedItem['A_Z_Reason'] = $extraValues[$sku]['A_Z_Reason'] ?? null;
            $processedItem['A_Z_ActionRequired'] = $extraValues[$sku]['A_Z_ActionRequired'] ?? null;
            $processedItem['A_Z_ActionTaken'] = $extraValues[$sku]['A_Z_ActionTaken'] ?? null;

            // Default values for other fields
            $processedItem['A L30'] = 0;
            $processedItem['Sess30'] = 0;
            $processedItem['price'] = 0;
            $processedItem['TOTAL PFT'] = 0;
            $processedItem['T Sales L30'] = 0;
            $processedItem['PFT %'] = 0;
            $processedItem['Roi'] = 0;
            $processedItem['percentage'] = $percentageValue;

            $processedData[] = $processedItem;
        }

        $processedData = array_filter($processedData, function ($item) {
            return isset($item['INV']) && $item['INV'] > 0;
        });
        $processedData = array_values($processedData); // Re-index array

        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $processedData,
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

        $row = TemuDataView::firstOrCreate(['sku' => $sku]);
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

        // Fetch ShopifySku records for those SKUs
        $shopifySkus = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        // Only count SKUs where INV > 0 (zero view logic can be adjusted as needed)
        $zeroViewCount = $productMasters->filter(function ($product) use ($shopifySkus) {
            $sku = $product->sku;
            $inv = $shopifySkus->has($sku) ? $shopifySkus[$sku]->inv : 0;
            // If you have a "views" or "Sess30" field, add the check here
            return $inv > 0;
        })->count();

        return $zeroViewCount;
    }

    public function getLivePendingAndZeroViewCounts()
    {
        $productMasters = ProductMaster::whereNull('deleted_at')->get();
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $temuDataViews = TemuListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');

        $temuMetrics = TemuMetric::whereIn('sku', $skus)->get()->keyBy('sku');


        $listedCount = 0;
        $zeroInvOfListed = 0;
        $liveCount = 0;
        $zeroViewCount = 0;

        foreach ($productMasters as $item) {
            $sku = trim($item->sku);
            $inv = $shopifyData[$sku]->inv ?? 0;
            $isParent = stripos($sku, 'PARENT') !== false;
            if ($isParent) continue;

            $status = $temuDataViews[$sku]->value ?? null;
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
            $views = $temuMetrics[$sku]->product_impressions_l30 ?? null;
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
