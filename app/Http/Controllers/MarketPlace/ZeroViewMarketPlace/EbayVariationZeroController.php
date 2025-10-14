<?php

namespace App\Http\Controllers\MarketPlace\ZeroViewMarketPlace;

use App\Http\Controllers\Controller;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\EbayVariationDataView;
use App\Models\EbayVariationListingStatus;
use Illuminate\Http\Request;

class EbayVariationZeroController extends Controller
{
    public function ebayVariationZeroview(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');
        return view('market-places.zero-market-places.ebayVariationZeroView', [
            'mode' => $mode,
            'demo' => $demo
        ]);
    }

    public function getViewEbayVariationZeroData(Request $request)
    {
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $variationDataViews = EbayVariationDataView::whereIn('sku', $skus)->get()->keyBy('sku');

        $result = [];
        foreach ($productMasters as $pm) {
            $sku = $pm->sku;
            $parent = $pm->parent;
            $shopify = $shopifyData[$sku] ?? null;

            $inv = $shopify ? $shopify->inv : 0;
            $ov_l30 = $shopify ? $shopify->quantity : 0;
            $ov_dil = ($inv > 0) ? round($ov_l30 / $inv, 4) : 0;

            if ($inv > 0) {
                $variationView = $variationDataViews[$sku] ?? null;
                $value = $variationView ? $variationView->value : [];
                if (is_string($value)) {
                    $value = json_decode($value, true) ?: [];
                }

                $row = [
                    'parent' => $parent,
                    'sku' => $sku,
                    'inv' => $inv,
                    'ov_l30' => $ov_l30,
                    'ov_dil' => $ov_dil,
                    'NR' => isset($value['NR']) && in_array($value['NR'], ['REQ', 'NR']) ? $value['NR'] : 'REQ',
                    'A_Z_Reason' => $value['A_Z_Reason'] ?? '',
                    'A_Z_ActionRequired' => $value['A_Z_ActionRequired'] ?? '',
                    'A_Z_ActionTaken' => $value['A_Z_ActionTaken'] ?? '',
                ];
                $result[] = $row;
            }
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $result,
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

        $row = EbayVariationDataView::firstOrCreate(
            ['sku' => $sku],
            ['value' => json_encode([])]
        );

        $value = $row->value;
        if (is_string($value)) {
            $value = json_decode($value, true) ?: [];
        }

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

    public function getLivePendingAndZeroViewCounts()
    {
        $productMasters = ProductMaster::whereNull('deleted_at')->get();
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayDataViews = EbayVariationListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');

        $listedCount = 0;
        $zeroInvOfListed = 0;
        $liveCount = 0;
        $zeroViewCount = 0; // eBay Variation doesn't have metrics/views data

        foreach ($productMasters as $item) {
            $sku = trim($item->sku);
            $inv = $shopifyData[$sku]->inv ?? 0;
            $isParent = stripos($sku, 'PARENT') !== false;
            if ($isParent) continue;

            $status = $ebayDataViews[$sku]->value ?? null;
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

            // Zero view: eBay Variation doesn't track views, so zero_view remains 0
        }

        // live pending = listed - 0-inv of listed - live
        $livePending = $listedCount - $zeroInvOfListed - $liveCount;

        return [
            'live_pending' => $livePending,
            'zero_view' => $zeroViewCount,
        ];
    }
}
