<?php

namespace App\Http\Controllers\Campaigns;

use App\Http\Controllers\Controller;
use App\Models\Product;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\WalmartCampaignReport;
use App\Models\WalmartDataView;
use App\Models\WalmartProductSheet;
use Illuminate\Http\Request;

class WalmartMissingAdsController extends Controller
{
    public function index()
    {
        return view('campaign.walmart-missing-ads');
    }

    public function getWalmartMissingAdsData()
    {
        $normalizeSku = fn($sku) => strtoupper(trim(preg_replace('/\s+/', ' ', str_replace("\xc2\xa0", ' ', $sku))));

        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')
            ->filter()
            ->unique()
            ->map(fn($sku) => $normalizeSku($sku))
            ->values()
            ->all();

        $walmartProductSheet = WalmartProductSheet::whereIn('sku', $skus)->get()->keyBy(fn($item) => $normalizeSku($item->sku));

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(fn($item) => $normalizeSku($item->sku));

        $nrValues = WalmartDataView::whereIn('sku', $skus)->pluck('value', 'sku');

        $walmartCampaignReportsAll = WalmartCampaignReport::whereIn('campaignName', $skus)->get()->keyBy(fn($item) => $normalizeSku($item->campaignName));

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = $normalizeSku($pm->sku);
            $parent = $pm->parent;

            $walmartSheet = $walmartProductSheet[$sku] ?? null;
            $shopify = $shopifyData[$sku] ?? null;

            $matchedCampaign = $walmartCampaignReportsAll[$sku] ?? null;

            $row = [];
            $row['parent'] = $parent;
            $row['sku']    = $pm->sku;
            $row['INV']    = $shopify->inv ?? 0;
            $row['L30']    = $shopify->quantity ?? 0;
            $row['WA_L30'] = $walmartSheet->l30 ?? 0;

            $row['campaignName'] = $matchedCampaign->campaignName ?? '';
            $row['campaignBudgetAmount'] = $matchedCampaign->budget ?? '';
            $row['campaignStatus'] = $matchedCampaign->status ?? '';

            $row['NRA']  = '';
            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) $raw = json_decode($raw, true);
                if (is_array($raw)) {
                    $row['NRA'] = $raw['NR'] ?? null;
                }
            }

            if($row['NRA'] != 'NRA' ){
                $result[] = (object) $row;
            }
        }

        $uniqueResult = collect($result)->unique('sku')->values()->all();

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $uniqueResult,
            'status'  => 200,
        ]);
    }
}
