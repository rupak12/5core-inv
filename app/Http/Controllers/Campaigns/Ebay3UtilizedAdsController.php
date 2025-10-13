<?php

namespace App\Http\Controllers\Campaigns;

use App\Http\Controllers\Controller;
use App\Models\Ebay3Metric;
use App\Models\Ebay3PriorityReport;
use App\Models\EbayThreeDataView;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Http\Request;

class Ebay3UtilizedAdsController extends Controller
{
    public function ebay3OverUtilizedAdsView()
    {
        return view('campaign.ebay-three.over-utilized-ads');
    }

    public function ebay3UnderUtilizedAdsView()
    {
        return view('campaign.ebay-three.under-utilized-ads');
    }

    public function ebay3CorrectlyUtilizedAdsView()
    {
        return view('campaign.ebay-three.correctly-utilized-ads');
    }

    public function getEbay3UtilizedAdsData()
    {
        $normalize = fn($value) => strtoupper(trim($value));

        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayMetricData = Ebay3Metric::whereIn('sku', $skus)->get()->keyBy('sku');
        $nrValues = EbayThreeDataView::whereIn('sku', $skus)->pluck('value', 'sku');

        $allCampaigns = Ebay3PriorityReport::whereIn('campaign_name', $skus)->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = $normalize($pm->sku);
            $parent = $pm->parent;

            $shopify = $shopifyData[$pm->sku] ?? null;
            $ebay = $ebayMetricData[$pm->sku] ?? null;

            $matchedCampaigns = $allCampaigns->filter(function ($c) use ($sku, $normalize) {
                return $normalize($c->campaign_name) === $sku;
            });

            if ($matchedCampaigns->isEmpty()) {
                continue;
            }

            $matchedL1  = $matchedCampaigns->firstWhere('report_range', 'L1');
            $matchedL7  = $matchedCampaigns->firstWhere('report_range', 'L7');
            $matchedL30 = $matchedCampaigns->firstWhere('report_range', 'L30');

            $row = [
                'parent' => $parent,
                'sku'    => $pm->sku,
                'INV'    => $shopify->inv ?? 0,
                'L30'    => $shopify->quantity ?? 0,
                'price'  => $ebay->ebay_price ?? 0,
                'campaign_id' => $matchedL7->campaign_id ?? ($matchedL1->campaign_id ?? ''),
                'campaignName' => $matchedL7->campaign_name ?? ($matchedL1->campaign_name ?? ''),
                'campaignStatus' => $matchedL7->campaignStatus ?? ($matchedL1->campaignStatus ?? ''),
                'campaignBudgetAmount' => $matchedL7->campaignBudgetAmount ?? ($matchedL1->campaignBudgetAmount ?? ''),
                'l7_spend' => (float) str_replace('USD ', '', $matchedL7->cpc_ad_fees_payout_currency ?? 0),
                'l7_cpc'   => (float) str_replace('USD ', '', $matchedL7->cost_per_click ?? 0),
                'l1_spend' => (float) str_replace('USD ', '', $matchedL1->cpc_ad_fees_payout_currency ?? 0),
                'l1_cpc'   => (float) str_replace('USD ', '', $matchedL1->cost_per_click ?? 0),
                'sbid'     => 0.10,
                'NR'       => '',
            ];

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) $raw = json_decode($raw, true);
                if (is_array($raw)) $row['NR'] = $raw['NR'] ?? null;
            }

            $adFees = (float) str_replace('USD ', '', $matchedL30->cpc_ad_fees_payout_currency ?? 0);
            $sales  = (float) str_replace('USD ', '', $matchedL30->cpc_sale_amount_payout_currency ?? 0);
            $row['acos'] = ($adFees > 0 && $sales === 0) ? 100 : ($sales > 0 ? ($adFees / $sales) * 100 : 0);

            if ($row['price'] < 30) {
                if ($row['price'] < 10) {
                    $row['sbid'] = 0.10;
                } elseif ($row['price'] <= 20) {
                    $row['sbid'] = 0.20;
                } else {
                    $row['sbid'] = 0.30;
                }
                $result[] = (object) $row;
            }
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

}
