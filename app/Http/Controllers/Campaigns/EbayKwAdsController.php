<?php

namespace App\Http\Controllers\Campaigns;

use App\Http\Controllers\Controller;
use App\Models\EbayDataView;
use App\Models\EbayMetric;
use App\Models\EbayPriorityReport;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Support\Facades\Log;

class EbayKwAdsController extends Controller
{
    public function index(){
        return view('campaign.ebay-kw-ads');
    }

    public function getEbayKwAdsData(){

        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        $nrValues = EbayDataView::whereIn('sku', $skus)->pluck('value', 'sku');

        $periods = ['L7', 'L15', 'L30', 'L60'];
        $campaignReports = [];
        foreach ($periods as $period) {
            $campaignReports[$period] = EbayPriorityReport::where('report_range', $period)
                ->where(function ($q) use ($skus) {
                    foreach ($skus as $sku) {
                        $q->orWhere('campaign_name', 'LIKE', '%' . $sku . '%');
                    }
                })->get();
        }

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $row = [
                'parent' => $parent,
                'sku'    => $pm->sku,
                'INV'    => $shopify->inv ?? 0,
                'L30'    => $shopify->quantity ?? 0,
                'NR'     => ''
            ];

            $matchedCampaignL30 = $campaignReports['L30']->first(function ($item) use ($sku) {
                return strtoupper(trim($item->campaign_name)) === strtoupper(trim($sku));
            });

            $row['campaignName'] = $matchedCampaignL30->campaign_name ?? '';
            $row['campaignBudgetAmount'] = $matchedCampaignL30->campaignBudgetAmount ?? 0;
            $row['campaignStatus'] = $matchedCampaignL30->campaignStatus ?? '';

            if(!$matchedCampaignL30){
                continue;
            }

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NR'] = $raw['NR'] ?? null;
                }
            }

            foreach ($periods as $period) {
                $matchedCampaign = $campaignReports[$period]->first(function ($item) use ($sku) {
                    return strtoupper(trim($item->campaign_name)) === strtoupper(trim($sku));
                });
                
                if (!$matchedCampaign) {
                    $row["impressions_" . strtolower($period)] = 0;
                    $row["clicks_" . strtolower($period)]      = 0;
                    $row["ad_sales_" . strtolower($period)]    = 0;
                    $row["ad_sold_" . strtolower($period)]     = 0;
                    $row["spend_" . strtolower($period)]       = 0;
                    $row["acos_" . strtolower($period)]        = 0;
                    $row["cpc_" . strtolower($period)]         = 0;
                    continue;
                }

                $adFees = (float) str_replace('USD ', '', $matchedCampaign->cpc_ad_fees_payout_currency ?? 0);
                $sales  = (float) str_replace('USD ', '', $matchedCampaign->cpc_sale_amount_payout_currency ?? 0);
                $clicks = (float) ($matchedCampaign->cpc_clicks ?? 0);
                $spend  = (float) ($matchedCampaign->cpc_cost ?? $adFees);
                $cpc    = $clicks > 0 ? ($spend / $clicks) : 0;
                $acos   = $sales > 0 ? ($adFees / $sales) * 100 : 0;

                if ($adFees > 0 && $sales === 0) {
                    $acos = 100;
                }

                $row["impressions_" . strtolower($period)] = $matchedCampaign->cpc_impressions ?? 0;
                $row["clicks_" . strtolower($period)]      = $matchedCampaign->cpc_clicks ?? 0;
                $row["ad_sales_" . strtolower($period)]    = $sales;
                $row["ad_sold_" . strtolower($period)]     = $matchedCampaign->unitsSold ?? 0;
                $row["spend_" . strtolower($period)]       = $adFees;
                $row["acos_" . strtolower($period)]        = $acos;
                $row["cpc_" . strtolower($period)]         = $cpc;
            }

            if ($row['NR'] !== "NRA") {
                $result[] = (object) $row;
            }
    
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    public function ebayPriceLessThanTwentyAdsView(){
        return view('campaign.ebay-less-twenty-kw-ads');
    }

    public function ebayPriceLessThanTwentyAdsData()
    {
        $normalizeSku = fn($sku) => strtoupper(trim($sku));

        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->map($normalizeSku)->unique()->values()->all();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(fn($item) => $normalizeSku($item->sku));

        $ebayMetricData = EbayMetric::whereIn('sku', $skus)->get()->keyBy(fn($item) => $normalizeSku($item->sku));

        $nrValues = EbayDataView::whereIn('sku', $skus)->pluck('value', 'sku');

        $ebayCampaignReportsL7 = EbayPriorityReport::where('report_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaign_name', 'LIKE', '%' . $sku . '%');
                }
            })
            ->get();

        $ebayCampaignReportsL1 = EbayPriorityReport::where('report_range', 'L1')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaign_name', 'LIKE', '%' . $sku . '%');
                }
            })
            ->get();

        $ebayCampaignReportsL30 = EbayPriorityReport::where('report_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaign_name', 'LIKE', '%' . $sku . '%');
                }
            })
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $shopify = $shopifyData[$sku] ?? null;

            $ebay = $ebayMetricData[$sku] ?? null;

            $matchedCampaignL7 = $ebayCampaignReportsL7->first(function ($item) use ($sku) {
                return strtoupper(trim($item->campaign_name)) === strtoupper(trim($sku));
            });

            $matchedCampaignL1 = $ebayCampaignReportsL1->first(function ($item) use ($sku) {
                return strtoupper(trim($item->campaign_name)) === strtoupper(trim($sku));
            });

            $matchedCampaignL30 = $ebayCampaignReportsL30->first(function ($item) use ($sku) {
                return strtoupper(trim($item->campaign_name)) === strtoupper(trim($sku));
            });

            $row = [];
            $row['parent'] = $parent;
            $row['sku']    = $pm->sku;
            $row['INV']    = $shopify->inv ?? 0;
            $row['L30']    = $shopify->quantity ?? 0;
            $row['e_l30']  = $ebay->ebay_l30 ?? 0;
            $row['price']  = $ebay->ebay_data_price ?? 0;
            $row['campaign_id'] = $matchedCampaignL7->campaign_id ?? ($matchedCampaignL1->campaign_id ?? '');
            $row['campaignName'] = $matchedCampaignL7->campaign_name ?? ($matchedCampaignL1->campaign_name ?? '');
            $row['campaignStatus'] = $matchedCampaignL7->campaignStatus ?? ($matchedCampaignL1->campaignStatus ?? '');
            $row['campaignBudgetAmount'] = $matchedCampaignL7->campaignBudgetAmount ?? ($matchedCampaignL1->campaignBudgetAmount ?? '');

            $adFees   = (float) str_replace('USD ', '', $matchedCampaignL30->cpc_ad_fees_payout_currency ?? 0);
            $sales    = (float) ($matchedCampaignL30 ? ($matchedCampaignL30->cpc_attributed_sales ?? 0) : 0);

            $acos = $sales > 0 ? ($adFees / $sales) * 100 : 0;
            
            if($adFees > 0 && $sales === 0){
                $row['acos'] = 100;
            }else{
                $row['acos'] = $acos;
            }

            $row['l7_spend'] = (float) str_replace('USD ', '', $matchedCampaignL7->cpc_ad_fees_payout_currency ?? 0);
            $row['l7_cpc'] = (float) str_replace('USD ', '', $matchedCampaignL7->cost_per_click ?? 0);
            $row['l1_spend'] = (float) str_replace('USD ', '', $matchedCampaignL1->cpc_ad_fees_payout_currency ?? 0);
            $row['l1_cpc'] = (float) str_replace('USD ', '', $matchedCampaignL1->cost_per_click ?? 0);
            $row['sbid'] = 0;

            $row['NR'] = '';
            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NR'] = $raw['NR'] ?? '';
                }
            }

            $budget = floatval($row['campaignBudgetAmount']);
            $l7_spend = floatval($row['l7_spend']);
            $l1_cpc = floatval($row['l1_cpc']);
            $l7_cpc = floatval($row['l7_cpc']);

            $ub7 = $budget > 0 ? ($l7_spend / ($budget * 7)) * 100 : 0;
            if($ub7 < 70){
                if($l1_cpc > $l7_cpc){
                    $row['sbid'] = floor($l1_cpc * 1.05 * 100) / 100;
                }else{
                    $row['sbid'] = floor($l7_cpc * 1.05 * 100) / 100;
                }
            }else if($ub7 > 90){
                $row['sbid'] = floor($l1_cpc * 0.90 * 100) / 100;
            }
            
            if($row['price'] < 30 && $row['campaignName'] !== ''){
                if($row['price'] <= 10 && $row['sbid'] > 0.10){
                    $row['sbid'] = 0.10;
                }
                elseif($row['price'] > 10 && $row['price'] <= 20 && $row['sbid'] > 0.20){
                    $row['sbid'] = 0.20;
                }
                elseif($row['price'] > 20 && $row['price'] <= 30 && $row['sbid'] > 0.30){
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
