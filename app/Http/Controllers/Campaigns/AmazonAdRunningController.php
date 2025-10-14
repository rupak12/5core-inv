<?php

namespace App\Http\Controllers\Campaigns;

use App\Http\Controllers\Controller;
use App\Models\AmazonDatasheet;
use App\Models\AmazonDataView;
use App\Models\AmazonSbCampaignReport;
use App\Models\AmazonSpCampaignReport;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;

class AmazonAdRunningController extends Controller
{
    public function index(){
        return view('campaign.amz-ad-running');
    }

    public function getAmazonAdRunningData()
    {
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();

        $amazonDatasheetsBySku = AmazonDatasheet::whereIn('sku', $skus)->get()->keyBy(function ($item) {
            return strtoupper($item->sku);
        });

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $nrValues = AmazonDataView::whereIn('sku', $skus)->pluck('value', 'sku');

        $amazonKwL30 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->get();

        $amazonKwL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->get();

        $amazonPtL30 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
            })
            ->get();

        $amazonPtL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
            })
            ->get();

        $amazonHlL30 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
            })
            ->get();

        $amazonHlL7 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
            })
            ->get();

        $parentSkuCounts = $productMasters
            ->filter(fn($pm) => $pm->parent && !str_starts_with(strtoupper($pm->sku), 'PARENT'))
            ->groupBy('parent')
            ->map->count();


        $result = [];
        $parentHlSpendData = [];
        
        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = trim($pm->parent);

            $matchedCampaignHlL30 = $amazonHlL30->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                return in_array($cleanName, [$sku, $sku . ' HEAD']) && strtoupper($item->campaignStatus) === 'ENABLED';
            });
            $matchedCampaignHlL7 = $amazonHlL7->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                return in_array($cleanName, [$sku, $sku . ' HEAD']) && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            if (str_starts_with($sku, 'PARENT')) {
                $childCount = $parentSkuCounts[$parent] ?? 0;
                $parentHlSpendData[$parent] = [
                    'total_L30' => $matchedCampaignHlL30->cost ?? 0,
                    'total_L7'  => $matchedCampaignHlL7->cost ?? 0,
                    'total_L30_sales' => $matchedCampaignHlL30->sales ?? 0,
                    'total_L7_sales'  => $matchedCampaignHlL7->sales ?? 0,
                    'total_L30_sold'  => $matchedCampaignHlL30->unitsSold ?? 0,
                    'total_L7_sold'   => $matchedCampaignHlL7->unitsSold ?? 0,
                    'total_L30_impr'  => $matchedCampaignHlL30->impressions ?? 0,
                    'total_L7_impr'   => $matchedCampaignHlL7->impressions ?? 0,
                    'total_L30_clicks'=> $matchedCampaignHlL30->clicks ?? 0,
                    'total_L7_clicks' => $matchedCampaignHlL7->clicks ?? 0,
                    'childCount'=> $childCount,
                ];
            }
        }

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = trim($pm->parent);

            $amazonSheet = $amazonDatasheetsBySku[$sku] ?? null;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignKwL30 = $amazonKwL30->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $matchedCampaignKwL7 = $amazonKwL7->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $matchedCampaignPtL30 = $amazonPtL30->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                return (str_ends_with($cleanName, $sku . ' PT') || str_ends_with($cleanName, $sku . ' PT.'))
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });
            $matchedCampaignPtL7 = $amazonPtL7->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                return (str_ends_with($cleanName, $sku . ' PT') || str_ends_with($cleanName, $sku . ' PT.'))
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            $row = [];
            $row['parent'] = $parent;
            $row['sku'] = $pm->sku;
            $row['INV'] = $shopify->inv ?? 0;
            $row['L30'] = $shopify->quantity ?? 0;
            $row['fba'] = $pm->fba ?? null;
            $row['A_L30'] = $amazonSheet->units_ordered_l30 ?? 0;

            // --- KW ---
            $row['kw_impr_L30'] = $matchedCampaignKwL30->impressions ?? 0;
            $row['kw_impr_L7']  = $matchedCampaignKwL7->impressions ?? 0;
            $row['kw_clicks_L30'] = $matchedCampaignKwL30->clicks ?? 0;
            $row['kw_clicks_L7']  = $matchedCampaignKwL7->clicks ?? 0;
            $row['kw_spend_L30']  = $matchedCampaignKwL30->spend ?? 0;
            $row['kw_spend_L7']   = $matchedCampaignKwL7->spend ?? 0;
            $row['kw_sales_L30']  = $matchedCampaignKwL30->sales30d ?? 0;
            $row['kw_sales_L7']   = $matchedCampaignKwL7->sales7d ?? 0;
            $row['kw_sold_L30']  = $matchedCampaignKwL30->unitsSoldClicks30d ?? 0;
            $row['kw_sold_L7']   = $matchedCampaignKwL7->unitsSoldClicks7d ?? 0;

            // --- PT ---
            $row['pt_impr_L30'] = $matchedCampaignPtL30->impressions ?? 0;
            $row['pt_impr_L7']  = $matchedCampaignPtL7->impressions ?? 0;
            $row['pt_clicks_L30'] = $matchedCampaignPtL30->clicks ?? 0;
            $row['pt_clicks_L7']  = $matchedCampaignPtL7->clicks ?? 0;
            $row['pt_spend_L30']  = $matchedCampaignPtL30->spend ?? 0;
            $row['pt_spend_L7']   = $matchedCampaignPtL7->spend ?? 0;
            $row['pt_sales_L30']  = $matchedCampaignPtL30->sales30d ?? 0;
            $row['pt_sales_L7']   = $matchedCampaignPtL7->sales7d ?? 0;
            $row['pt_sold_L30']  = $matchedCampaignPtL30->unitsSoldClicks30d ?? 0;
            $row['pt_sold_L7']   = $matchedCampaignPtL7->unitsSoldClicks7d ?? 0;

            // --- HL  ---
            $row['hl_impr_L30'] = $matchedCampaignHlL30->impressions ?? 0;
            $row['hl_impr_L7']  = $matchedCampaignHlL7->impressions ?? 0;
            $row['hl_clicks_L30'] = $matchedCampaignHlL30->clicks ?? 0;
            $row['hl_clicks_L7']  = $matchedCampaignHlL7->clicks ?? 0;
            $row['hl_campaign_L30'] = $matchedCampaignHlL30->campaignName ?? null;
            $row['hl_campaign_L7']  = $matchedCampaignHlL7->campaignName ?? null;
            $row['hl_sales_L30']  = 0;
            $row['hl_sales_L7']   = 0;
            $row['hl_sold_L30']  = 0;
            $row['hl_sold_L7']   = 0;

            if (str_starts_with($sku, 'PARENT')) {
                $row['hl_spend_L30'] = $matchedCampaignHlL30->cost ?? 0;
                $row['hl_spend_L7']  = $matchedCampaignHlL7->cost ?? 0;
                $row['hl_sales_L30']  = $matchedCampaignHlL30->sales ?? 0;
                $row['hl_sales_L7']   = $matchedCampaignHlL7->sales ?? 0;
                $row['hl_sold_L30']  = $matchedCampaignHlL30->unitsSold ?? 0;
                $row['hl_sold_L7']   = $matchedCampaignHlL7->unitsSold ?? 0;
                $row['hl_impr_L30'] = $matchedCampaignHlL30->impressions ?? 0;
                $row['hl_impr_L7']  = $matchedCampaignHlL7->impressions ?? 0;
                $row['hl_clicks_L30'] = $matchedCampaignHlL30->clicks ?? 0;
                $row['hl_clicks_L7']  = $matchedCampaignHlL7->clicks ?? 0;
            } 
            elseif (isset($parentHlSpendData[$parent]) && $parentHlSpendData[$parent]['childCount'] > 0) {
                $row['hl_spend_L30'] = $parentHlSpendData[$parent]['total_L30'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_spend_L7']  = $parentHlSpendData[$parent]['total_L7'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_sales_L30']  = $parentHlSpendData[$parent]['total_L30_sales'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_sales_L7']   = $parentHlSpendData[$parent]['total_L7_sales'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_sold_L30']  = $parentHlSpendData[$parent]['total_L30_sold'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_sold_L7']   = $parentHlSpendData[$parent]['total_L7_sold'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_impr_L30'] = $parentHlSpendData[$parent]['total_L30_impr'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_impr_L7']  = $parentHlSpendData[$parent]['total_L7_impr'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_clicks_L30'] = $parentHlSpendData[$parent]['total_L30_clicks'] / $parentHlSpendData[$parent]['childCount'];
                $row['hl_clicks_L7']  = $parentHlSpendData[$parent]['total_L7_clicks'] / $parentHlSpendData[$parent]['childCount'];
            } else {
                $row['hl_spend_L30'] = 0;
                $row['hl_spend_L7']  = 0;
                $row['hl_sales_L30'] = 0;
                $row['hl_sales_L7']  = 0;
                $row['hl_sold_L30']  = 0;
                $row['hl_sold_L7']   = 0;
                $row['hl_impr_L30'] = 0;
                $row['hl_impr_L7']  = 0;
                $row['hl_clicks_L30'] = 0;
                $row['hl_clicks_L7']  = 0;
            }
            
            $childCount = $parentSkuCounts[$parent] ?? 0;
            $childCount = max($childCount, 1);

            $hl_share_L30 = ($matchedCampaignHlL30->impressions ?? 0) / $childCount;
            $hl_share_L7  = ($matchedCampaignHlL7->impressions ?? 0) / $childCount;

            $hl_share_clicks_L30 = ($matchedCampaignHlL30->impressions ?? 0) / $childCount;
            $hl_share_clicks_L7  = ($matchedCampaignHlL7->impressions ?? 0) / $childCount;

            $row['IMP_L30'] = ($row['pt_impr_L30'] + $row['kw_impr_L30'] + $hl_share_L30);
            $row['IMP_L7']  = ($row['pt_impr_L7']  + $row['kw_impr_L7']  + $hl_share_L7);

            $row['CLICKS_L30'] = ($row['pt_clicks_L30'] + $row['kw_clicks_L30'] + $hl_share_clicks_L30);
            $row['CLICKS_L7']  = ($row['pt_clicks_L7']  + $row['kw_clicks_L7']  + $hl_share_clicks_L7);

            $row['SPEND_L30'] = $row['pt_spend_L30'] + $row['kw_spend_L30'] + $row['hl_spend_L30'];
            $row['SPEND_L7']  = $row['pt_spend_L7'] + $row['kw_spend_L7'] + $row['hl_spend_L7'];

            $row['SALES_L30'] = $row['pt_sales_L30'] + $row['kw_sales_L30'] + $row['hl_sales_L30'];
            $row['SALES_L7']  = $row['pt_sales_L7'] + $row['kw_sales_L7'] + $row['hl_sales_L7'];

            $row['SOLD_L30'] = $row['pt_sold_L30'] + $row['kw_sold_L30'] + $row['hl_sold_L30'];
            $row['SOLD_L7']  = $row['pt_sold_L7'] + $row['kw_sold_L7'] + $row['hl_sold_L7'];

            $row['NRL'] = '';
            $row['NRA'] = '';
            $row['FBA'] = '';
            $row['start_ad'] = '';
            $row['stop_ad'] = '';

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) $raw = json_decode($raw, true);
                if (is_array($raw)) {
                    $row['NRL'] = $raw['NRL'] ?? null;
                    $row['NRA'] = $raw['NRA'] ?? null;
                    $row['FBA'] = $raw['FBA'] ?? null;
                    $row['start_ad'] = $raw['start_ad'] ?? null;
                    $row['stop_ad'] = $raw['stop_ad'] ?? null;
                }
            }

            $result[] = $row;
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

}
