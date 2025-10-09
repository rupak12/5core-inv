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
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;

class AmazonCampaignReportsController extends Controller
{

    public function index(){
        $data = DB::table('amazon_sp_campaign_reports')
            ->selectRaw('
                report_date_range,
                SUM(clicks) as clicks, 
                SUM(spend) as spend, 
                SUM(purchases1d) as orders, 
                SUM(sales1d) as sales
            ')
            ->whereIn('report_date_range', ['L60','L30','L15','L7','L1'])
            ->groupBy('report_date_range')
            ->orderByRaw("FIELD(report_date_range, 'L60','L30','L15','L7','L1')")
            ->get();

        $dates  = $data->pluck('report_date_range');
        $clicks = $data->pluck('clicks')->map(fn($v) => (int) $v);
        $spend  = $data->pluck('spend')->map(fn($v) => (float) $v);
        $orders = $data->pluck('orders')->map(fn($v) => (int) $v);
        $sales  = $data->pluck('sales')->map(fn($v) => (float) $v);

        
        return view('campaign.amazon-campaign-reports',compact('dates', 'clicks', 'spend', 'orders', 'sales'));
    }

    public function amazonKwAdsView() {
        $thirtyDaysAgo = \Carbon\Carbon::now()->subDays(30)->format('Y-m-d');

        $data = DB::connection('apicentral')->table('amazon_sp_campaign_reports')
            ->selectRaw('
                report_date_range,
                SUM(clicks) as clicks, 
                SUM(spend) as spend, 
                SUM(purchases1d) as orders, 
                SUM(sales1d) as sales
            ')
            ->whereDate('report_date_range', '>=', $thirtyDaysAgo)
            ->where(function($query) {
                $query->whereRaw("campaignName NOT LIKE '%PT'")
                    ->whereRaw("campaignName NOT LIKE '%PT.'");
            })
            ->groupBy('report_date_range')
            ->orderBy('report_date_range', 'asc')
            ->get();

        $dates  = $data->pluck('report_date_range')->map(fn($d) => \Carbon\Carbon::parse($d)->format('Y-m-d'));
        $clicks = $data->pluck('clicks')->map(fn($v) => (int) $v);
        $spend  = $data->pluck('spend')->map(fn($v) => (float) $v);
        $orders = $data->pluck('orders')->map(fn($v) => (int) $v);
        $sales  = $data->pluck('sales')->map(fn($v) => (float) $v);
        
        return view('campaign.amazon-kw-ads', compact('dates', 'clicks', 'spend', 'orders', 'sales'));
    }


    public function filterKwAds(Request $request)
    {
        $start = $request->startDate;
        $end   = $request->endDate;  

        $data = DB::connection('apicentral')->table('amazon_sp_campaign_reports')
            ->selectRaw('
                report_date_range,
                SUM(clicks) as clicks,
                SUM(spend) as spend,
                SUM(purchases1d) as orders,
                SUM(sales1d) as sales
            ')
            ->whereBetween('report_date_range', [$start, $end])
            ->where(function($query) {
                $query->whereRaw("campaignName NOT LIKE '%PT'")
                    ->whereRaw("campaignName NOT LIKE '%PT.'");
            })
            ->groupBy('report_date_range')
            ->orderBy('report_date_range', 'asc')
            ->get();

        return response()->json([
            'dates'  => $data->pluck('report_date_range'),
            'clicks' => $data->pluck('clicks')->map(fn($v) => (int) $v),
            'spend'  => $data->pluck('spend')->map(fn($v) => (float) $v),
            'orders' => $data->pluck('orders')->map(fn($v) => (int) $v),
            'sales'  => $data->pluck('sales')->map(fn($v) => (float) $v),
            'totals' => [
                'clicks' => $data->sum('clicks'),
                'spend'  => $data->sum('spend'),
                'orders' => $data->sum('orders'),
                'sales'  => $data->sum('sales'),
            ]
        ]);
    }

    public function filterPtAds(Request $request)
    {
        $start = $request->startDate;
        $end   = $request->endDate;  

        $data = DB::connection('apicentral')->table('amazon_sp_campaign_reports')
            ->selectRaw('
                report_date_range,
                SUM(clicks) as clicks,
                SUM(spend) as spend,
                SUM(purchases1d) as orders,
                SUM(sales1d) as sales
            ')
            ->whereBetween('report_date_range', [$start, $end])
            ->where(function($query) {
                $query->whereRaw("campaignName LIKE '%PT'")
                    ->whereRaw("campaignName LIKE '%PT.'");
            })
            ->groupBy('report_date_range')
            ->orderBy('report_date_range', 'asc')
            ->get();

        return response()->json([
            'dates'  => $data->pluck('report_date_range'),
            'clicks' => $data->pluck('clicks')->map(fn($v) => (int) $v),
            'spend'  => $data->pluck('spend')->map(fn($v) => (float) $v),
            'orders' => $data->pluck('orders')->map(fn($v) => (int) $v),
            'sales'  => $data->pluck('sales')->map(fn($v) => (float) $v),
            'totals' => [
                'clicks' => $data->sum('clicks'),
                'spend'  => $data->sum('spend'),
                'orders' => $data->sum('orders'),
                'sales'  => $data->sum('sales'),
            ]
        ]);
    }


    public function getAmazonKwAdsData(){

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

        $amazonSpCampaignReportsL60 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L60')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->where('campaignStatus', '!=', 'ARCHIVED')
            ->get();

        $amazonSpCampaignReportsL30 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->where('campaignStatus', '!=', 'ARCHIVED')
            ->get();

        $amazonSpCampaignReportsL15 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L15')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->where('campaignStatus', '!=', 'ARCHIVED')
            ->get();

        $amazonSpCampaignReportsL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->where('campaignStatus', '!=', 'ARCHIVED')
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $amazonSheet = $amazonDatasheetsBySku[$sku] ?? null;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignL60 = $amazonSpCampaignReportsL60->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $matchedCampaignL15 = $amazonSpCampaignReportsL15->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $matchedCampaignL7 = $amazonSpCampaignReportsL7->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $row = [
                'parent' => $parent,
                'sku' => $pm->sku,
                'INV' => $shopify->inv ?? 0,
                'L30' => $shopify->quantity ?? 0,
                'A_L30' => $amazonSheet->units_ordered_l30 ?? 0,
                'campaignName' => $matchedCampaignL7->campaignName ?? '',
                'campaignStatus' => $matchedCampaignL30->campaignStatus ?? '',
                'campaignBudgetAmount' => $matchedCampaignL30->campaignBudgetAmount ?? 0,
                // L60
                'impressions_l60' => $matchedCampaignL60->impressions ?? 0,
                'clicks_l60'      => $matchedCampaignL60->clicks ?? 0,
                'spend_l60'       => $matchedCampaignL60->spend ?? 0,
                'ad_sales_l60'    => $matchedCampaignL60->sales60d ?? 0,
                'ad_sold_l60'     => $matchedCampaignL60->unitsSoldSameSku60d ?? 0,
                'acos_l60'        => ($matchedCampaignL60 && $matchedCampaignL60->sales60d > 0) ? round(($matchedCampaignL60->spend / $matchedCampaignL60->sales60d) * 100, 2) : 0,
                'cpc_l60'         => $matchedCampaignL60->costPerClick ?? 0,

                // L30
                'impressions_l30' => $matchedCampaignL30->impressions ?? 0,
                'clicks_l30'      => $matchedCampaignL30->clicks ?? 0,
                'spend_l30'       => $matchedCampaignL30->spend ?? 0,
                'ad_sales_l30'    => $matchedCampaignL30->sales30d ?? 0,
                'ad_sold_l30'     => $matchedCampaignL30->unitsSoldSameSku30d ?? 0,
                'acos_l30'        => ($matchedCampaignL30 && $matchedCampaignL30->sales30d > 0) ? round(($matchedCampaignL30->spend / $matchedCampaignL30->sales30d) * 100, 2) : 0,
                'cpc_l30'         => $matchedCampaignL30->costPerClick ?? 0,

                // L15
                'impressions_l15' => $matchedCampaignL15->impressions ?? 0,
                'clicks_l15'      => $matchedCampaignL15->clicks ?? 0,
                'spend_l15'       => $matchedCampaignL15->spend ?? 0,
                'ad_sales_l15'    => $matchedCampaignL15->sales14d ?? 0,
                'ad_sold_l15'     => $matchedCampaignL15->unitsSoldSameSku14d ?? 0,
                'acos_l15'        => ($matchedCampaignL15 && $matchedCampaignL15->sales14d > 0) ? round(($matchedCampaignL15->spend / $matchedCampaignL15->sales14d) * 100, 2) : 0,
                'cpc_l15'         => $matchedCampaignL15->costPerClick ?? 0,

                // L7
                'impressions_l7'  => $matchedCampaignL7->impressions ?? 0,
                'clicks_l7'       => $matchedCampaignL7->clicks ?? 0,
                'spend_l7'        => $matchedCampaignL7->spend ?? 0,
                'ad_sales_l7'     => $matchedCampaignL7->sales7d ?? 0,
                'ad_sold_l7'      => $matchedCampaignL7->unitsSoldSameSku7d ?? 0,
                'acos_l7'         => ($matchedCampaignL7 && $matchedCampaignL7->sales7d > 0) ? round(($matchedCampaignL7->spend / $matchedCampaignL7->sales7d) * 100, 2) : 0,
                'cpc_l7'          => $matchedCampaignL7->costPerClick ?? 0,

                'NRL' => '',
                'NRA' => '',
                'FBA' => '',
            ];

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NRL']  = $raw['NRL'] ?? null;
                    $row['NRA'] = $raw['NRA'] ?? null;
                    $row['FBA'] = $raw['FBA'] ?? null;
                    $row['TPFT'] = $raw['TPFT'] ?? null;
                }
            }

            if($row['campaignName'] != ''){
                $result[] = (object) $row;
            }
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    public function amazonPtAdsView(){
        $thirtyDaysAgo = \Carbon\Carbon::now()->subDays(30)->format('Y-m-d');

        $data = DB::connection('apicentral')->table('amazon_sp_campaign_reports')
            ->selectRaw('
                report_date_range,
                SUM(clicks) as clicks, 
                SUM(spend) as spend, 
                SUM(purchases1d) as orders, 
                SUM(sales1d) as sales
            ')
            ->whereDate('report_date_range', '>=', $thirtyDaysAgo)
            ->where(function($query) {
                $query->whereRaw("campaignName LIKE '%PT'")
                    ->orWhereRaw("campaignName LIKE '%PT.'");
            })
            ->groupBy('report_date_range')
            ->orderBy('report_date_range', 'asc')
            ->get();

        $dates  = $data->pluck('report_date_range')->map(fn($d) => \Carbon\Carbon::parse($d)->format('Y-m-d'));
        $clicks = $data->pluck('clicks')->map(fn($v) => (int) $v);
        $spend  = $data->pluck('spend')->map(fn($v) => (float) $v);
        $orders = $data->pluck('orders')->map(fn($v) => (int) $v);
        $sales  = $data->pluck('sales')->map(fn($v) => (float) $v);
        
        return view('campaign.amazon-pt-ads', compact('dates', 'clicks', 'spend', 'orders', 'sales'));
    }

    public function getAmazonPtAdsData(){

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

        $amazonSpCampaignReportsL60 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L60')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();
        
        $amazonSpCampaignReportsL30 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $amazonSpCampaignReportsL15 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L15')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $amazonSpCampaignReportsL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $amazonSheet = $amazonDatasheetsBySku[$sku] ?? null;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignL60 = $amazonSpCampaignReportsL60->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));

                return (
                    (str_ends_with($cleanName, $sku . ' PT') || str_ends_with($cleanName, $sku . ' PT.'))
                    && strtoupper($item->campaignStatus) === 'ENABLED'
                );
            });

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));

                return (
                    (str_ends_with($cleanName, $sku . ' PT') || str_ends_with($cleanName, $sku . ' PT.'))
                );
            });

            $matchedCampaignL15 = $amazonSpCampaignReportsL15->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));

                return (
                    (str_ends_with($cleanName, $sku . ' PT') || str_ends_with($cleanName, $sku . ' PT.'))
                );
            });

            $matchedCampaignL7 = $amazonSpCampaignReportsL7->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));

                return (
                    (str_ends_with($cleanName, $sku . ' PT') || str_ends_with($cleanName, $sku . ' PT.'))
                );
            });

            $row = [
                'parent' => $parent,
                'sku' => $pm->sku,
                'INV' => $shopify->inv ?? 0,
                'L30' => $shopify->quantity ?? 0,
                'A_L30' => $amazonSheet->units_ordered_l30 ?? 0,
                'campaignName' => $matchedCampaignL30->campaignName ?? '',
                'campaignStatus' => $matchedCampaignL30->campaignStatus ?? '',
                'campaignBudgetAmount' => $matchedCampaignL30->campaignBudgetAmount ?? 0,
                
                // L60
                'impressions_l60' => $matchedCampaignL60->impressions ?? 0,
                'clicks_l60'      => $matchedCampaignL60->clicks ?? 0,
                'spend_l60'       => $matchedCampaignL60->spend ?? 0,
                'ad_sales_l60'    => $matchedCampaignL60->sales30d ?? 0,
                'ad_sold_l60'     => $matchedCampaignL60->unitsSoldSameSku30d ?? 0,
                'acos_l60'        => ($matchedCampaignL60 && ($matchedCampaignL60->sales30d ?? 0) > 0) ? round(($matchedCampaignL60->spend / $matchedCampaignL60->sales30d) * 100, 2) : 0,
                'cpc_l60'         => $matchedCampaignL60->costPerClick ?? 0,

                // L30
                'impressions_l30' => $matchedCampaignL30->impressions ?? 0,
                'clicks_l30'      => $matchedCampaignL30->clicks ?? 0,
                'spend_l30'       => $matchedCampaignL30->spend ?? 0,
                'ad_sales_l30'    => $matchedCampaignL30->sales30d ?? 0,
                'ad_sold_l30'     => $matchedCampaignL30->unitsSoldSameSku30d ?? 0,
                'acos_l30'        => ($matchedCampaignL30 && ($matchedCampaignL30->sales30d ?? 0) > 0) ? round(($matchedCampaignL30->spend / $matchedCampaignL30->sales30d) * 100, 2) : 0,
                'cpc_l30'         => $matchedCampaignL30->costPerClick ?? 0,

                // L15
                'impressions_l15' => $matchedCampaignL15->impressions ?? 0,
                'clicks_l15'      => $matchedCampaignL15->clicks ?? 0,
                'spend_l15'       => $matchedCampaignL15->spend ?? 0,
                'ad_sales_l15'    => ($matchedCampaignL15->sales1d ?? 0) + ($matchedCampaignL15->sales14d ?? 0),
                'ad_sold_l15'     => ($matchedCampaignL15->unitsSoldSameSku1d ?? 0) + ($matchedCampaignL15->unitsSoldSameSku14d ?? 0),
                'acos_l15'        => (($matchedCampaignL15->spend ?? 0) > 0 && (($matchedCampaignL15->sales1d ?? 0) + ($matchedCampaignL15->sales14d ?? 0)) > 0) 
                                    ? round(($matchedCampaignL15->spend / (($matchedCampaignL15->sales1d ?? 0) + ($matchedCampaignL15->sales14d ?? 0))) * 100, 2) 
                                    : 0,
                'cpc_l15'         => $matchedCampaignL15->costPerClick ?? 0,

                // L7
                'impressions_l7'  => $matchedCampaignL7->impressions ?? 0,
                'clicks_l7'       => $matchedCampaignL7->clicks ?? 0,
                'spend_l7'        => $matchedCampaignL7->spend ?? 0,
                'ad_sales_l7'     => $matchedCampaignL7->sales7d ?? 0,
                'ad_sold_l7'      => $matchedCampaignL7->unitsSoldSameSku7d ?? 0,
                'acos_l7'         => ($matchedCampaignL7 && ($matchedCampaignL7->sales7d ?? 0) > 0) ? round(($matchedCampaignL7->spend / $matchedCampaignL7->sales7d) * 100, 2) : 0,
                'cpc_l7'          => $matchedCampaignL7->costPerClick ?? 0,



                'NRL' => '',
                'NRA' => '',
                'FBA' => '',
            ];

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NRL']  = $raw['NRL'] ?? null;
                    $row['NRA'] = $raw['NRA'] ?? null;
                    $row['FBA'] = $raw['FBA'] ?? null;
                    $row['TPFT'] = $raw['TPFT'] ?? null;
                }
            }

            if($row['campaignName'] != ''){
                $result[] = (object) $row;
            }
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    public function amazonHlAdsView(){
        return view('campaign.amazon-hl-ads');
    }

    public function getAmazonHlAdsData(){

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

        $amazonSpCampaignReportsL60 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L60')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $amazonSpCampaignReportsL30 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $amazonSpCampaignReportsL15 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L15')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $amazonSpCampaignReportsL7 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $amazonSheet = $amazonDatasheetsBySku[$sku] ?? null;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignL60 = $amazonSpCampaignReportsL60->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                $expected1 = $sku;                
                $expected2 = $sku . ' HEAD';      

                return ($cleanName === $expected1 || $cleanName === $expected2)
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                $expected1 = $sku;                
                $expected2 = $sku . ' HEAD';      

                return ($cleanName === $expected1 || $cleanName === $expected2)
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            $matchedCampaignL15 = $amazonSpCampaignReportsL15->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                $expected1 = $sku;                
                $expected2 = $sku . ' HEAD';      

                return ($cleanName === $expected1 || $cleanName === $expected2)
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            $matchedCampaignL7 = $amazonSpCampaignReportsL7->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                $expected1 = $sku;                
                $expected2 = $sku . ' HEAD';      

                return ($cleanName === $expected1 || $cleanName === $expected2)
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            if(!$matchedCampaignL60 && !$matchedCampaignL30 && !$matchedCampaignL15 && !$matchedCampaignL7){
                continue;
            }

            // L7
            $costPerClick7 = ($matchedCampaignL7 && $matchedCampaignL7->clicks > 0)
                ? ($matchedCampaignL7->cost / $matchedCampaignL7->clicks)
                : 0;

            // L15
            $costPerClick15 = ($matchedCampaignL15 && $matchedCampaignL15->clicks > 0)
                ? ($matchedCampaignL15->cost / $matchedCampaignL15->clicks)
                : 0;

            // L30
            $costPerClick30 = ($matchedCampaignL30 && $matchedCampaignL30->clicks > 0)
                ? ($matchedCampaignL30->cost / $matchedCampaignL30->clicks)
                : 0;

            // L60
            $costPerClick60 = ($matchedCampaignL60 && $matchedCampaignL60->clicks > 0)
                ? ($matchedCampaignL60->cost / $matchedCampaignL60->clicks)
                : 0;

            $row = [
                'parent' => $parent,
                'sku' => $pm->sku,
                'INV' => $shopify->inv ?? 0,
                'L30' => $shopify->quantity ?? 0,
                'A_L30' => $amazonSheet->units_ordered_l30 ?? 0,
                'campaignName' => $matchedCampaignL30->campaignName ?? '',
                'campaignStatus' => $matchedCampaignL30->campaignStatus ?? '',
                'campaignBudgetAmount' => $matchedCampaignL30->campaignBudgetAmount ?? 0,

                'impressions_l7' => $matchedCampaignL7->impressions ?? 0,
                'clicks_l7'      => $matchedCampaignL7->clicks ?? 0,
                'spend_l7'       => $matchedCampaignL7->cost ?? 0,
                'ad_sales_l7'    => $matchedCampaignL7->sales ?? 0,
                'ad_sold_l7'     => $matchedCampaignL7->unitsSold ?? 0,
                'acos_l7'        => ($matchedCampaignL7 && $matchedCampaignL7->sales > 0)
                                        ? round(($matchedCampaignL7->cost / $matchedCampaignL7->sales) * 100, 2) : 0,
                'cpc_l7'         => $costPerClick7,

                'impressions_l15' => $matchedCampaignL15->impressions ?? 0,
                'clicks_l15'      => $matchedCampaignL15->clicks ?? 0,
                'spend_l15'       => $matchedCampaignL15->cost ?? 0,
                'ad_sales_l15'    => $matchedCampaignL15->sales ?? 0,
                'ad_sold_l15'     => $matchedCampaignL15->unitsSold ?? 0,
                'acos_l15'        => ($matchedCampaignL15 && $matchedCampaignL15->sales > 0)
                                        ? round(($matchedCampaignL15->cost / $matchedCampaignL15->sales) * 100, 2) : 0,
                'cpc_l15'         => $costPerClick15,

                'impressions_l30' => $matchedCampaignL30->impressions ?? 0,
                'clicks_l30'      => $matchedCampaignL30->clicks ?? 0,
                'spend_l30'       => $matchedCampaignL30->cost ?? 0,
                'ad_sales_l30'    => $matchedCampaignL30->sales ?? 0,
                'ad_sold_l30'     => $matchedCampaignL30->unitsSold ?? 0,
                'acos_l30'        => ($matchedCampaignL30 && $matchedCampaignL30->sales > 0)
                                        ? round(($matchedCampaignL30->cost / $matchedCampaignL30->sales) * 100, 2) : 0,
                'cpc_l30'         => $costPerClick30,

                'impressions_l60' => $matchedCampaignL60->impressions ?? 0,
                'clicks_l60'      => $matchedCampaignL60->clicks ?? 0,
                'spend_l60'       => $matchedCampaignL60->cost ?? 0,
                'ad_sales_l60'    => $matchedCampaignL60->sales ?? 0,
                'ad_sold_l60'     => $matchedCampaignL60->unitsSold ?? 0,
                'acos_l60'        => ($matchedCampaignL60 && $matchedCampaignL60->sales > 0)
                                        ? round(($matchedCampaignL60->cost / $matchedCampaignL60->sales) * 100, 2) : 0,
                'cpc_l60'         => $costPerClick60,


                'NRL' => '',
                'NRA' => '',
                'FBA' => '',
            ];

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NRL']  = $raw['NRL'] ?? null;
                    $row['NRA'] = $raw['NRA'] ?? null;
                    $row['FBA'] = $raw['FBA'] ?? null;
                    $row['TPFT'] = $raw['TPFT'] ?? null;
                }
            }

            $result[] = (object) $row;
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    public function getAmazonCampaignsData(){

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

        $amazonSpCampaignReportsL30 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
                }
            })
            ->get();

        $amazonSpCampaignReportsL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
                }
            })
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $amazonSheet = $amazonDatasheetsBySku[$sku] ?? null;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
                return strcasecmp(trim($item->campaignName), $sku) === 0;
            });

            $matchedCampaignL7 = $amazonSpCampaignReportsL7->first(function ($item) use ($sku) {
                return strcasecmp(trim($item->campaignName), $sku) === 0;
            });

            $row = [];
            $row['parent'] = $parent;
            $row['sku']    = $pm->sku;
            $row['INV']    = $shopify->inv ?? 0;
            $row['L30']    = $shopify->quantity ?? 0;
            $row['fba']    = $pm->fba ?? null;
            $row['A_L30']  = $amazonSheet->units_ordered_l30 ?? 0;
            $row['campaign_id'] = $matchedCampaignL30->campaign_id ??  '';
            $row['campaignName'] = $matchedCampaignL30->campaignName ?? '';
            $row['campaignStatus'] = $matchedCampaignL30->campaignStatus ?? '';
            $row['campaignBudgetAmount'] = $matchedCampaignL30->campaignBudgetAmount ?? '';
            $row['l7_cpc'] = $matchedCampaignL7->costPerClick ?? 0;
            
            $row['acos_L30'] = ($matchedCampaignL30 && ($matchedCampaignL30->sales30d ?? 0) > 0)
                ? round(($matchedCampaignL30->spend / $matchedCampaignL30->sales30d) * 100, 2)
                : null;

            $row['clicks_L30'] = $matchedCampaignL30->clicks ?? 0;
            $row['spend_L30'] = $matchedCampaignL30->spend ?? 0;
            $row['sales_L30'] = $matchedCampaignL30->sales30d ?? 0;
            $row['sold_L30'] = $matchedCampaignL30->unitsSoldSameSku30d ?? 0;

            $row['NRL']  = '';
            $row['NRA'] = '';
            $row['FBA'] = '';
            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NRL']  = $raw['NRL'] ?? null;
                    $row['NRA'] = $raw['NRA'] ?? null;
                    $row['FBA'] = $raw['FBA'] ?? null;
                }
            }

            $result[] = (object) $row;
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }
}
