<?php

namespace App\Http\Controllers\MarketPlace\ACOSControl;

use App\Http\Controllers\Controller;
use App\Models\AmazonDatasheet;
use App\Models\AmazonDataView;
use App\Models\AmazonSbCampaignReport;
use App\Models\AmazonSpCampaignReport;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use AWS\CRT\Log;
use GuzzleHttp\Client;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log as FacadesLog;

class AmazonACOSController extends Controller
{
    protected $profileId;

    public function __construct()
    {
        parent::__construct();
        $this->profileId = env('AMAZON_ADS_PROFILE_IDS');
    }

    public function getAccessToken()
    {
        return cache()->remember('amazon_ads_access_token', 55 * 60, function () {
            $client = new Client();

            $response = $client->post('https://api.amazon.com/auth/o2/token', [
                'form_params' => [
                    'grant_type' => 'refresh_token',
                    'refresh_token' => env('AMAZON_ADS_REFRESH_TOKEN'),
                    'client_id' => env('AMAZON_ADS_CLIENT_ID'),
                    'client_secret' => env('AMAZON_ADS_CLIENT_SECRET'),
                ]
            ]);

            $data = json_decode($response->getBody(), true);
            return $data['access_token'];
        });
    }

    public function updateAutoAmazonCampaignBgt(array $campaignIds, array $newBgts)
    {
        ini_set('max_execution_time', 300);
        ini_set('memory_limit', '512M');

        if (empty($campaignIds) || empty($newBgts)) {
            return response()->json([
                'message' => 'Campaign IDs and new budgets are required',
                'status' => 400
            ]);
        }

        $allCampaigns = [];

        foreach ($campaignIds as $index => $campaignId) {
            $newBgt = floatval($newBgts[$index] ?? 0);

            $allCampaigns[] = [
                'campaignId' => $campaignId,
                'budget' => [
                    'budget' => $newBgt,
                    'budgetType' => 'DAILY'
                ]
            ];
        }

        if (empty($allCampaigns)) {
            return response()->json([
                'message' => 'No campaigns found to update',
                'status' => 404,
            ]);
        }

        $accessToken = $this->getAccessToken();
        $client = new Client();
        $url = 'https://advertising-api.amazon.com/sp/campaigns';
        $results = [];

        try {
            $chunks = array_chunk($allCampaigns, 100);
            foreach ($chunks as $chunk) {
                $response = $client->put($url, [
                    'headers' => [
                        'Amazon-Advertising-API-ClientId' => env('AMAZON_ADS_CLIENT_ID'),
                        'Authorization' => 'Bearer ' . $accessToken,
                        'Amazon-Advertising-API-Scope' => $this->profileId,
                        'Content-Type' => 'application/vnd.spCampaign.v3+json',
                        'Accept' => 'application/vnd.spCampaign.v3+json',
                    ],
                    'json' => [
                        'campaigns' => $chunk
                    ],
                    'timeout' => 60,
                    'connect_timeout' => 30,
                ]);

                $results[] = json_decode($response->getBody(), true);
            }
            return [
                'message' => 'BGT updated successfully',
                'data' => $results,
                'status' => 200,
            ];

        } catch (\Exception $e) {
            return [
                'message' => 'Error updating BGT',
                'error' => $e->getMessage(),
                'status' => 500,
            ];
        }
    }

    public function updateAmazonCampaignBgt(Request $request)
    {
        ini_set('max_execution_time', 300);
        ini_set('memory_limit', '512M');

        $campaignIds = $request->input('campaign_ids', []);
        $newBgts = $request->input('bgts', []);

        if (empty($campaignIds) || empty($newBgts)) {
            return response()->json([
                'message' => 'Campaign IDs and new budgets are required',
                'status' => 400
            ]);
        }

        $allCampaigns = [];

        foreach ($campaignIds as $index => $campaignId) {
            $newBgt = floatval($newBgts[$index] ?? 0);

            $allCampaigns[] = [
                'campaignId' => $campaignId,
                'budget' => [
                    'budget' => $newBgt,
                    'budgetType' => 'DAILY'
                ]
            ];
        }

        if (empty($allCampaigns)) {
            return response()->json([
                'message' => 'No campaigns found to update',
                'status' => 404,
            ]);
        }

        $accessToken = $this->getAccessToken();
        $client = new Client();
        $url = 'https://advertising-api.amazon.com/sp/campaigns';
        $results = [];

        try {
            $chunks = array_chunk($allCampaigns, 100);
            foreach ($chunks as $chunk) {
                $response = $client->put($url, [
                    'headers' => [
                        'Amazon-Advertising-API-ClientId' => env('AMAZON_ADS_CLIENT_ID'),
                        'Authorization' => 'Bearer ' . $accessToken,
                        'Amazon-Advertising-API-Scope' => $this->profileId,
                        'Content-Type' => 'application/vnd.spCampaign.v3+json',
                        'Accept' => 'application/vnd.spCampaign.v3+json',
                    ],
                    'json' => [
                        'campaigns' => $chunk
                    ],
                    'timeout' => 60,
                    'connect_timeout' => 30,
                ]);

                $results[] = json_decode($response->getBody(), true);
            }
            return response()->json([
                'message' => 'Campaign budget updated successfully',
                'data' => $results,
                'status' => 200,
            ]);

        } catch (\Exception $e) {
            return response()->json([
                'message' => 'Error updating campaign budgets',
                'error' => $e->getMessage(),
                'status' => 500,
            ]);
        }
    }

    public function updateAutoAmazonSbCampaignBgt(array $campaignIds, array $newBgts)
    {
        ini_set('max_execution_time', 300);
        ini_set('memory_limit', '512M');

        if (empty($campaignIds) || empty($newBgts)) {
            return response()->json([
                'message' => 'Campaign IDs and new budgets are required',
                'status' => 400
            ]);
        }

        $allCampaigns = [];

        foreach ($campaignIds as $index => $campaignId) {
            $newBgt = floatval($newBgts[$index] ?? 0);

            $allCampaigns[] = [
                'campaignId' => $campaignId,
                'budget' => $newBgt,
            ];
        }

        if (empty($allCampaigns)) {
            return response()->json([
                'message' => 'No campaigns found to update',
                'status' => 404,
            ]);
        }

        $accessToken = $this->getAccessToken();
        $client = new Client();
        $url = 'https://advertising-api.amazon.com/sb/v4/campaigns';
        $results = [];

        try {
            $chunks = array_chunk($allCampaigns, 10);
            foreach ($chunks as $chunk) {
                $response = $client->put($url, [
                    'headers' => [
                        'Amazon-Advertising-API-ClientId' => env('AMAZON_ADS_CLIENT_ID'),
                        'Authorization' => 'Bearer ' . $accessToken,
                        'Amazon-Advertising-API-Scope' => $this->profileId,
                        'Content-Type' => 'application/vnd.sbcampaignresource.v4+json',
                        'Accept' => 'application/vnd.sbcampaignresource.v4+json',
                    ],
                    'json' => [
                        'campaigns' => $chunk
                    ],
                    'timeout' => 60,
                    'connect_timeout' => 30,
                ]);

                $results[] = json_decode($response->getBody(), true);
            }
            return [
                'message' => 'Campaign bgt updated successfully',
                'data' => $results,
                'status' => 200,
            ];

        } catch (\Exception $e) {
            return [
                'message' => 'Error updating campaign bgt',
                'error' => $e->getMessage(),
                'status' => 500,
            ];
        }
    }

    public function updateAmazonSbCampaignBgt(Request $request)
    {
        ini_set('max_execution_time', 300);
        ini_set('memory_limit', '512M');
        
        $campaignIds = $request->input('campaign_ids', []);
        $newBgts = $request->input('bgts', []);

        if (empty($campaignIds) || empty($newBgts)) {
            return response()->json([
                'message' => 'Campaign IDs and new budgets are required',
                'status' => 400
            ]);
        }

        $allCampaigns = [];

        foreach ($campaignIds as $index => $campaignId) {
            $newBgt = floatval($newBgts[$index] ?? 0);

            $allCampaigns[] = [
                'campaignId' => $campaignId,
                'budget' => $newBgt,
            ];
        }

        if (empty($allCampaigns)) {
            return response()->json([
                'message' => 'No campaigns found to update',
                'status' => 404,
            ]);
        }

        $accessToken = $this->getAccessToken();
        $client = new Client();
        $url = 'https://advertising-api.amazon.com/sb/v4/campaigns';
        $results = [];

        try {
            $chunks = array_chunk($allCampaigns, 10);
            foreach ($chunks as $chunk) {
                $response = $client->put($url, [
                    'headers' => [
                        'Amazon-Advertising-API-ClientId' => env('AMAZON_ADS_CLIENT_ID'),
                        'Authorization' => 'Bearer ' . $accessToken,
                        'Amazon-Advertising-API-Scope' => $this->profileId,
                        'Content-Type' => 'application/vnd.sbcampaignresource.v4+json',
                        'Accept' => 'application/vnd.sbcampaignresource.v4+json',
                    ],
                    'json' => [
                        'campaigns' => $chunk
                    ],
                    'timeout' => 60,
                    'connect_timeout' => 30,
                ]);

                $results[] = json_decode($response->getBody(), true);
            }
            return response()->json([
                'message' => 'Campaign budget updated successfully',
                'data' => $results,
                'status' => 200,
            ]);

        } catch (\Exception $e) {
            return response()->json([
                'message' => 'Error updating campaign budgets',
                'error' => $e->getMessage(),
                'status' => 500,
            ]);
        }
    }

    public function index(){
        return view('market-places.acos-control.amazon-acos-control');
    }

    public function getAmzonAcOSData()
    {
        $data = DB::table(function($query) {
            $query->select(
                'id', 'campaign_id', 'note', 'sbid', 'yes_sbid', 'profile_id', 'ad_type',
                'report_date_range', 'campaignName', 'clicks', 'cost', 'impressions',
                'startDate', 'endDate', 'sales',
                'campaignBudgetAmount', 'campaignBudgetCurrencyCode', DB::raw("'SB' as source"), 'campaignStatus'
            )
            ->from('amazon_sb_campaign_reports')
            ->unionAll(
                DB::table('amazon_sd_campaign_reports')->select(
                    'id', 'campaign_id', 'note', 'sbid', 'yes_sbid', 'profile_id', 'ad_type',
                    'report_date_range', 'campaignName', 'clicks', 'cost', 'impressions',
                    'startDate', 'endDate', 'sales',
                    DB::raw('NULL as campaignBudgetAmount'), 'campaignBudgetCurrencyCode', DB::raw("'SD' as source"), 'campaignStatus'
                )
            )
            ->unionAll(
                DB::table('amazon_sp_campaign_reports')->select(
                    'id', 'campaign_id', 'note', 'sbid', 'yes_sbid', 'profile_id', 'ad_type',
                    'report_date_range', 'campaignName', 'clicks', 'cost', 'impressions',
                    'startDate', 'endDate', 'sales30d as sales',
                    'campaignBudgetAmount', 'campaignBudgetCurrencyCode', DB::raw("'SP' as source"), 'campaignStatus'
                )
            );
        }, 'base')
        ->leftJoin('campaign_entries as ce', DB::raw('TRIM(base.campaignName)'), '=', DB::raw('TRIM(ce.campaign_name)'))
        ->leftJoin(DB::raw('(SELECT campaign_id, profile_id, source,
                SUM(CASE WHEN report_date_range="L7" THEN clicks ELSE 0 END) as l7_clicks,
                SUM(CASE WHEN report_date_range="L15" THEN clicks ELSE 0 END) as l15_clicks,
                SUM(CASE WHEN report_date_range="L30" THEN clicks ELSE 0 END) as l30_clicks,
                SUM(CASE WHEN report_date_range="L60" THEN clicks ELSE 0 END) as l60_clicks
            FROM (
                SELECT campaign_id, profile_id, report_date_range, clicks, "SB" as source FROM amazon_sb_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, clicks, "SD" as source FROM amazon_sd_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, clicks, "SP" as source FROM amazon_sp_campaign_reports
            ) AS reports
            GROUP BY campaign_id, profile_id, source
        ) as clicks_by_range'), function($join){
            $join->on('base.campaign_id', '=', 'clicks_by_range.campaign_id')
                 ->on('base.profile_id', '=', 'clicks_by_range.profile_id')
                 ->on('base.source', '=', 'clicks_by_range.source');
        })
        ->leftJoin(DB::raw('(SELECT campaign_id, profile_id, source,
                SUM(CASE WHEN report_date_range="L1" THEN cost ELSE 0 END) as l1_spend,
                SUM(CASE WHEN report_date_range="L7" THEN cost ELSE 0 END) as l7_spend,
                SUM(CASE WHEN report_date_range="L15" THEN cost ELSE 0 END) as l15_spend,
                SUM(CASE WHEN report_date_range="L30" THEN cost ELSE 0 END) as l30_spend,
                SUM(CASE WHEN report_date_range="L60" THEN cost ELSE 0 END) as l60_spend
            FROM (
                SELECT campaign_id, profile_id, report_date_range, cost, "SB" as source FROM amazon_sb_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, cost, "SD" as source FROM amazon_sd_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, cost, "SP" as source FROM amazon_sp_campaign_reports
            ) as spend_reports
            GROUP BY campaign_id, profile_id, source
        ) as spend_by_range'), function($join){
            $join->on('base.campaign_id', '=', 'spend_by_range.campaign_id')
                 ->on('base.profile_id', '=', 'spend_by_range.profile_id')
                 ->on('base.source', '=', 'spend_by_range.source');
        })
        ->leftJoin(DB::raw('(SELECT campaign_id, profile_id, source,
                SUM(CASE WHEN report_date_range="L7" THEN sales ELSE 0 END) as l7_sales,
                SUM(CASE WHEN report_date_range="L15" THEN sales ELSE 0 END) as l15_sales,
                SUM(CASE WHEN report_date_range="L30" THEN sales ELSE 0 END) as l30_sales,
                SUM(CASE WHEN report_date_range="L60" THEN sales ELSE 0 END) as l60_sales
            FROM (
                SELECT campaign_id, profile_id, report_date_range, sales, "SB" as source FROM amazon_sb_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, sales, "SD" as source FROM amazon_sd_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, sales30d as sales, "SP" as source FROM amazon_sp_campaign_reports
            ) as sales_reports
            GROUP BY campaign_id, profile_id, source
        ) as sales_by_range'), function($join){
            $join->on('base.campaign_id', '=', 'sales_by_range.campaign_id')
                 ->on('base.profile_id', '=', 'sales_by_range.profile_id')
                 ->on('base.source', '=', 'sales_by_range.source');
        })
        ->leftJoin(DB::raw('(SELECT campaign_id, profile_id, source,
                ROUND(SUM(CASE WHEN report_date_range="L7" THEN cost ELSE 0 END)/NULLIF(SUM(CASE WHEN report_date_range="L7" THEN clicks ELSE 0 END),0),2) as l7_cpc,
                ROUND(SUM(CASE WHEN report_date_range="L15" THEN cost ELSE 0 END)/NULLIF(SUM(CASE WHEN report_date_range="L15" THEN clicks ELSE 0 END),0),2) as l15_cpc,
                ROUND(SUM(CASE WHEN report_date_range="L30" THEN cost ELSE 0 END)/NULLIF(SUM(CASE WHEN report_date_range="L30" THEN clicks ELSE 0 END),0),2) as l30_cpc,
                ROUND(SUM(CASE WHEN report_date_range="L60" THEN cost ELSE 0 END)/NULLIF(SUM(CASE WHEN report_date_range="L60" THEN clicks ELSE 0 END),0),2) as l60_cpc
            FROM (
                SELECT campaign_id, profile_id, report_date_range, clicks, cost, "SB" as source FROM amazon_sb_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, clicks, cost, "SD" as source FROM amazon_sd_campaign_reports
                UNION ALL
                SELECT campaign_id, profile_id, report_date_range, clicks, cost, "SP" as source FROM amazon_sp_campaign_reports
            ) as cpc_reports
            GROUP BY campaign_id, profile_id, source
        ) as cpc_by_range'), function($join){
            $join->on('base.campaign_id', '=', 'cpc_by_range.campaign_id')
                 ->on('base.profile_id', '=', 'cpc_by_range.profile_id')
                 ->on('base.source', '=', 'cpc_by_range.source');
        })
        ->select('base.*', 'ce.parent',
            'clicks_by_range.l7_clicks','clicks_by_range.l15_clicks','clicks_by_range.l30_clicks','clicks_by_range.l60_clicks',
            'spend_by_range.l1_spend','spend_by_range.l7_spend','spend_by_range.l15_spend','spend_by_range.l30_spend','spend_by_range.l60_spend',
            'sales_by_range.l7_sales','sales_by_range.l15_sales','sales_by_range.l30_sales','sales_by_range.l60_sales',
            'cpc_by_range.l7_cpc','cpc_by_range.l15_cpc','cpc_by_range.l30_cpc','cpc_by_range.l60_cpc',
            DB::raw('ROUND(spend_by_range.l7_spend/NULLIF(sales_by_range.l7_sales,0)*100,2) as l7_acos'),
            DB::raw('ROUND(spend_by_range.l15_spend/NULLIF(sales_by_range.l15_sales,0)*100,2) as l15_acos'),
            DB::raw('ROUND(spend_by_range.l30_spend/NULLIF(sales_by_range.l30_sales,0)*100,2) as l30_acos'),
            DB::raw('ROUND(spend_by_range.l60_spend/NULLIF(sales_by_range.l60_sales,0)*100,2) as l60_acos')
        )
        ->get();

        return response()->json($data);
    }

    public function amazonAcosKwControl(){
        return view('market-places.acos-control.amazon-acos-kw-control');
    }

    public function amazonAcosKwControlData(){

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
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->get();

        $amazonSpCampaignReportsL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
            })
            ->where('campaignName', 'NOT LIKE', '%PT')
            ->where('campaignName', 'NOT LIKE', '%PT.')
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $amazonSheet = $amazonDatasheetsBySku[$sku] ?? null;
            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
            });

            $matchedCampaignL7 = $amazonSpCampaignReportsL7->first(function ($item) use ($sku) {
                $campaignName = strtoupper(trim(rtrim($item->campaignName, '.')));
                $cleanSku = strtoupper(trim(rtrim($sku, '.')));
                return $campaignName === $cleanSku;
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
                    $row['TPFT'] = $raw['TPFT'] ?? null;
                }
            }

            if ($row['NRA'] !== 'NRA') {
                $result[] = (object) $row;
            }

        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    public function amazonAcosHlControl(){
        return view('market-places.acos-control.amazon-acos-hl-control');
    }

    public function amazonAcosHlControlData(){

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

        $amazonSpCampaignReportsL30 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L30')
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

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
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
            
            $row['acos_L30'] = ($matchedCampaignL30 && ($matchedCampaignL30->sales ?? 0) > 0)
                ? round(($matchedCampaignL30->cost / $matchedCampaignL30->sales) * 100, 2)
                : null;

            $row['clicks_L30'] = $matchedCampaignL30->clicks ?? 0;

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
                    $row['TPFT'] = $raw['TPFT'] ?? null;
                }
            }

            if ($row['NRA'] !== 'NRA') {
                $result[] = (object) $row;
            }
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    public function amazonAcosPtControl(){
        return view('market-places.acos-control.amazon-acos-pt-control');
    }

    public function amazonAcosPtControlData(){

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

            $matchedCampaignL30 = $this->matchCampaign($sku, $amazonSpCampaignReportsL30);
            $matchedCampaignL7  = $this->matchCampaign($sku, $amazonSpCampaignReportsL7);

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
                    $row['TPFT'] = $raw['TPFT'] ?? null;
                }
            }

            if ($row['NRA'] !== 'NRA') {
                $result[] = (object) $row;
            }
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data'    => $result,
            'status'  => 200,
        ]);
    }

    function matchCampaign($sku, $campaignReports) {
        $skuClean = preg_replace('/\s+/', ' ', strtoupper(trim($sku)));

        $expected1 = $skuClean . ' PT';
        $expected2 = $skuClean . ' PT.';

        return $campaignReports->first(function ($item) use ($expected1, $expected2) {
            $campaignName = preg_replace('/\s+/', ' ', strtoupper(trim($item->campaignName)));

            return in_array($campaignName, [$expected1, $expected2], true)
                && strtoupper($item->campaignStatus) === 'ENABLED';
        });
    }
}
