<?php

namespace App\Http\Controllers\Campaigns;

use App\Http\Controllers\Controller;
use App\Models\Ebay3GeneralReport;
use App\Models\Ebay3Metric;
use App\Models\Ebay3PriorityReport;
use App\Models\EbayGeneralReport;
use App\Models\EbayThreeDataView;
use App\Models\MarketplacePercentage;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;

class Ebay3PmtAdsController extends Controller
{
    public function index()
    {
        $marketplaceData = MarketplacePercentage::where("marketplace", "Ebay3" )->first();
        $ebayPercentage = $marketplaceData ? $marketplaceData->percentage : 100;
        $ebayAdPercentage = $marketplaceData ? $marketplaceData->ad_updates : 100;

        return view('campaign.ebay-three.pmt-ads', compact('ebayPercentage','ebayAdPercentage'));
    }

    public function getEbay3PmtAdsData()
    {
        $productMasters = ProductMaster::orderBy("parent", "asc")
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy("sku", "asc")
            ->get();

        $skus = $productMasters->pluck("sku")->filter()->unique()->values()->all();

        $shopifyData = ShopifySku::whereIn("sku", $skus)->get()->keyBy("sku");
        $ebayMetrics = Ebay3Metric::whereIn("sku", $skus)->get()->keyBy("sku");
        $nrValues = EbayThreeDataView::whereIn("sku", $skus)->pluck("value", "sku");

        $itemIdToSku = $ebayMetrics->pluck('sku', 'item_id')->toArray();
        $campaignIdToSku = $ebayMetrics->pluck('sku', 'campaign_id')->toArray();

        $extraClicksData = Ebay3GeneralReport::whereIn('listing_id', array_keys($itemIdToSku))
            ->where('report_range', 'L30')
            ->pluck('clicks', 'listing_id')
            ->toArray();

        $generalReports = Ebay3GeneralReport::whereIn('listing_id', array_keys($itemIdToSku))
            ->whereIn('report_range', ['L60', 'L30', 'L7'])
            ->get();

        $priorityReports = Ebay3PriorityReport::whereIn('campaign_id', array_keys($campaignIdToSku))
            ->whereIn('report_range', ['L60', 'L30', 'L7'])
            ->get();

        $campaignListings = DB::connection('apicentral')->table('ebay3_campaign_ads_listings')->select('listing_id', 'bid_percentage', 'suggested_bid')->get()->keyBy('listing_id')->toArray();

        $adMetricsBySku = [];

        foreach ($generalReports as $report) {
            $sku = $itemIdToSku[$report->listing_id] ?? null;
            if (!$sku) continue;

            $range = strtoupper($report->report_range);

            $adMetricsBySku[$sku][$range]['GENERAL_SPENT'] =
                ($adMetricsBySku[$sku][$range]['GENERAL_SPENT'] ?? 0) + $this->extractNumber($report->ad_fees);

            $adMetricsBySku[$sku][$range]['Imp'] =
                ($adMetricsBySku[$sku][$range]['Imp'] ?? 0) + (int) $report->impressions;

            $adMetricsBySku[$sku][$range]['Clk'] =
                ($adMetricsBySku[$sku][$range]['Clk'] ?? 0) + (int) $report->clicks;

            $adMetricsBySku[$sku][$range]['Ctr'] =
                ($adMetricsBySku[$sku][$range]['Ctr'] ?? 0) + (float) $report->ctr;

            $adMetricsBySku[$sku][$range]['Sls'] =
                ($adMetricsBySku[$sku][$range]['Sls'] ?? 0) + (int) $report->sales;
        }

        foreach ($priorityReports as $report) {
            $sku = $campaignIdToSku[$report->campaign_id] ?? null;
            if (!$sku) continue;

            $range = strtoupper($report->report_range);

            $adMetricsBySku[$sku][$range]['PRIORITY_SPENT'] =
                ($adMetricsBySku[$sku][$range]['PRIORITY_SPENT'] ?? 0) + $this->extractNumber($report->cpc_ad_fees_payout_currency);
        }

        $marketplaceData = MarketplacePercentage::where("marketplace", "Ebay" )->first();
        $percentage = $marketplaceData ? ($marketplaceData->percentage / 100) : 1;
        $adPercentage = $marketplaceData ? ($marketplaceData->ad_updates / 100) : 1;

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);
            $parent = $pm->parent;

            $shopify = $shopifyData[$pm->sku] ?? null;
            $ebayMetric = $ebayMetrics[$pm->sku] ?? null;

            $row = [];
            $row["Parent"] = $parent;
            $row["(Child) sku"] = $pm->sku;
            $row['fba'] = $pm->fba;

            $row["INV"] = $shopify->inv ?? 0;
            $row["L30"] = $shopify->quantity ?? 0;

            $row["eBay L30"] = $ebayMetric->ebay_l30 ?? 0;
            $row["eBay L60"] = $ebayMetric->ebay_l60 ?? 0;
            $row["eBay Price"] = $ebayMetric->ebay_price ?? 0;
            $row['price_lmpa'] = $ebayMetric->price_lmpa ?? null;
            $row['eBay_item_id'] = $ebayMetric->item_id ?? null;
            $row['ebay_views'] = $ebayMetric->views ?? 0;

            if ($ebayMetric && isset($campaignListings[$ebayMetric->item_id])) {
                $row['bid_percentage'] = $campaignListings[$ebayMetric->item_id]->bid_percentage ?? null;
                $row['suggested_bid']  = $campaignListings[$ebayMetric->item_id]->suggested_bid ?? null;
            } else {
                $row['bid_percentage'] = null;
                $row['suggested_bid']  = null;
            }


            $row["E Dil%"] = ($row["eBay L30"] && $row["INV"] > 0)
                ? round(($row["eBay L30"] / $row["INV"]), 2)
                : 0;

            $pmtData = $adMetricsBySku[$sku] ?? [];
            foreach (['L60', 'L30', 'L7'] as $range) {
                $metrics = $pmtData[$range] ?? [];
                foreach (['Imp', 'Clk', 'Ctr', 'Sls', 'GENERAL_SPENT', 'PRIORITY_SPENT'] as $suffix) {
                    $key = "Pmt{$suffix}{$range}";
                    $row[$key] = $metrics[$suffix] ?? 0;
                }
            }

            $row["PmtClkL30"] = $adMetricsBySku[$sku]['L30']['Clk'] ?? 0;
            if ($ebayMetric && isset($extraClicksData[$ebayMetric->item_id])) {
                $row["PmtClkL30"] += (int) $extraClicksData[$ebayMetric->item_id];
            }

            $values = is_array($pm->Values) ? $pm->Values : (is_string($pm->Values) ? json_decode($pm->Values, true) : []);
            $lp = 0;
            foreach ($values as $k => $v) {
                if (strtolower($k) === "lp") {
                    $lp = floatval($v);
                    break;
                }
            }
            if ($lp === 0 && isset($pm->lp)) {
                $lp = floatval($pm->lp);
            }

            $ship = isset($values["ship"]) ? floatval($values["ship"]) : (isset($pm->ship) ? floatval($pm->ship) : 0);

            $price = floatval($row["eBay Price"] ?? 0);
            $units_ordered_l30 = floatval($row["eBay L30"] ?? 0);

            $generalSpent = $adMetricsBySku[$sku]['L30']['GENERAL_SPENT'] ?? 0;
            $prioritySpent = $adMetricsBySku[$sku]['L30']['PRIORITY_SPENT'] ?? 0;
            $denominator = ($price * $units_ordered_l30);
            $row["TacosL30"] = $denominator > 0 ? round((($generalSpent + $prioritySpent) / $denominator), 4) : 0;

            $row["Total_pft"] = round(($price * $percentage - $lp - $ship) * $units_ordered_l30, 2);
            $row["T_Sale_l30"] = round($price * $units_ordered_l30, 2);
            $row["PFT %"] = round(
                $price > 0 ? (($price * $percentage - $lp - $ship) / $price) : 0,
                2
            );
            $row["ROI%"] = round(
                $lp > 0 ? (($price * $percentage - $lp - $ship) / $lp) : 0,
                2
            );
            $row["TPFT"] = $row["PFT %"] + $adPercentage - $row["bid_percentage"];
            $row["percentage"] = $percentage;
            $row["LP_productmaster"] = $lp;
            $row["Ship_productmaster"] = $ship;

            $row['NRL'] = "";
            $row['SPRICE'] = null;
            $row['SPFT'] = null;
            $row['SROI'] = null;
            $row['Listed'] = null;
            $row['Live'] = null;
            $row['APlus'] = null;

            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) {
                    $raw = json_decode($raw, true);
                }
                if (is_array($raw)) {
                    $row['NRL'] = $raw['NRL'] ?? null;
                    $row['SPRICE'] = $raw['SPRICE'] ?? null;
                    $row['SPFT'] = $raw['SPFT'] ?? null;
                    $row['SROI'] = $raw['SROI'] ?? null;
                    $row['Listed'] = isset($raw['Listed']) ? filter_var($raw['Listed'], FILTER_VALIDATE_BOOLEAN) : null;
                    $row['Live'] = isset($raw['Live']) ? filter_var($raw['Live'], FILTER_VALIDATE_BOOLEAN) : null;
                    $row['APlus'] = isset($raw['APlus']) ? filter_var($raw['APlus'], FILTER_VALIDATE_BOOLEAN) : null;
                }
            }

            $row["image_path"] = $shopify->image_src ?? ($values["image_path"] ?? ($pm->image_path ?? null));

            if($row['NRL'] !== 'NRL'){
                $result[] = (object) $row;
            }
        }

        return response()->json([
            "message" => "eBay Data Fetched Successfully",
            "data" => $result,
            "status" => 200,
        ]);
    }

    private function extractNumber($value)
    {
        if (empty($value)) return 0;
        return (float) preg_replace('/[^\d.]/', '', $value);
    }

    public function saveEbay3PMTSpriceToDatabase(Request $request)
    {
        $sku = $request->input('sku');
        $spriceData = $request->only(['sprice', 'spft_percent', 'sroi_percent']);

        if (!$sku || !$spriceData['sprice']) {
            return response()->json(['error' => 'SKU and sprice are required.'], 400);
        }


        $ebayDataView = EbayThreeDataView::firstOrNew(['sku' => $sku]);

        $existing = is_array($ebayDataView->value)
            ? $ebayDataView->value
            : (json_decode($ebayDataView->value, true) ?: []);

        $merged = array_merge($existing, [
            'SPRICE' => $spriceData['sprice'],
            'SPFT' => $spriceData['spft_percent'],
            'SROI' => $spriceData['sroi_percent'],
        ]);

        $ebayDataView->value = $merged;
        $ebayDataView->save();

        return response()->json(['message' => 'Data saved successfully.']);
    }

    public function updateEbay3Percentage(Request $request)
    {
        try {
            $type = $request->input('type');
            $value = $request->input('value');

            $marketplace = MarketplacePercentage::where('marketplace', 'Ebay3')->first();

            $percent = $marketplace->percentage ?? 0;
            $adUpdates = $marketplace->ad_updates ?? 0;

            if ($type === 'percentage') {
                if (!is_numeric($value) || $value < 0 || $value > 100) {
                    return response()->json(['status' => 400, 'message' => 'Invalid percentage value'], 400);
                }
                $percent = $value;
            }

            if ($type === 'ad_updates') {
                if (!is_numeric($value) || $value < 0) {
                    return response()->json(['status' => 400, 'message' => 'Invalid ad_updates value'], 400);
                }
                $adUpdates = $value;
            }

            $marketplace = MarketplacePercentage::updateOrCreate(
                ['marketplace' => 'Ebay3'],
                [
                    'percentage' => $percent,
                    'ad_updates' => $adUpdates,
                ]
            );

            return response()->json([
                'status' => 200,
                'message' => ucfirst($type) . ' updated successfully!',
                'data' => [
                    'percentage' => $marketplace->percentage,
                    'ad_updates' => $marketplace->ad_updates
                ]
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'status' => 500,
                'message' => 'Error updating Ebay3 marketplace values',
                'error' => $e->getMessage()
            ], 500);
        }
    }
}
