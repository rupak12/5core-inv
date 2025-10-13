<?php

namespace App\Console\Commands;

use App\Http\Controllers\Campaigns\EbayOverUtilizedBgtController;
use App\Models\EbayDataView;
use App\Models\EbayMetric;
use App\Models\EbayPriorityReport;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Console\Command;
use Illuminate\Support\Facades\Log;

class EbayPriceLessBidsAutoUpdate extends Command
{
    protected $signature = 'ebay:auto-update-price-less-bids';
    protected $description = 'Automatically update Ebay campaign keyword bids for price less than 20';

    protected $profileId;

    public function __construct()
    {
        parent::__construct();
    }

    public function handle()
    {
        $this->info("Starting Ebay bids auto-update...");

        $updateOverUtilizedBids = new EbayOverUtilizedBgtController;

        $campaigns = $this->getEbayPriceLessBidsCampaign();

        if (empty($campaigns)) {
            $this->warn("No campaigns matched filter conditions.");
            return 0;
        }

        $campaignIds = collect($campaigns)->pluck('campaign_id')->toArray();
        $newBids = collect($campaigns)->pluck('sbid')->toArray();

        $result = $updateOverUtilizedBids->updateAutoKeywordsBidDynamic($campaignIds, $newBids);
        $this->info("Update Result: " . json_encode($result));

    }

    public function getEbayPriceLessBidsCampaign(){

        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        $ebayMetricData = EbayMetric::whereIn('sku', $skus)->get()->keyBy('sku');

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

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);

            $shopify = $shopifyData[$pm->sku] ?? null;

            $ebay = $ebayMetricData[$pm->sku] ?? null;

            $matchedCampaignL7 = $ebayCampaignReportsL7->first(function ($item) use ($sku) {
                return stripos($item->campaign_name, $sku) !== false;
            });

            $matchedCampaignL1 = $ebayCampaignReportsL1->first(function ($item) use ($sku) {
                return stripos($item->campaign_name, $sku) !== false;
            });

            if (!$matchedCampaignL7 && !$matchedCampaignL1) {
                continue;
            }

            $row = [];
            $row['INV']    = $shopify->inv ?? 0;
            $row['L30']    = $shopify->quantity ?? 0;
            $row['price']  = $ebay->ebay_price ?? 0;
            $row['campaign_id'] = $matchedCampaignL7->campaign_id ?? ($matchedCampaignL1->campaign_id ?? '');
            $row['sbid'] = 0.10;

            if($row['price'] < 30){
                if($row['price'] < 10){
                    $row['sbid'] = 0.10;
                }elseif($row['price'] > 10 && $row['price'] <= 20){
                    $row['sbid'] = 0.20;
                }elseif($row['price'] > 20 && $row['price'] <= 30){
                    $row['sbid'] = 0.30;
                }
                $result[] = (object) $row;
            }

        }

        return $result;
    }


}