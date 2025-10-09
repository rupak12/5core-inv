<?php

namespace App\Console\Commands;

use App\Http\Controllers\Campaigns\AmazonSpBudgetController;
use Illuminate\Console\Command;
use App\Models\AmazonSpCampaignReport;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Support\Facades\Log;

class AutoUpdateAmazonFbaOverKwBids extends Command
{
    protected $signature = 'amazon-fba:auto-update-over-kw-bids';
    protected $description = 'Automatically update Amazon FBA campaign keyword bids';

    protected $profileId;

    public function __construct()
    {
        parent::__construct();
    }

    public function handle()
    {
        $this->info("Starting Amazon bids auto-update...");

        $updateKwBids = new AmazonSpBudgetController;

        $campaigns = $this->getAutomateAmzUtilizedBgtKw();

        if (empty($campaigns)) {
            $this->warn("No campaigns matched filter conditions.");
            return 0;
        }

        $campaignIds = collect($campaigns)->pluck('campaign_id')->toArray();
        $newBids = collect($campaigns)->pluck('sbid')->toArray();

        $result = $updateKwBids->updateAutoCampaignKeywordsBid($campaignIds, $newBids);
        $this->info("Update Result: " . json_encode($result));

    }

    public function getAutomateAmzUtilizedBgtKw()
    {
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        $amazonSpCampaignReportsL7 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L7')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
                }
            })
            ->where(function ($q) {
                $q->where('campaignName', 'LIKE', '%FBA%')
                ->orWhere('campaignName', 'LIKE', '%fba%')
                ->orWhere('campaignName', 'LIKE', '%FBA.%')
                ->orWhere('campaignName', 'LIKE', '%fba.%');
            })
            ->whereRaw("LOWER(TRIM(TRAILING '.' FROM campaignName)) NOT LIKE '% pt'")
            ->get();

        $amazonSpCampaignReportsL1 = AmazonSpCampaignReport::where('ad_type', 'SPONSORED_PRODUCTS')
            ->where('report_date_range', 'L1')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . $sku . '%');
                }
            })
            ->where(function ($q) {
                $q->where('campaignName', 'LIKE', '%FBA%')
                ->orWhere('campaignName', 'LIKE', '%fba%')
                ->orWhere('campaignName', 'LIKE', '%FBA.%')
                ->orWhere('campaignName', 'LIKE', '%fba.%');
            })
            ->whereRaw("LOWER(TRIM(TRAILING '.' FROM campaignName)) NOT LIKE '% pt'")
            ->get();


        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);

            $shopify = $shopifyData[$pm->sku] ?? null;

            $matchedCampaignL7 = $amazonSpCampaignReportsL7->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim(rtrim($item->campaignName, '.')));

                return (
                    (str_ends_with($cleanName, $sku . ' FBA') || str_ends_with($cleanName, $sku . ' FBA.'))
                    && !str_ends_with($cleanName, $sku . ' PT')
                    && !str_ends_with($cleanName, $sku . ' PT.')
                );
            });

            $matchedCampaignL1 = $amazonSpCampaignReportsL1->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim(rtrim($item->campaignName, '.')));

                return (
                    (str_ends_with($cleanName, $sku . ' FBA') || str_ends_with($cleanName, $sku . ' FBA.'))
                    && !str_ends_with($cleanName, $sku . ' PT')
                    && !str_ends_with($cleanName, $sku . ' PT.')
                );
            });

            $row = [];
            $row['INV']    = $shopify->inv ?? 0;
            $row['campaign_id'] = $matchedCampaignL7->campaign_id ?? ($matchedCampaignL1->campaign_id ?? '');
            $row['campaignName'] = $matchedCampaignL7->campaignName ?? ($matchedCampaignL1->campaignName ?? '');
            $row['campaignBudgetAmount'] = $matchedCampaignL7->campaignBudgetAmount ?? ($matchedCampaignL1->campaignBudgetAmount ?? '');
            $row['l7_spend'] = $matchedCampaignL7->spend ?? 0;
            $row['l1_cpc'] = $matchedCampaignL1->costPerClick ?? 0;

            $l1_cpc = floatval($row['l1_cpc']);
            $row['sbid'] = floor($l1_cpc * 0.90 * 100) / 100;

            $budget = floatval($row['campaignBudgetAmount']);
            $l7_spend = floatval($row['l7_spend']);

            $ub7 = $budget > 0 ? ($l7_spend / ($budget * 7)) * 100 : 0;

            if($row['campaignName'] != '' && $ub7 > 90) {
                $result[] = (object) $row;
            }
        }
        return $result;
    }
}