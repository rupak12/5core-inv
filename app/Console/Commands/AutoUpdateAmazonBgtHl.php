<?php

namespace App\Console\Commands;

use App\Http\Controllers\Campaigns\AmazonSpBudgetController;
use App\Http\Controllers\MarketPlace\ACOSControl\AmazonACOSController;
use App\Models\AmazonDatasheet;
use App\Models\AmazonDataView;
use App\Models\AmazonSbCampaignReport;
use Illuminate\Console\Command;
use App\Models\AmazonSpCampaignReport;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use GuzzleHttp\Client;
use Illuminate\Support\Facades\Log;

class AutoUpdateAmazonBgtHl extends Command
{
    protected $signature = 'amazon:auto-update-amz-bgt-hl';
    protected $description = 'Automatically update Amazon campaign bgt price';

    protected $profileId;

    public function __construct()
    {
        parent::__construct();
    }

    public function handle()
    {
        $this->info("Starting Amazon bgts auto-update...");

        $updateKwBgts = new AmazonACOSController;

        $campaigns = $this->amazonAcosHlControlData();

        if (empty($campaigns)) {
            $this->warn("No campaigns matched filter conditions.");
            return 0;
        }

        $campaignIds = collect($campaigns)->pluck('campaign_id')->toArray();
        $newBgts = collect($campaigns)->pluck('sbgt')->toArray();
        
        $result = $updateKwBgts->updateAutoAmazonSbCampaignBgt($campaignIds, $newBgts);
        $this->info("Update Result: " . json_encode($result));

    }

    public function amazonAcosHlControlData()
    {
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->filter()->unique()->values()->all();

        $nrValues = AmazonDataView::whereIn('sku', $skus)->pluck('value', 'sku');
        
        $amazonSpCampaignReportsL30 = AmazonSbCampaignReport::where('ad_type', 'SPONSORED_BRANDS')
            ->where('report_date_range', 'L30')
            ->where(function ($q) use ($skus) {
                foreach ($skus as $sku) {
                    $q->orWhere('campaignName', 'LIKE', '%' . strtoupper($sku) . '%');
                }
            })
            ->get();

        $result = [];

        foreach ($productMasters as $pm) {
            $sku = strtoupper($pm->sku);

            $matchedCampaignL30 = $amazonSpCampaignReportsL30->first(function ($item) use ($sku) {
                $cleanName = strtoupper(trim($item->campaignName));
                $expected1 = $sku;                
                $expected2 = $sku . ' HEAD';      

                return ($cleanName === $expected1 || $cleanName === $expected2)
                    && strtoupper($item->campaignStatus) === 'ENABLED';
            });

            if (!$matchedCampaignL30) {
                continue;
            }

            // clicks must be >= 25
            if (($matchedCampaignL30->clicks ?? 0) < 25) {
                continue;
            }

            $row = [];
            $row['campaign_id'] = $matchedCampaignL30->campaign_id ?? '';
            $row['campaign_name'] = $matchedCampaignL30->campaignName ?? '';
            $row['acos_L30'] = ($matchedCampaignL30 && ($matchedCampaignL30->sales ?? 0) > 0)
                ? round(($matchedCampaignL30->cost / $matchedCampaignL30->sales) * 100, 2)
                : 0;

            $row['spend_l30']       = $matchedCampaignL30->spend ?? 0;
            $row['ad_sales_l30']    = $matchedCampaignL30->sales30d ?? 0;

            $acos = (float) ($row['acos_L30'] ?? 0);

            $tpft = 0;
            if (isset($nrValues[$pm->sku])) {
                $raw = $nrValues[$pm->sku];
                if (!is_array($raw)) $raw = json_decode($raw, true);
                if (is_array($raw)) $tpft = isset($raw['TPFT']) ? (int) floor($raw['TPFT']) : 0;
            }
            $row['TPFT'] = $tpft;

            $acos = (float) ($row['acos_L30'] ?? 0);
            $spend = (float) ($row['spend_l30'] ?? 0);
            $sales = (float) ($row['ad_sales_l30'] ?? 0);

            if($spend > 0 && $sales === 0) {
                $sbgt = 1;
            }else if ($acos < 10) {
                $sbgt = 10;        
            } else if ($acos >= 10 && $acos < 15) {
                $sbgt = 9;         
            } else if ($acos >= 15 && $acos < 20) {
                $sbgt = 8;         
            } else if ($acos >= 20 && $acos <= 50) {
                $sbgt = 2;        
            } else {  
                $sbgt = 1;
            }

            // $l30 = (float) ($shopify->quantity ?? 0);
            // $inv = (float) ($shopify->inv ?? 0);
            // $dilColor = "";
            // if ($inv != 0) {
            //     $dilDecimal = $l30 / $inv;
            //     $dilColor = $this->getDilColor($dilDecimal);
            // }

            // if (($dilColor === "red" && $tpft > 18) ||
            //     ($dilColor === "yellow" && $tpft > 22) ||
            //     ($dilColor === "green" && $tpft > 26) ||
            //     ($dilColor === "pink" && $tpft > 30)) {
            //     $sbgt = $sbgt * 2;
            // }

            $row['sbgt'] = $sbgt;

            $result[] = (object) $row;
        }

        return $result;
    }
}