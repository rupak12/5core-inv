<?php

namespace App\Http\Controllers\Campaigns;

use App\Http\Controllers\Controller;
use App\Models\MarketplacePercentage;
use Illuminate\Http\Request;

class Ebay3PmtAdsController extends Controller
{
    public function index()
    {
        $marketplaceData = MarketplacePercentage::where("marketplace", "Ebay3" )->first();
        $ebayPercentage = $marketplaceData ? $marketplaceData->percentage : 100;
        $ebayAdPercentage = $marketplaceData ? $marketplaceData->ad_updates : 100;

        return view('campaign.ebay-three.pmt-ads', compact('ebayPercentage','ebayAdPercentage'));
    }
}
