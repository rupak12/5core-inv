<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Ebay3PriorityReport extends Model
{
    use HasFactory;

    protected $table = 'ebay_3_priority_reports';

    protected $fillable = [
        'report_range',
        'start_date',
        'end_date',
        'campaign_id',
        'campaign_name',
        'campaignBudgetAmount',
        'campaignStatus',
        'cpc_impressions',
        'cpc_clicks',
        'cpc_attributed_sales',
        'cpc_ctr',
        'cpc_ad_fees_listingsite_currency',
        'cpc_sale_amount_listingsite_currency',
        'cpc_avg_cost_per_sale',
        'cpc_return_on_ad_spend',
        'cpc_conversion_rate',
        'cpc_sale_amount_payout_currency',
        'cost_per_click',
        'cpc_ad_fees_payout_currency',
        'channels',
    ];
}
