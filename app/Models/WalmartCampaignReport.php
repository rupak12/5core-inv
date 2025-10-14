<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class WalmartCampaignReport extends Model
{
    use HasFactory;

    protected $table = 'walmart_campaign_reports'; 

    public $timestamps = true; 

    protected $fillable = [
        'report_range',
        'campaign_id',
        'campaignName',
        'budget',
        'spend',
        'cpc',
        'impression',
        'clicks',
        'status',
        'created_at',
        'updated_at',
    ];
}
