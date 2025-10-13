<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class FbaReportsMaster extends Model
{
    protected $table = 'fba_reports_master';

    protected $fillable = [
        'seller_sku',
        'asin',
        'year',
        'total_views',
        'current_month_views',
        'fulfillment_fee',
        'referral_fee',
        'storage_fee',
        'total_fee',
    ];

    public $timestamps = true;
}
