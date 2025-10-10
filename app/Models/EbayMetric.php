<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class EbayMetric extends Model
{
    use HasFactory;

    protected $table = 'fetch_api_for_ebay_data_metric_data';

    protected $fillable = [
        'id',
        'item_id',
        'sku',
        'ebay_data_price',
        'ebay_data_l30',
        'ebay_data_l60',
        'ebay_data_views',
    ];

}
