<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class EbayMetric extends Model
{
    use HasFactory;

    protected $fillable = [
        'item_id',
        'sku',
        'ebay_l30',
        'ebay_l60',
        'ebay_price',
        'price_lmpa',
        'campaign_id',
        'views',
        'total_pft',
        't_sale_l30',
        'pft_percentage',
        'roi_percentage',
        't_cogs',
    ];
}
