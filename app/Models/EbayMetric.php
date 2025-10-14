<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class EbayMetric extends Model
{
    use HasFactory;

    protected $table = 'ebay_metrics';

    protected $fillable = [
        'id',
        'item_id',
        'sku',
        'ebay_price',
        'ebay_l30',
        'ebay_l60',
        'ebay_views',
    ];

}
