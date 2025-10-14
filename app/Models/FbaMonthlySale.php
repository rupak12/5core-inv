<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class FbaMonthlySale extends Model
{
    protected $table = 'fba_monthly_sales';

    protected $fillable = [
        'seller_sku','asin','year',
        'jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec',
        'total_units','avg_price',
        'l30_units','l30_revenue','l60_units','l60_revenue'
    ];
}
