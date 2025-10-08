<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class FbShopSheetdata extends Model
{
    use HasFactory;

     protected $table = 'fb_shop_sheet_data';

    protected $fillable = [
        'sku',
        'l30',
        'l60',
        'price',
        'views',
    ];
}
