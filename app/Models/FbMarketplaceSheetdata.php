<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class FbMarketplaceSheetdata extends Model
{
    use HasFactory;

    protected $table = 'fb_marketplace_sheet_data';

    protected $fillable = [
        'sku',
        'l30',
        'l60',
        'price',
        'views',
    ];
}
