<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class BusinessFiveCoreSheetdata extends Model
{
    use HasFactory;

    protected $table = 'business_five_core_sheet_data';

    protected $fillable = [
        'sku',
        'l30',
        'l60',
        'price',
        'views',
    ];
}
