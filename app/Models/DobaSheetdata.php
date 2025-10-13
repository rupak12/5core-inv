<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class DobaSheetdata extends Model
{
    use HasFactory;

    protected $table = 'doba_sheet_data';

    protected $fillable = [
        'sku',
        'item_id',
        'l30',
        'l60',
        'price',
        'views',
        'pickup_price',
    ];
}
