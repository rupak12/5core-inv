<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class TopDawgSheetdata extends Model
{
    use HasFactory;

    protected $table = 'top_dawg_sheet_data';

    protected $fillable = [
        'sku',
        'l30',
        'l60',
        'price',
        'views',
    ];
}
