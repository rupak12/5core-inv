<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class WalmartProductSheet extends Model
{
    protected $table = 'walmart_product_sheet'; // Ensure the model uses the correct table name
    protected $fillable = [
        'sku', 'price', 'pft', 'roi', 'l30', 'dil', 'buy_link'
    ];
    use HasFactory;
}
