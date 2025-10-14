<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class WaifairProductSheet extends Model
{
    protected $table = 'wayfair_product_sheets'; // Ensure the model uses the correct table name
     protected $fillable = [
        'sku', 'price', 'pft', 'roi', 'l30', 'dil', 'buy_link','l60'
    ];
    use HasFactory;
}
