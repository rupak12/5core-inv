<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class ProductStockMapping extends Model
{
    use HasFactory;
    protected $fillable = ['sku','title','inventory_shopify','inventory_shopify_product','inventory_amazon','inventory_amazon_product'];
}
