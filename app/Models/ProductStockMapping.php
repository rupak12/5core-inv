<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class ProductStockMapping extends Model
{
    use HasFactory;
    protected $fillable = ['image','variant_id','sku','title','inventory_shopify','inventory_shopify_product','inventory_amazon','inventory_amazon_product','not_required','is_parent','parent','product_id','inventory_walmart','inventory_reverb','inventory_shein','inventory_doba','inventory_temu','inventory_macy','inventory_ebay1','inventory_ebay2','inventory_ebay3'];
}
