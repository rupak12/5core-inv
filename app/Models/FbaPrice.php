<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class FbaPrice extends Model
{
    protected $table = 'fba_prices';
    protected $fillable = ['seller_sku', 'price'];
}
