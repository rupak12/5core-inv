<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class FbaTable extends Model
{
    use HasFactory;

    protected $table = 'fba_table';

    protected $fillable = [
        'seller_sku',
        'fulfillment_channel_sku',
        'asin',
        'condition_type',
        'quantity_available',
        // 'price'
    ];

    protected $casts = [
        'quantity_available' => 'integer',
        // 'price' => 'decimal:2'
    ];
}