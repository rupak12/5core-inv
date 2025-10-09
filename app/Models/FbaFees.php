<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class FbaFees extends Model
{
    use HasFactory;

    protected $fillable = [
        'seller_sku',
        'fnsku',
        'asin',
        'amazon_store',
        'product_name',
        'product_group',
        'brand',
        'fulfilled_by',
        'your_price',
        'sales_price',
        'longest_side',
        'median_side',
        'shortest_side',
        'length_and_girth',
        'unit_of_dimension',
        'item_package_weight',
        'unit_of_weight',
        'product_size_tier',
        'currency',
        'estimated_fee_total',
        'estimated_referral_fee_per_unit',
        'estimated_variable_closing_fee',
        'estimated_fixed_closing_fee',
        'estimated_order_handling_fee_per_order',
        'estimated_pick_pack_fee_per_unit',
        'estimated_weight_handling_fee_per_unit',
        'expected_fulfillment_fee_per_unit',
        'estimated_future_fee',
        'estimated_future_order_handling_fee_per_order',
        'estimated_future_pick_pack_fee_per_unit',
        'estimated_future_weight_handling_fee_per_unit',
        'expected_future_fulfillment_fee_per_unit',
        'report_generated_at'
    ];

    protected $casts = [
        'your_price' => 'decimal:2',
        'sales_price' => 'decimal:2',
        'longest_side' => 'decimal:2',
        'median_side' => 'decimal:2',
        'shortest_side' => 'decimal:2',
        'length_and_girth' => 'decimal:2',
        'item_package_weight' => 'decimal:3',
        'estimated_fee_total' => 'decimal:2',
        'estimated_referral_fee_per_unit' => 'decimal:2',
        'estimated_variable_closing_fee' => 'decimal:2',
        'estimated_fixed_closing_fee' => 'decimal:2',
        'estimated_order_handling_fee_per_order' => 'decimal:2',
        'estimated_pick_pack_fee_per_unit' => 'decimal:2',
        'estimated_weight_handling_fee_per_unit' => 'decimal:2',
        'expected_fulfillment_fee_per_unit' => 'decimal:2',
        'estimated_future_fee' => 'decimal:2',
        'estimated_future_order_handling_fee_per_order' => 'decimal:2',
        'estimated_future_pick_pack_fee_per_unit' => 'decimal:2',
        'estimated_future_weight_handling_fee_per_unit' => 'decimal:2',
        'expected_future_fulfillment_fee_per_unit' => 'decimal:2',
        'report_generated_at' => 'datetime'
    ];
}
