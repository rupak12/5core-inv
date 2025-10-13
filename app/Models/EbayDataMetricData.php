<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class EbayDataMetricData extends Model
{
    use HasFactory;

    protected $table = 'fetch_api_for_ebay_data_metric_data';

    protected $fillable = [
        'item_id',
        'sku',
        'ebay_data_price',
        'ebay_data_l30',
        'ebay_data_l60',
        'ebay_data_views',
    ];

    protected $casts = [
        'ebay_data_price' => 'decimal:2',
    ];

    /**
     * Get the product master associated with this metric data
     */
    public function productMaster()
    {
        return $this->belongsTo(ProductMaster::class, 'sku', 'sku');
    }

    /**
     * Scope to get data for specific SKUs
     */
    public function scopeForSkus($query, array $skus)
    {
        return $query->whereIn('sku', $skus);
    }

    /**
     * Scope to get data for specific item IDs
     */
    public function scopeForItemIds($query, array $itemIds)
    {
        return $query->whereIn('item_id', $itemIds);
    }

    /**
     * Get the latest data entry for this item
     */
    public function scopeLatest($query)
    {
        return $query->orderBy('created_at', 'desc');
    }
}