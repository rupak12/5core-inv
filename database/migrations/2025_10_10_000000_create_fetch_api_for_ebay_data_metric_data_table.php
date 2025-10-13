<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('fetch_api_for_ebay_data_metric_data', function (Blueprint $table) {
            $table->id();
            
            $table->string('item_id')->nullable();
            $table->string('sku')->nullable();
            
            // eBay data fields
            $table->decimal('ebay_price', 10, 2)->nullable();
            $table->integer('ebay_data_l30')->nullable();
            $table->integer('ebay_data_l60')->nullable();
            $table->integer('ebay_views')->nullable();
            
            $table->timestamps();
            
            // Indexes for better performance
            $table->index('item_id');
            $table->index('sku');
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('fetch_api_for_ebay_data_metric_data');
    }
};