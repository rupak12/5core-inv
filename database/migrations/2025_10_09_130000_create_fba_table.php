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
        Schema::create('fba_table', function (Blueprint $table) {
            $table->id();
            $table->string('seller_sku');
            $table->string('fulfillment_channel_sku')->nullable();
            $table->string('asin')->nullable();
            $table->string('condition_type')->default('NewItem');
            $table->integer('quantity_available')->default(0);
            $table->decimal('price', 10, 2)->nullable();
            $table->timestamps();
            
            // Add indexes for better performance
            $table->index('seller_sku');
            $table->index('asin');
            $table->index('quantity_available');
            $table->index('price');
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('fba_table');
    }
};