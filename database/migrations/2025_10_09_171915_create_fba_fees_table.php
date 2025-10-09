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
        Schema::create('fba_fees', function (Blueprint $table) {
            $table->id();
            $table->string('seller_sku')->index();
            $table->string('fnsku')->nullable();
            $table->string('asin')->nullable()->index();
            $table->string('amazon_store')->nullable();
            $table->text('product_name')->nullable();
            $table->string('product_group')->nullable();
            $table->string('brand')->nullable();
            $table->string('fulfilled_by')->nullable();
            $table->decimal('your_price', 10, 2)->nullable();
            $table->decimal('sales_price', 10, 2)->nullable();
            $table->decimal('longest_side', 8, 2)->nullable();
            $table->decimal('median_side', 8, 2)->nullable();
            $table->decimal('shortest_side', 8, 2)->nullable();
            $table->decimal('length_and_girth', 8, 2)->nullable();
            $table->string('unit_of_dimension')->nullable();
            $table->decimal('item_package_weight', 8, 3)->nullable();
            $table->string('unit_of_weight')->nullable();
            $table->string('product_size_tier')->nullable();
            $table->string('currency', 3)->default('USD');
            $table->decimal('estimated_fee_total', 10, 2)->nullable();
            $table->decimal('estimated_referral_fee_per_unit', 10, 2)->nullable();
            $table->decimal('estimated_variable_closing_fee', 10, 2)->nullable();
            $table->decimal('estimated_fixed_closing_fee', 10, 2)->nullable();
            $table->decimal('estimated_order_handling_fee_per_order', 10, 2)->nullable();
            $table->decimal('estimated_pick_pack_fee_per_unit', 10, 2)->nullable();
            $table->decimal('estimated_weight_handling_fee_per_unit', 10, 2)->nullable();
            $table->decimal('expected_fulfillment_fee_per_unit', 10, 2)->nullable();
            $table->decimal('estimated_future_fee', 10, 2)->nullable();
            $table->decimal('estimated_future_order_handling_fee_per_order', 10, 2)->nullable();
            $table->decimal('estimated_future_pick_pack_fee_per_unit', 10, 2)->nullable();
            $table->decimal('estimated_future_weight_handling_fee_per_unit', 10, 2)->nullable();
            $table->decimal('expected_future_fulfillment_fee_per_unit', 10, 2)->nullable();
            $table->timestamp('report_generated_at')->nullable();
            $table->timestamps();
            
            $table->unique(['seller_sku', 'report_generated_at']);
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('fba_fees');
    }
};
