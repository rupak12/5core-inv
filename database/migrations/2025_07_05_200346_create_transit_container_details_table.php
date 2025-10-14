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
        if (!Schema::hasTable('transit_container_details')) {
            Schema::create('transit_container_details', function (Blueprint $table) {
                $table->id();
                $table->string('tab_name'); 
                $table->string('supplier_name')->nullable();
                $table->string('company_name')->nullable();
                $table->string('parent')->nullable();
                $table->string('our_sku')->nullable();
                $table->text('photos')->nullable(); 
                $table->text('specification')->nullable();
                $table->string('package_size')->nullable();
                $table->string('product_size_link')->nullable();
                $table->string('status')->nullable(); 
                $table->string('changes')->nullable(); 
                $table->string('no_of_units')->nullable();
                $table->string('total_ctn')->nullable();
                $table->decimal('rate', 10, 3)->nullable();
                $table->string('unit')->nullable();
                $table->string('order_link')->nullable();
                $table->string('comparison_link')->nullable();
                $table->timestamps();
            });
        }
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('transit_container_details');
    }
};
