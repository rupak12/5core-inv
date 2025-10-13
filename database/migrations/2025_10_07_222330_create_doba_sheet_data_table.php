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
        Schema::create('doba_sheet_data', function (Blueprint $table) {
            $table->id();
            $table->string('sku')->index();
            $table->integer('l30')->nullable();
            $table->integer('l60')->nullable();
            $table->decimal('price', 10, 2)->nullable();
            $table->integer('views')->nullable();
            $table->decimal('pickup_price', 10, 2)->nullable();
            $table->string('item_id')->nullable();
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('doba_sheet_data');
    }
};
