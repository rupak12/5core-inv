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
        Schema::create('wmp_mark_as_dones', function (Blueprint $table) {
            $table->id();
            $table->string('parent')->nullable();
            $table->string('sku')->unique();
            $table->date('done_date')->nullable();
            $table->boolean('is_done')->default(false);
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('wmp_mark_as_dones');
    }
};
