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
        Schema::table('fba_monthly_sales', function (Blueprint $table) {
            $table->integer('l30_units')->default(0);
            $table->decimal('l30_revenue', 10, 2)->nullable();
            $table->integer('l60_units')->default(0);
            $table->decimal('l60_revenue', 10, 2)->nullable();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::table('fba_monthly_sales', function (Blueprint $table) {
            $table->dropColumn(['l30_units', 'l30_revenue', 'l60_units', 'l60_revenue']);
        });
    }
};
