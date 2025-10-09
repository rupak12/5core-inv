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
        Schema::table('fba_fees', function (Blueprint $table) {
            $table->dropUnique(['seller_sku', 'report_generated_at']);
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::table('fba_fees', function (Blueprint $table) {
            $table->unique(['seller_sku', 'report_generated_at']);
        });
    }
};
