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
        Schema::table('ebay_3_priority_reports', function (Blueprint $table) {
            $table->date('start_date')->nullable()->after('report_range');
            $table->date('end_date')->nullable()->after('start_date');
            $table->string('campaign_name')->nullable()->after('end_date');
            $table->decimal('campaignBudgetAmount', 15, 2)->nullable()->after('campaign_name');
            $table->string('campaignStatus')->nullable()->after('campaignBudgetAmount');
        });
    }

    public function down(): void
    {
        Schema::table('ebay_3_priority_reports', function (Blueprint $table) {
            $table->dropColumn(['start_date', 'end_date', 'campaign_name', 'campaignBudgetAmount', 'campaignStatus']);
        });
    }
};
