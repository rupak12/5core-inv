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
        if (!Schema::hasTable('negative_reviews')) {
        Schema::create('negative_reviews', function (Blueprint $table) {
            $table->id();
            $table->date('review_date');
            $table->string('marketplace');
            $table->string('sku');
            $table->tinyInteger('rating');
            $table->string('review_category')->nullable();
            $table->text('review_text')->nullable();
            $table->text('review_summary')->nullable();
            $table->string('reviewer_name')->nullable();
            $table->string('action_status')->default('Pending');
            $table->text('action_taken')->nullable();
            $table->date('action_date')->nullable();
            $table->timestamps();
        });
    }
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('negative_reviews');
    }
};
