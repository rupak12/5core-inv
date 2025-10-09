<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use App\Services\FetchEbayDataMetricServiceClass;

class FetchApiForEbayDataMetricData extends Command
{
    protected $signature = 'app:fetch-ebay-table-data';
    protected $description = 'Fetch eBay data metrics and insert into database';

    public function handle(FetchEbayDataMetricServiceClass $service)
    {
        $this->info('🔄 Starting eBay data metric fetching...');

        if ($service->fetchAndInsertEbayMetrics()) {
            $this->info('✅ eBay data metrics inserted successfully!');
        } else {
            $this->error('❌ Failed to fetch or insert eBay metrics.');
        }
    }
}
