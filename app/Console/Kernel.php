<?php

namespace App\Console;

use Illuminate\Console\Scheduling\Schedule;
use Illuminate\Foundation\Console\Kernel as ConsoleKernel;
use App\Console\Commands\FetchReverbData;
use App\Console\Commands\FetchMacyProducts;
use App\Console\Commands\FetchWayfairData;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;

class Kernel extends ConsoleKernel
{
    protected $commands = [
        FetchReverbData::class,
        FetchMacyProducts::class,
        FetchWayfairData::class,
        \App\Console\Commands\LogClear::class,
        \App\Console\Commands\SyncTemuSheet::class,
        \App\Console\Commands\AutoUpdateAmazonKwBids::class,
        \App\Console\Commands\AutoUpdateAmazonPtBids::class,
        \App\Console\Commands\AutoUpdateAmazonHlBids::class,
        \App\Console\Commands\AutoUpdateAmzUnderKwBids::class,
        \App\Console\Commands\AutoUpdateAmzUnderPtBids::class,
        \App\Console\Commands\AutoUpdateAmazonBgtKw::class,
        \App\Console\Commands\AutoUpdateAmazonBgtPt::class,
        \App\Console\Commands\AutoUpdateAmazonBgtHl::class,
        \App\Console\Commands\AutoUpdateAmazonPinkDilKwAds::class,
        \App\Console\Commands\AutoUpdateAmazonPinkDilPtAds::class,
        \App\Console\Commands\AutoUpdateAmazonPinkDilHlAds::class,
        \App\Console\Commands\EbayOverUtilzBidsAutoUpdate::class,
        \App\Console\Commands\EbayPinkDilKwBidsAutoUpdate::class,
        \App\Console\Commands\EbayPriceLessBidsAutoUpdate::class,
        \App\Console\Commands\AutoUpdateAmazonFbaOverKwBids::class,
        \App\Console\Commands\AutoUpdateAmazonFbaUnderKwBids::class,
        \App\Console\Commands\AutoUpdateAmazonFbaOverPtBids::class,
        \App\Console\Commands\AutoUpdateAmazonFbaUnderPtBids::class,

    ];

    /**
     * Define the application's command schedule.
     */
    protected function schedule(Schedule $schedule)
    {
        // Test scheduler to verify it's working
        $schedule->call(function () {
            Log::info('Test scheduler is working at ' . now());
        })->everyMinute()->name('test-scheduler-log');

        // Clear Laravel log after test log
        $schedule->call(function () {
            $logPath = storage_path('logs/laravel.log');
            if (file_exists($logPath)) {
                file_put_contents($logPath, '');
            }
        })->everyFiveMinutes()->name('clear-laravel-log');

        // All commands running every 5 minutes
        $schedule->command('shopify:save-daily-inventory')
            ->everyFiveMinutes()
            ->timezone('UTC');
        $schedule->command('app:process-jungle-scout-sheet-data')
            ->dailyAt('00:30')
            ->timezone('America/Los_Angeles');
        $schedule->command('app:fetch-amazon-listings')
            ->dailyAt('06:00')
            ->timezone('America/Los_Angeles');
        $schedule->command('reverb:fetch')
            ->everyFiveMinutes()
            ->timezone('UTC');
        $schedule->command('app:fetch-ebay-reports')
            ->hourly()
            ->timezone('UTC');
        $schedule->command('app:fetch-macy-products')
            ->everyFiveMinutes()
            ->timezone('UTC');
        $schedule->command('app:fetch-wayfair-data')
            ->everyFiveMinutes()
            ->timezone('UTC');
        $schedule->command('app:amazon-campaign-reports')
            ->dailyAt('04:00')
            ->timezone('UTC');
        $schedule->command('app:ebay-campaign-reports')
            ->dailyAt('05:00')
            ->timezone('UTC');
        $schedule->command('app:fetch-doba-metrics')
            ->dailyAt('00:00')
            ->timezone('UTC');

        // Sync Main sheet update command
        $schedule->command('app:sync-sheet')
            ->dailyAt('02:10')
            ->timezone('UTC');
        // Sync Temu sheet command
      
        // Sync Newegg sheet command
        $schedule->command('sync:neweegg-sheet')->everyTenMinutes();
        // Sync Wayfair sheet command
        $schedule->command('sync:wayfair-sheet')->everyTenMinutes();

        $schedule->command('sync:shein-sheet')->daily();

        // Sync Walmart sheet command
        $schedule->command('sync:walmart-sheet')->everyTenMinutes();
        $schedule->command('sync:temu-sheet-data')->everyTwelveHours();


        // // Sync eBay 2 sheet command
        // $schedule->command('sync:ebay-two-sheet')->everyTenMinutes();
        // // Sync eBay 3 sheet command
        // $schedule->command('sync:ebay-three-sheet')->everyTenMinutes();

        // Sync Shopify sheet command
        $schedule->command('sync:shopify-quantity')->everyTenMinutes()
            ->timezone('UTC');   

        $schedule->command('app:fetch-ebay-three-metrics')
            ->dailyAt('02:00')
            ->timezone('America/Los_Angeles');
            
        $schedule->command('app:ebay3-campaign-reports')
            ->dailyAt('04:00')
            ->timezone('America/Los_Angeles');
        $schedule->command('app:fetch-temu-metrics')
            ->dailyAt('03:00')
            ->timezone('America/Los_Angeles');
        $schedule->command('app:fetch-ebay-two-metrics')
            ->dailyAt('01:00')
            ->timezone('America/Los_Angeles');   
        $schedule->command('app:ebay2-campaign-reports')
            ->dailyAt('01:15')
            ->timezone('America/Los_Angeles');
        // Amazon over utilized bids update commands
        $schedule->command('amazon:auto-update-over-kw-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('amazon:auto-update-over-pt-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata'); 
        $schedule->command('amazon:auto-update-over-hl-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        // amazon acos bgt update commands
        $schedule->command('amazon:auto-update-amz-bgt-kw')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata'); 
        $schedule->command('amazon:auto-update-amz-bgt-pt')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata'); 
        $schedule->command('amazon:auto-update-amz-bgt-hl')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata'); 
        // Pink Dil ads update command
        $schedule->command('amazon:auto-update-pink-dil-kw-ads')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('amazon:auto-update-pink-dil-pt-ads')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('amazon:auto-update-pink-dil-hl-ads')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        // FBA bids update command

        $schedule->command('amazon-fba:auto-update-over-kw-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('amazon-fba:auto-update-under-kw-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('amazon-fba:auto-update-over-pt-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('amazon-fba:auto-update-under-pt-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');

        // Ebay bids update command
        $schedule->command('ebay:auto-update-over-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('ebay:auto-update-pink-dil-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        $schedule->command('ebay:auto-update-price-less-bids')
            ->dailyAt('12:00')
            ->timezone('Asia/Kolkata');
        // end of bids update commands
        $schedule->command('sync:amazon-prices')->everyMinute();
        $schedule->command('sync:sync-temu-sip')->everyMinute();
        $schedule->command('sync:walmart-metrics-data')->everyMinute();
        $schedule->command('sync:tiktok-sheet-data')->everyMinute();
        $schedule->command('app:aliexpress-sheet-sync')->everyMinute();
        $schedule->command('app:fetch-ebay-table-data')->dailyAt('00:00');
          $schedule->call(function () {
            DB::connection('apicentral')
                ->table('google_ads_campaigns')
                ->where('id', 1)
                ->update(['sbid_status' => 0]);
        })->dailyAt('00:00');
        $schedule->command('sbid:update')
            ->dailyAt('00:01') 
            ->timezone('Asia/Kolkata');
    }

    /**
     * Register the commands for the application.
     */
    protected function commands(): void
    {
        $this->load(__DIR__ . '/Commands');
        require base_path('routes/console.php');
    }
}
 