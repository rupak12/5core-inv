<?php

namespace App\Http\Controllers\MarketingMaster;

use App\Http\Controllers\Controller;
use App\Http\Controllers\ApiController;
use App\Http\Controllers\MarketPlace\AmazonZeroController;
use App\Http\Controllers\MarketPlace\EbayZeroController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingAmazonController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingEbayController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingTemuController;
use App\Http\Controllers\MarketPlace\MacyZeroController;
use App\Http\Controllers\MarketPlace\Neweggb2cZeroController;
use App\Http\Controllers\MarketPlace\Shopifyb2cZeroController;
use App\Http\Controllers\MarketPlace\TemuZeroController;
use App\Http\Controllers\MarketPlace\WayfairZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\AliexpressZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\DobaZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\Ebay2ZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\Ebay3ZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\SheinZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\TiktokShopZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\WalmartZeroController;
use App\Models\AliexpressDataView;
use App\Models\AmazonDataView;
use App\Models\DobaDataView;
use App\Models\EbayDataView;
use App\Models\EbayThreeDataView;
use App\Models\EbayTwoDataView;
use App\Models\MacyDataView;
use App\Models\MarketplacePercentage;
use App\Models\ZeroVisibilityMaster;
use App\Models\ProductMaster;
use App\Models\SheinDataView;
use App\Models\Shopifyb2cDataView;
use App\Models\ShopifySku;
use App\Models\TemuDataView;
use App\Models\TiktokShopDataView;
use App\Models\TiktokVideoAd;
use App\Models\WalmartDataView;
use App\Models\WayfairDataView;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\DB;


class ZeroVisibilityMasterController extends Controller
{

    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }


    /**
     * Display a listing of the resource.
     */
    public function index()
    {
        $productSKUs = ProductMaster::where('sku', 'NOT LIKE', '%PARENT%')
            ->pluck('sku')
            ->toArray();

        $zeroInvCount = ShopifySku::whereIn('sku', $productSKUs)
            ->where('inv', '<=', 0)
            ->count();


        $totalSkuCount = count($productSKUs);

        // --- Get eBay zero view count ---
        // $ebayZeroCount = app(EbayZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', 'eBay')
        //     ->update(['zero_visibility_sku_count' => $ebayZeroCount]);

        // // --- Get Amazon zero view count ---
        // $amazonZeroCount = app(AmazonZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', 'Amazon')
        //     ->update(['zero_visibility_sku_count' => $amazonZeroCount]);

        // // --- Get Shopify B2C zero view count ---
        // $shopifyB2CZeroCount = app(Shopifyb2cZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', 'Shopify B2C')
        //     ->update(['zero_visibility_sku_count' => $shopifyB2CZeroCount]);

        // // --- Get Macy's zero view count ---
        // $macyZeroCount = app(MacyZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', 'Macys')
        //     ->update(['zero_visibility_sku_count' => $macyZeroCount]);

        // // --- Get Newegg B2C zero view count ---
        // $neweggB2CZeroCount = app(Neweggb2cZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', `Newegg B2C`)
        //     ->update(['zero_visibility_sku_count' => $neweggB2CZeroCount]);

        // // --- Get Wayfair zero view count ---
        // $wayfairZeroCount = app(WayfairZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', 'Wayfair')
        //     ->update(['zero_visibility_sku_count' => $wayfairZeroCount]);

        // // --- Get Temu zero view count ---
        // $temuZeroCount = app(TemuZeroController::class)->getZeroViewCount();
        // ZeroVisibilityMaster::where('channel_name', 'Temu')
        //     ->update(['zero_visibility_sku_count' => $temuZeroCount]);

        $channels = ZeroVisibilityMaster::all();

        foreach ($channels as $channel) {
            $livePending = null;
            $zeroView = null;

            // Build controller class name dynamically (e.g., "Amazon" => AmazonZeroController)
            $baseName = str_replace([' ', '&', '-', '/'], '', ucwords(strtolower(trim($channel->channel_name))));
            $controllerClass = "App\\Http\\Controllers\\MarketPlace\\{$baseName}ZeroController";

            if (class_exists($controllerClass)) {
                $controller = app($controllerClass);
                if (method_exists($controller, 'getLivePendingAndZeroViewCounts')) {
                    $counts = $controller->getLivePendingAndZeroViewCounts();
                    $livePending = $counts['live_pending'] ?? null;
                    $zeroView = $counts['zero_view'] ?? null;
                }
            }

            $channel->live_pending = $livePending;
            $channel->zero_view = $zeroView;
        }

        return view('marketing-masters.zero-visibility-master', compact('totalSkuCount', 'zeroInvCount','channels'));
    }

    /**
     * Show the form for creating a new resource.
     */
    public function create()
    {
        //
    }

    /**
     * Store a newly created resource in storage.
     */
    public function store(Request $request)
    {
        $data = $request->all();

        // Store or update based on channel name
        $record = ZeroVisibilityMaster::updateOrCreate(
            ['channel_name' => $data['channel_name']],
            [
                'sheet_link' => $data['sheet_link'] ?? null,
                'is_ra_checked' => $data['is_ra_checked'] ?? false,
                'total_sku' => $data['total_sku'] ?? 0,
                'nr' => $data['nr'] ?? 0,
                'listed_req' => $data['listed_req'] ?? 0,
                'listed' => $data['listed'] ?? 0,
                'listing_pending' => $data['listing_pending'] ?? 0,
                'zero_inv' => $data['zero_inv'] ?? 0,
                'live_req' => $data['live_req'] ?? 0,
                'active_and_live' => $data['active_and_live'] ?? 0,
                'live_pending' => $data['live_pending'] ?? 0,
                'zero_visibility_sku_count' => $data['zero_visibility_sku_count'] ?? 0,
                'reason' => $data['reason'] ?? '',
                'step_taken' => $data['step_taken'] ?? '',
            ]
        );

        return response()->json(['message' => 'Saved successfully']);

    }

    /**
     * Display the specified resource.
     */
    public function show(string $id)
    {
        //
    }

    /**
     * Show the form for editing the specified resource.
     */
    public function edit(string $id)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     */
    public function update(Request $request, string $id)
    {
        $row = ZeroVisibilityMaster::findOrFail($request->id);
        $row->update($request->except('id'));
        return response()->json(['status' => true]);
    }

    /**
     * Remove the specified resource from storage.
     */
    public function destroy(string $id)
    {
        //
    }


    public function getMergedChannelData(Request $request)
    {
        // Get all channels from DB
        $channels = MarketplacePercentage::pluck('marketplace')->toArray();

        $data = [];


        // Mapping for special channel/controller names
        $controllerMap = [
            'ebaythree' => 'Ebay3ZeroController',
            'ebaytwo'   => 'Ebay2ZeroController',
            'tiktokshop' => 'TiktokShopZeroController',
            'doba'      => 'DobaZeroController',
            'walmart'      => 'WalmartZeroController',
            'shein'      => 'SheinZeroController',
            'bestbuyusa' => 'BestbuyUSAZeroController',
            'aliexpress' => 'AliexpressZeroController',
            // Add more mappings as needed
        ];

        foreach ($channels as $channel) {
            $livePending = 0;
            $zeroViews = 0;

            $key = strtolower(str_replace([' ', '-', '&', '/'], '', trim($channel)));
            if (isset($controllerMap[$key])) {
                $controllerClass = "App\\Http\\Controllers\\MarketPlace\\ZeroViewMarketPlace\\" . $controllerMap[$key];
            } else {
                $controllerClass = "App\\Http\\Controllers\\MarketPlace\\" . ucfirst($channel) . "ZeroController";
            }

            if (class_exists($controllerClass)) {
                $controller = app($controllerClass);
                // Try getLivePendingAndZeroViewCounts (preferred, returns both counts)
                if (method_exists($controller, 'getLivePendingAndZeroViewCounts')) {
                    $counts = $controller->getLivePendingAndZeroViewCounts();
                    $livePending = $counts['live_pending'] ?? 0;
                    $zeroViews = $counts['zero_view'] ?? 0;
                }
                // Fallback: try getZeroViewCount (returns only zero view count)
                else if (method_exists($controller, 'getZeroViewCount')) {
                    $zeroViews = $controller->getZeroViewCount();
                }
                // Fallback: try getZeroViewCounts (plural, some controllers use this)
                else if (method_exists($controller, 'getZeroViewCounts')) {
                    $zeroViews = $controller->getZeroViewCounts();
                }
            }

            $data[] = [
                'Channel ' => $channel,          // keep space to match your DataTable
                'R&A' => false,                  // placeholder
                'Live Pending' => $livePending,
                'Zero Visibility SKU Count' => $zeroViews,
            ];
        }

        return response()->json(['data' => $data]);
    }


    public function exportCsv()
    {
        $sheetResponse = (new ApiController)->fetchDataFromChannelMasterGoogleSheet();
        if ($sheetResponse->getStatusCode() !== 200) {
            return response()->json(['error' => 'Failed to fetch Google Sheet'], 500);
        }

        $sheetData = $sheetResponse->getData()->data ?? [];
        $dbRecords = ZeroVisibilityMaster::all()->keyBy(fn($row) => strtolower(trim($row->channel_name)));

        $headers = [
            'Content-Type' => 'text/csv',
            'Content-Disposition' => 'attachment; filename="zero_visibility_master_data.csv"',
        ];

        $columns = [
            'SL',
            'Channel Name',
            'R&A',
            'URL LINK',
            'Total SKU',
            'NR',
            'Listed Req',
            'Listed',
            'Listing Pending',
            'Zero Inv',
            'Live Req',
            'Active & Live',
            'Live Pending',
            'Zero Visibility SKU Count',
            'Reason',
            'Step Taken',
        ];

        $callback = function () use ($sheetData, $dbRecords, $columns) {
            $file = fopen('php://output', 'w');
            fputcsv($file, $columns);

            $sl = 1;
            foreach ($sheetData as $item) {
                $channelName = trim($item->{'Channel '} ?? '');
                if (!$channelName)
                    continue;

                $lower = strtolower($channelName);
                $dbRow = $dbRecords[$lower] ?? null;

                fputcsv($file, [
                    $sl++,
                    $channelName,
                    trim($item->{'R&A'} ?? ''),
                    // trim($item->{'URL LINK'} ?? ''),
                    $dbRow->sheet_link ?? '',
                    $dbRow->total_sku ?? '',
                    $dbRow->nr ?? '',
                    $dbRow->listed_req ?? '',
                    $dbRow->listed ?? '',
                    $dbRow->listing_pending ?? '',
                    $dbRow->zero_inv ?? '',
                    $dbRow->live_req ?? '',
                    $dbRow->active_and_live ?? '',
                    $dbRow->live_pending ?? '',
                    $dbRow->zero_visibility_sku_count ?? '',
                    $dbRow->reason ?? '',
                    $dbRow->step_taken ?? '',
                ]);
            }

            fclose($file);
        };

        return response()->stream($callback, 200, $headers);
    }


    public function updateRaCheckbox(Request $request)
    {
        $channel = trim($request->input('channel'));
        $checked = $request->input('checked') ? true : false;

        Log::info('Received update-checkbox request', [
            'channel' => $channel,
            'checked' => $checked,
        ]);

        // Update Google Sheet
        $url = 'https://script.google.com/macros/s/AKfycbzhlu7KV3dx3PS-9XPFBI9FMgI0JZIAgsuZY48Lchr_60gkSmx1hNAukKwFGZXgPwid/exec';

        $response = Http::post($url, [
            'channel' => $channel,
            'checked' => $checked
        ]);

        if ($response->failed()) {
            Log::error('Failed to send to GSheet:', [$response->body()]);
            return response()->json(['success' => false, 'message' => 'Failed to update GSheet'], 500);
        }

        Log::info("Google Sheet updated successfully");

        // Update Laravel DB
        ZeroVisibilityMaster::updateOrCreate(
            ['channel_name' => $channel],
            ['is_ra_checked' => $checked]
        );

        Log::info("Database updated for channel: $channel");

        return response()->json(['success' => true, 'message' => 'Updated GSheet & DB']);
    }


    private function getNRCount($channel)
    {
        $channel = strtolower(trim($channel));

        try {
            switch ($channel) {
                case 'amazon':
                    return app(AmazonZeroController::class)->getNrReqCount()['NR'] ?? 0;

                case 'ebay':
                    return app(EbayZeroController::class)->getNrReqCount()['NR'] ?? 0;

                case 'temu':
                    return app(TemuZeroController::class)->getNrReqCount()['NR'] ?? 0;

                case 'doba':
                    return app(DobaZeroController::class)->getNrReqCount()['NR'] ?? 0;

                case 'macys':
                    return app(MacyZeroController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'walmart':
                //     return app(ListingWalmartController::class)->getNrReqCount()['NR'] ?? 0;
                
                case 'wayfair':
                    return app(WayfairZeroController::class)->getNrReqCount()['NR'] ?? 0;
                
                // case 'ebay 3':
                //     return app(ListingEbayThreeController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'shopify b2c':
                //     return app(ListingShopifyB2CController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'reverb':
                //     return app(ListingReverbController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'aliexpress':
                //     return app(ListingAliexpressController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'shein':
                //     return app(ListingSheinController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'tiktok shop':
                //     return app(ListingTiktokShopController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'shopify wholesale/ds':
                //     return app(ListingShopifyWholesaleController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'faire':
                //     return app(ListingFaireController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'ebay 2':
                //     return app(ListingEbayTwoController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'mercari w ship':
                //     return app(ListingMercariWShipController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'newegg b2c':
                //     return app(ListingNeweggB2CController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'fb marketplace':
                //     return app(ListingFBMarketplaceController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'syncee':
                //     return app(ListingSynceeController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'auto ds':
                //     return app(ListingAutoDSController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'mercari w/o ship':
                //     return app(ListingMercariWoShipController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'business 5core':
                //     return app(ListingBusiness5CoreController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'zendrop':
                //     return app(ListingZendropController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'poshmark':
                //     return app(ListingPoshmarkController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'appscenic':
                //     return app(ListingAppscenicController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'tiendamia':
                //     return app(ListingTiendamiaController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'spocket':
                //     return app(ListingSpocketController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'offerup':
                //     return app(ListingOfferupController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'newegg b2b':
                //     return app(ListingNeweggB2BController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'fb shop':
                //     return app(ListingFBShopController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'instagram shop':
                //     return app(ListingInstagramShopController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'Yamibuy':
                //     return app(ListingYamibuyController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'dhgate':
                //     return app(ListingDHGateController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'bestbuy usa':
                //     return app(ListingBestbuyUSAController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'sw gear exchange':
                //     return app(ListingSWGearExchangeController::class)->getNrReqCount()['NR'] ?? 0;

                // case 'dhgate':
                //     return app(ListingDHGateController::class)->getNrReqCount()['NR'] ?? 0;
  

                default:
                    return 0;
            }
        } catch (\Throwable $e) {
            return 0;
        }
    }






}
