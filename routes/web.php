<?php

use App\Http\Controllers\AdsMaster\AdsMasterController;
use App\Http\Controllers\Channels\ChannelPromotionMasterController;
use App\Http\Controllers\MarketingMaster\CvrLQSMasterController;
use App\Http\Controllers\MarketingMaster\ListingMasterController;
use App\Http\Controllers\MarketPlace\AmazonFbaInvController;
use App\Http\Controllers\MarketPlace\AmazonLowVisibilityController;
use App\Http\Controllers\MarketPlace\AmazonZeroController;
use App\Http\Controllers\MarketPlace\EbayController;
use App\Http\Controllers\MarketPlace\EbayLowVisibilityController;
use App\Http\Controllers\MarketPlace\EbayZeroController;
use App\Http\Controllers\MarketPlace\ListingAuditAmazonController;
use App\Http\Controllers\MarketPlace\ListingAuditEbayController;
use App\Http\Controllers\MarketPlace\ListingAuditMacyController;
use App\Http\Controllers\MarketPlace\ListingAuditNeweggb2cController;
use App\Http\Controllers\MarketPlace\ListingAuditReverbController;
use App\Http\Controllers\MarketPlace\ListingAuditShopifyb2cController;
use App\Http\Controllers\MarketPlace\ListingAuditTemuController;
use App\Http\Controllers\MarketPlace\ListingAuditWayfairController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingAliexpressController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingAppscenicController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingAutoDSController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingBestbuyUSAController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingBusiness5CoreController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingDHGateController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingDobaController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingEbayController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingEbayThreeController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingEbayTwoController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingEbayVariationController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingFaireController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingFBMarketplaceController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingFBShopController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingInstagramShopController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingMacysController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingMercariWoShipController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingMercariWShipController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingNeweggB2BController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingNeweggB2CController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingOfferupController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingPlsController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingPoshmarkController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingReverbController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingSheinController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingShopifyB2CController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingShopifyWholesaleController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingSpocketController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingSWGearExchangeController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingSynceeController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingTemuController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingTiendamiaController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingTiktokShopController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingWalmartController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingWayfairController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingYamibuyController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingZendropController;
use App\Http\Controllers\MarketPlace\MacyController;
use App\Http\Controllers\MarketPlace\MacyLowVisibilityController;
use App\Http\Controllers\MarketPlace\MacyZeroController;
use App\Http\Controllers\MarketPlace\Neweggb2cController;
use App\Http\Controllers\MarketPlace\Neweggb2cLowVisibilityController;
use App\Http\Controllers\MarketPlace\Neweggb2cZeroController;
use App\Http\Controllers\MarketPlace\OverallAmazonFbaController;
use App\Http\Controllers\MarketPlace\OverallAmazonPriceController;
use App\Http\Controllers\MarketPlace\ReverbLowVisibilityController;
use App\Http\Controllers\MarketPlace\Shopifyb2cController;
use App\Http\Controllers\Channels\ChannelMasterController;
use App\Http\Controllers\Channels\ChannelwiseController;
use App\Http\Controllers\Channels\ReturnController;
use App\Http\Controllers\Channels\ExpensesController;
use App\Http\Controllers\Channels\ReviewController;
use App\Http\Controllers\Channels\HealthController;
use App\Http\Controllers\MarketPlace\OverallAmazonController;
use App\Http\Controllers\MarketPlace\Shopifyb2cLowVisibilityController;
use App\Http\Controllers\MarketPlace\Shopifyb2cZeroController;
use App\Http\Controllers\MarketPlace\TemuController;
use App\Http\Controllers\MarketingMaster\VideoPostedController;
use App\Http\Controllers\MarketPlace\TemuLowVisibilityController;
use App\Http\Controllers\MarketPlace\TemuZeroController;
use App\Http\Controllers\MarketPlace\WayfairController;
use App\Http\Controllers\MarketingMaster\ListingLQSMasterController;
use App\Http\Controllers\MarketPlace\WayfairLowVisibilityController;
use App\Http\Controllers\MarketPlace\WayfairZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\AliexpressZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\AppscenicZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\AutoDSZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\BestbuyUSAZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\Business5CoreZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\DHGateZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\DobaZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\Ebay2ZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\Ebay3ZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\EbayVariationZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\FaireZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\FBMarketplaceZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\FBShopZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\InstagramShopZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\MercariWoShipZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\MercariWShipZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\NeweggB2BZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\OfferupZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\PLSZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\PoshmarkZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\SheinZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\ShopifyWholesaleZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\SpocketZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\SWGearExchangeZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\SynceeZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\TiendamiaZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\TiktokShopZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\WalmartZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\YamibuyZeroController;
use App\Http\Controllers\MarketPlace\ZeroViewMarketPlace\ZendropZeroController;
use App\Http\Controllers\ProductMaster\PrAnalysisController;
use App\Http\Controllers\ProductMaster\ProductMasterController;
use App\Http\Controllers\Catalouge\CatalougeManagerController;
use App\Http\Controllers\Channels\AccountHealthMasterController;
use App\Http\Controllers\Channels\AccountHealthMasterDashboardController;
use App\Http\Controllers\ProductMaster\ReturnAnalysisController;
use App\Http\Controllers\ProductMaster\ReviewAnalysisController;
use App\Http\Controllers\ProductMaster\PricingAnalysisController;
use App\Http\Controllers\ProductMaster\ShortFallAnalysisController;
use App\Http\Controllers\ProductMaster\StockAnalysisController;
use App\Http\Controllers\ProductMaster\CostpriceAnalysisController;
use App\Http\Controllers\ProductMaster\MovementAnalysisController;
use App\Http\Controllers\ProductMaster\ForecastAnalysisController;
use App\Http\Controllers\InventoryManagement\VerificationAdjustmentController;
use App\Http\Controllers\ShopifyController;
use App\Http\Controllers\Listing\ListingManagerController;
use App\Http\Controllers\PurchaseMaster\SupplierController;
use App\Http\Controllers\ProductMaster\ToBeDCController;
use App\Http\Controllers\ProductMaster\ToOrderAnalysisController;
use App\Http\Controllers\PurchaseMaster\CategoryController;
use App\Http\Controllers\SkuMatchController;
use Illuminate\Foundation\Console\RouteCacheCommand;
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\RoutingController;
use App\Http\Controllers\MarketPlace\ReverbController;
use App\Http\Controllers\MarketPlace\ReverbZeroController;
use App\Http\Controllers\PurchaseMaster\ChinaLoadController;
use App\Http\Controllers\PurchaseMaster\MFRGInProgressController;
use App\Http\Controllers\PurchaseMaster\OnRoadTransitController;
use App\Http\Controllers\PurchaseMaster\OnSeaTransitController;
use App\Http\Controllers\Warehouse\WarehouseController;
use Illuminate\Support\Facades\Auth;
use App\Http\Controllers\PurchaseMaster\PurchaseOrderController;
use App\Http\Controllers\PurchaseMaster\ReadyToShipController;
use App\Http\Controllers\PricingMaster\PricingMasterController;
use App\Http\Controllers\MarketingMaster\ZeroVisibilityMasterController;
use App\Http\Controllers\MarketingMaster\ListingAuditMasterController;
use App\Http\Controllers\AdvertisementMaster\Kw_Advt\WalmartController;
use App\Http\Controllers\AdvertisementMaster\Kw_Advt\KwEbayController;
use App\Http\Controllers\AdvertisementMaster\Kw_Advt\KwAmazonController;
use App\Http\Controllers\AdvertisementMaster\Prod_Target_Advt\ProdTargetAmazonController;
use App\Http\Controllers\AdvertisementMaster\Headline_Advt\HeadlineAmazonController;
use App\Http\Controllers\AdvertisementMaster\Promoted_Advt\PromotedEbayController;
use App\Http\Controllers\AdvertisementMaster\Shopping_Advt\GoogleShoppingController;
use App\Http\Controllers\AdvertisementMaster\Demand_Gen_parent\GoogleNetworksController;
use App\Http\Controllers\AdvertisementMaster\MetaParent\ProductWiseMetaParentController;
use App\Http\Controllers\ArrivedContainerController;
use App\Http\Controllers\Campaigns\AmazonAdRunningController;
use App\Http\Controllers\Campaigns\AmazonCampaignReportsController;
use App\Http\Controllers\Campaigns\AmazonFbaAdsController;
use App\Http\Controllers\Campaigns\AmazonPinkDilAdController;
use App\Http\Controllers\Campaigns\AmazonSbBudgetController;
use App\Http\Controllers\Campaigns\AmazonSpBudgetController;
use App\Http\Controllers\Campaigns\AmzCorrectlyUtilizedController;
use App\Http\Controllers\Campaigns\AmzUnderUtilizedBgtController;
use App\Http\Controllers\Campaigns\CampaignImportController;
use App\Http\Controllers\Campaigns\EbayKwAdsController;
use App\Http\Controllers\Campaigns\EbayOverUtilizedBgtController;
use App\Http\Controllers\Campaigns\EbayPinkDilAdController;
use App\Http\Controllers\Campaigns\EbayPMPAdsController;
use App\Http\Controllers\Campaigns\GoogleShoppingAdsController;
use App\Http\Controllers\Campaigns\WalmartUtilisationController;
use App\Http\Controllers\Channels\ApprovalsChannelMasterController;
use App\Http\Controllers\EbayDataUpdateController;
use App\Http\Controllers\PurchaseMaster\PurchaseController;
use App\Http\Controllers\PurchaseMaster\TransitContainerDetailsController;
use App\Http\Controllers\InventoryManagement\IncomingController;
use App\Http\Controllers\InventoryManagement\OutgoingController;
use App\Http\Controllers\InventoryManagement\StockAdjustmentController;
use App\Http\Controllers\InventoryManagement\StockTransferController;
use App\Http\Controllers\Channels\ChannelMovementAnalysisController;
use App\Http\Controllers\Channels\NewMarketplaceController;
use App\Http\Controllers\Channels\OpportunityController;
use App\Http\Controllers\Channels\ReviewMaster\AmazonReviewController;
use App\Http\Controllers\Channels\ReviewMaster\ReviewDashboardController;
use App\Http\Controllers\Channels\SetupAccountChannelController;
use App\Http\Controllers\Channels\ShippingMasterController;
use App\Http\Controllers\Channels\TrafficMasterController;
use App\Http\Controllers\InventoryManagement\AutoStockBalanceController;
use App\Http\Controllers\InventoryManagement\StockBalanceController;
use App\Http\Controllers\InventoryWarehouseController;
use App\Http\Controllers\MarketPlace\DobaController;
use App\Http\Controllers\PurchaseMaster\ClaimReimbursementController;
use App\Http\Controllers\MarketingMaster\VideoAdsMasterController;
use App\Http\Controllers\MarketingMaster\DmMarketingController;
use App\Http\Controllers\MarketingMaster\EmailMarketingController;
use App\Http\Controllers\MarketingMaster\LetterMarketingController;
use App\Http\Controllers\MarketingMaster\PhoneAppMarketingController;
use App\Http\Controllers\MarketingMaster\SmsMarketingController;
use App\Http\Controllers\MarketingMaster\WhatsappMarketingController;
use App\Http\Controllers\MarketPlace\EbayTwoController;
use App\Http\Controllers\MarketPlace\EbayThreeController;
use App\Http\Controllers\MarketPlace\WalmartControllerMarket;
use App\Http\Controllers\MarketingMaster\CarouselSalesController;
use App\Http\Controllers\MarketingMaster\EbayCvrLqsController;
use App\Http\Controllers\MarketingMaster\ShoppableVideoController;
use App\Http\Controllers\MarketPlace\ACOSControl\AmazonACOSController;
use App\Http\Controllers\MarketPlace\ACOSControl\EbayACOSController;
use App\Http\Controllers\MarketPlace\AliexpressController;
use App\Http\Controllers\MarketPlace\Ebay2LowVisibilityController;
use App\Http\Controllers\MarketPlace\Ebay3LowVisibilityController;
use App\Http\Controllers\MarketPlace\ListingMarketPlace\ListingAmazonController;
use App\Http\Controllers\MarketPlace\SheinController;
use App\Http\Controllers\MarketPlace\TiktokShopController;
use App\Http\Controllers\PurchaseMaster\LedgerMasterController;
use App\Http\Controllers\PricingIncDsc\MasterIncDscController;
use App\Http\Controllers\PricingMaster\PricingMasterViewsController;
use App\Http\Controllers\PurchaseMaster\ContainerPlanningController;
use App\Http\Controllers\PurchaseMaster\QualityEnhanceController;
use App\Http\Controllers\PurchaseMaster\RFQController;
use App\Http\Controllers\PurchaseMaster\SourcingController;
use App\Http\Controllers\MarketingMaster\FacebookAddsManagerController;
use App\Http\Controllers\MarketingMaster\MovementPricingMaster;
use App\Http\Controllers\MarketingMaster\OverallCvrLqsController;
use App\Http\Controllers\PurchaseMaster\SupplierRFQController;
use App\Http\Controllers\StockMappingController;
/*  
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

Route::prefix('auth')->group(function () {
    require __DIR__ . '/auth.php';
});
Route::get('/auth/logout-page', function () {
    // Prevent access if user is still logged in
    if (Auth::check()) {
        return redirect('/home');
    }

    return view('auth.logout');
})->name('logout.page');

Route::get('/sku-match', [SkuMatchController::class, 'index']);
Route::post('/sku-match', [SkuMatchController::class, 'match'])->name('sku.match.process');
Route::post('/sku-match/update', [SkuMatchController::class, 'update'])->name('sku-match.update');




Route::group(['prefix' => '/', 'middleware' => 'auth'], function () {


    Route::get('/amazon-summary-data', [OverallAmazonController::class, 'getAmazonDataSummary']);
    Route::get('/ebay-data-view', [EbayController::class, 'getViewEbayData']);
    Route::get('/ebay2-data-view', [EbayTwoController::class, 'getViewEbayData']);
    Route::get('/shopifyB2C-data-view', [Shopifyb2cController::class, 'getViewShopifyB2CData']);
    Route::get('/listing-master-data', [ListingManagerController::class, 'getViewListingMasterData']);
    Route::get('/macy-data-view', [MacyController::class, 'getViewMacyData']);
    Route::get('/macy-pricing-cvr', [MacyController::class, 'macyPricingCvr']);
    Route::get('/macy-pricing-increase-decrease', [MacyController::class, 'macyPricingIncreaseandDecrease']);

    Route::get('/product-master-data-view', [ProductMasterController::class, 'getViewProductData']);
    Route::get('/neweggB2C-data-view', [Neweggb2cController::class, 'getViewNeweggB2CData']);
    Route::get('/review-analysis-data-view', [ReviewAnalysisController::class, 'getViewReviewAnalysisData']);
    Route::get('/pricing-analysis-data-view', [PricingAnalysisController::class, 'getViewPricingAnalysisData']);
    Route::get('/pRoi-analysis-data-view', [PrAnalysisController::class, 'getViewPRoiAnalysisData']);
    Route::get('/return-analysis-data-view', [ReturnAnalysisController::class, 'getViewReturnAnalysisData']);
    Route::get('/stock-analysis-data-view', [StockAnalysisController::class, 'getViewStockAnalysisData']);
    Route::get('/shortfall-analysis-data-view', [ShortFallAnalysisController::class, 'getViewShortFallAnalysisData']);
    Route::get('/costprice-analysis-data-view', [CostpriceAnalysisController::class, 'getViewCostpriceAnalysisData']);
    Route::get('/movement-analysis-data-view', [MovementAnalysisController::class, 'getViewMovementAnalysisData']);
    Route::get('/forecast-analysis-data-view', [ForecastAnalysisController::class, 'getViewForecastAnalysisData']);
    Route::post('/updateForcastSheet', [ForecastAnalysisController::class, 'updateForcastSheet']);
    Route::get('/wayfair-data-view', [WayfairController::class, 'getViewWayfairData']);

    //channel master
    Route::post('/update-executive', [ChannelMasterController::class, 'updateExecutive']);
    Route::post('/update-checkbox', [ChannelMasterController::class, 'sendToGoogleSheet']);
    Route::get('/channels-master-data', [ChannelMasterController::class, 'getViewChannelData']);

    //Zero Visibility Master
    Route::get('/zero-visibility-master', [ZeroVisibilityMasterController::class, 'index'])->name('zero.visibility');
    // Route::get('/zero-visibility-data', [ZeroVisibilityMasterController::class, 'getViewChannelData']);
    Route::post('/store-zero-visibility', [ZeroVisibilityMasterController::class, 'store']);
    Route::get('/show-zero-visibility-data', [ZeroVisibilityMasterController::class, 'getMergedChannelData']);
    Route::get('/export-zero-visibility-csv', [ZeroVisibilityMasterController::class, 'exportCsv'])->name('zero.export.csv');
    Route::post('/update-ra-checkbox', [ZeroVisibilityMasterController::class, 'updateRaCheckbox']);

    //Listing Audit Master
    Route::get('/listing-audit-master', [ListingAuditMasterController::class, 'index'])->name('listing.audit');
    Route::post('/store-list-audit-amazon-data', [ListingAuditMasterController::class, 'storeListingAuditAmazonData']);
    Route::post('/store-list-audit-ebay-data', [ListingAuditMasterController::class, 'storeListingAuditEbayData']);
    Route::get('/show-list-audit-master-data', [ListingAuditMasterController::class, 'getListingAuditSummaryWithChannelInfo']);
    Route::get('/listing-audit-master-data', [ListingAuditMasterController::class, 'getAuditMasterTableData']);
    Route::get('/export-listing-audit-csv', [ListingAuditMasterController::class, 'exportListingAuditCSV']);
    Route::post('/update-ra-checkbox', [ListingAuditMasterController::class, 'updateRaCheckbox']);

    //Email Marketing

    Route::get('/email-marketing-master', [EmailMarketingController::class, 'index'])->name('email.marketing');
    Route::get('/whatsapp-marketing-master', [WhatsappMarketingController::class, 'index'])->name('whatsapp.marketing');
    Route::get('/sms-marketing-master', [SmsMarketingController::class, 'index'])->name('sms.marketing');
    Route::get('/dm-marketing-master', [DmMarketingController::class, 'index'])->name('dm.marketing');
    Route::get('/phone-marketing-master', [PhoneAppMarketingController::class, 'index'])->name('phone.marketing');
    Route::get('/letter-marketing-master', [LetterMarketingController::class, 'index'])->name('letter.marketing');
    Route::get('/carousel-sales-master', [CarouselSalesController::class, 'index'])->name('carousel.sales');


    //Account Health Master
    // Route::get('/channel/account-health-test', [AccountHealthMasterController::class, 'test'])->name('account.health.master');
    Route::controller(AccountHealthMasterController::class)->group(function () {
        Route::get('/account-health-master/odr-rate', 'odrRateIndex')->name('odr.rate');
        Route::post('/odr-rate-save', 'saveOdrRate')->name('odr.rate.save');
        Route::get('/fetchOdrRates', 'fetchOdrRates');
        Route::post('/odr-rate/update', 'updateOdrRate');
        Route::post('/odr-health-link/update', 'updateOdrHealthLink')->name('odr-health.link.update');

        Route::get('/account-health-master/fullfillment-rate', 'fullfillmentRateIndex')->name('fullfillment.rate');
        Route::post('/fullfillment-rate-save', 'saveFullfillmentRate')->name('fullfillment.rate.save');
        Route::get('/fetchFullfillmentRates', 'fetchFullfillmentRates');
        Route::post('/fullfillment-rate/update', 'updateFullfillmentRate');
        Route::post('/fullfillment-health-link/update', 'updateFullfillmentHealthLink')->name('fullfillment-health.link.update');

        Route::get('/account-health-master/valid-tracking-rate', 'validTrackingRateIndex')->name('valid.tracking.rate');
        Route::post('/validtracking-rate-save', 'saveValidTrackingRate')->name('validtracking.rate.save');
        Route::get('/fetchValidTrackingRates', 'fetchValidTrackingRates');
        Route::post('/validtracking-rate/update', 'updateValidTrackingRate');
        Route::post('/validtracking-health-link/update', 'updateValidTrackingHealthLink')->name('validtracking-health.link.update');

        Route::get('/account-health-master/late-shipment', 'lateShipmentRateIndex')->name('late.shipment.rate');
        Route::post('/lateshipment-rate-save', 'saveLateShipmentRate')->name('lateshipment.rate.save');
        Route::get('/fetchLateShipmentRates', 'fetchLateShipmentRates');
        Route::post('/lateshipment-rate/update', 'updateLateShipmentRate');
        Route::post('/lateshipment-health-link/update', 'updateLateShipmentHealthLink')->name('lateshipment-health.link.update');

        Route::get('/account-health-master/on-time-delivery', 'onTimeDeliveryIndex')->name('on.time.delivery.rate');
        Route::post('/onTimeDelivery-rate-save', 'saveOnTimeDeliveryRate')->name('onTimeDelivery.rate.save');
        Route::get('/fetchOnTimeDeliveryRates', 'fetchOnTimeDeliveryRates');
        Route::post('/onTimeDelivery-rate/update', 'updateOnTimeDeliveryRate');
        Route::post('/onTimeDelivery-health-link/update', 'updateOnTimeDeliveryHealthLink')->name('onTimeDelivery-health.link.update');

        Route::get('/account-health-master/negative-seller', 'negativeSellerIndex')->name('negative.seller.rate');
        Route::post('/negativeSeller-rate-save', 'saveNegativeSellerRate')->name('negativeSeller.rate.save');
        Route::get('/fetchNegativeSellerRates', 'fetchNegativeSellerRates');
        Route::post('/negativeSeller-rate/update', 'updateNegativeSellerRate');
        Route::post('/negativeSeller-health-link/update', 'updateNegativeSellerHealthLink')->name('negativeSeller-health.link.update');

        Route::get('/account-health-master/a-z-claims', 'aTozClaimsIndex')->name('a_z.claims.rate');
        Route::post('/AtoZClaims-rate-save', 'saveAtoZClaimsRate')->name('AtoZClaims.rate.save');
        Route::get('/fetchAtoZClaimsRates', 'fetchAtoZClaimsRates');
        Route::post('/AtoZClaims-rate/update', 'updateAtoZClaimsRate');
        Route::post('/AtoZClaims-health-link/update', 'updateAtoZClaimsHealthLink')->name('AtoZClaims-health.link.update');

        Route::get('/account-health-master/voilation-compliance', 'voilationIndex')->name('voilation.rate');
        Route::post('/voilance-rate-save', 'saveVoilanceRate')->name('voilance.rate.save');
        Route::get('/fetchVoilanceRates', 'fetchVoilanceRates');
        Route::post('/voilance-rate/update', 'updateVoilanceRate');
        Route::post('/voilance-health-link/update', 'updateVoilanceHealthLink')->name('voilance-health.link.update');

        Route::get('/account-health-master/refund-return', 'refundIndex')->name('refund.rate');
        Route::post('/refund-rate-save', 'saveRefundRate')->name('refund.rate.save');
        Route::get('/fetchRefundRates', 'fetchRefundRates');
        Route::post('/refund-rate/update', 'updateRefundRate');
        Route::post('/refund-health-link/update', 'updateRefundHealthLink')->name('refund-health.link.update');
    });

    // Account Health Master Channel Dashboard
    Route::get('/channel/dashboard', [AccountHealthMasterDashboardController::class, 'dashboard'])->name('account.health.master.channel.dashboard');
    Route::get('/account-health-master/dashboard-data', [AccountHealthMasterDashboardController::class, 'getMasterChannelDataHealthDashboard'])->name('account.health.master.dashboard.data');
    Route::get('/account-health-master/export', [AccountHealthMasterDashboardController::class, 'export'])->name('account-health-master.export');
    Route::post('/account-health-master/import', [AccountHealthMasterDashboardController::class, 'import'])->name('account-health-master.import');
    Route::get('/account-health-master/sample/{type?}', [AccountHealthMasterDashboardController::class, 'downloadSample'])->name('account-health-master.sample');

    Route::controller(OpportunityController::class)->group(function () {
        Route::get('/channerl-master-active/opportunity', 'index')->name('opportunity.index');
        Route::get('/opportunities/data', 'getOpportunitiesData');
        Route::post('/opportunities/save', 'saveOpportunity');
        Route::post('/import-opportunities', 'importOpportunities')->name('import.opportunities');
        Route::get('/opportunities/export', 'exportOpportunities')->name('opportunities.export');
        Route::post('/opportunities/delete', 'deleteOpportunities')->name('opportunities.export');
    });
    Route::controller(ApprovalsChannelMasterController::class)->group(function () {
        Route::get('/channerl-master-active/application-and-approvals', 'index')->name('application.approvals.index');
        Route::get('/approvals/data', 'fetchApprovalsData');
        Route::post('/approvals-channel-master/save', 'saveApprovalsData');
        Route::post('/approvals/delete', 'deleteApprovals')->name('approvals.export');
    });
    Route::controller(SetupAccountChannelController::class)->group(function () {
        Route::get('/channerl-master-active/setup-account-and-shop', 'index')->name('setup.account.index');
        Route::get('/setup-account/fetch-data', 'fetchSetupAccountData');
        Route::post('/setup-account-channel-master/save', 'saveSetupAccountData');
    });

    //Shipping Master
    Route::controller(ShippingMasterController::class)->group(function () {
        Route::get('/shipping-master/list', 'index')->name('shipping.master.list');
        Route::get('/fetch-shipping-rate/data', 'fetchShippingRate');
        Route::post('/update-shipping-rate', 'storeOrUpdateShippingRate');
    });

    Route::controller(TrafficMasterController::class)->group(function () {
        Route::get('/traffic-master/list', 'index')->name('traffic.master.list');
        Route::get('/fetch-traffic-rate/data', 'fetchTraficReport');
    });

    Route::get('/channel/account-health-master', [AccountHealthMasterController::class, 'index'])->name('account.health.master');
    Route::post('/channel/account-health-master/store', [AccountHealthMasterController::class, 'store'])->name('account.health.store');
    Route::post('/account-health/link/update', [AccountHealthMasterController::class, 'updateLink'])->name('account.health.link.update');
    Route::post('/account-health/update', [AccountHealthMasterController::class, 'update']);

    //verification & Adjustment
    Route::get('/verification-adjustment-data-view', [VerificationAdjustmentController::class, 'getViewVerificationAdjustmentData']);
    Route::get('/verification-adjustment-view', [VerificationAdjustmentController::class, 'index'])->name('verify-adjust');
    Route::post('/update-verified-stock', [VerificationAdjustmentController::class, 'updateVerifiedStock']);
    Route::get('/get-verified-stock', [VerificationAdjustmentController::class, 'getVerifiedStock']);
    Route::post('/update-to-adjust', [ShopifyController::class, 'updateToAdjust']);
    Route::post('/update-approved-by-ih', [VerificationAdjustmentController::class, 'updateApprovedByIH']);
    Route::post('/update-ra-status', [VerificationAdjustmentController::class, 'updateRAStatus']);
    Route::get('/verified-stock-activity-log', [VerificationAdjustmentController::class, 'getVerifiedStockActivityLog']);
    Route::get('/view-inventory-data', [VerificationAdjustmentController::class, 'viewInventory'])->name('view-inventory');
    Route::get('/inventory-history', [VerificationAdjustmentController::class, 'getSkuWiseHistory']);
    Route::post('/row-hide-toggle', [VerificationAdjustmentController::class, 'toggleHide']);
    Route::get('/get-hidden-rows', [VerificationAdjustmentController::class, 'getHiddenRows']);
    Route::post('/unhide-multiple-rows', [VerificationAdjustmentController::class, 'unhideMultipleRows']);

    //incoming
    Route::get('/incoming-view', [IncomingController::class, 'index'])->name('incoming.view');
    Route::post('/incoming-data-store', [IncomingController::class, 'store'])->name('incoming.store');
    Route::get('/incoming-data-list', [IncomingController::class, 'list']);

    //incoming orders
    Route::get('/incoming-orders-view', [IncomingController::class, 'incomingOrderIndex'])->name('incoming.orders.view');
    Route::post('/incoming-orders-store', [IncomingController::class, 'incomingOrderStore'])->name('incoming.orders.store');
    Route::get('/incoming-orders-list', [IncomingController::class, 'incomingOrderList']);

    //outgoing
    Route::get('/outgoing-view', [OutgoingController::class, 'index'])->name('outgoing.view');
    Route::post('/outgoing-data-store', [OutgoingController::class, 'store'])->name('outgoing.store');
    Route::get('/outgoing-data-list', [OutgoingController::class, 'list']);

    //Auto Stock Balance
    Route::get('/auto-stock-balance-view', [AutoStockBalanceController::class, 'index'])->name('autostock.balance.view');
    Route::post('/auto-stock-balance-store', [AutoStockBalanceController::class, 'store'])->name('autostock.balance.store');
    Route::get('/auto-stock-balance-data-list', [AutoStockBalanceController::class, 'list']);

    //Linked products
    Route::get('/linked-products-view', [ProductMasterController::class, 'linkedProductsView'])->name('linked.products.view');
    Route::post('/linked-products-store', [ProductMasterController::class, 'linkedProductStore'])->name('linked.products.store');
    Route::get('/linked-products-data-list', [ProductMasterController::class, 'linkedProductsList']);

    //show updated qty
    Route::get('/show-updated-qty', [ProductMasterController::class, 'showUpdatedQty'])->name('show.updated.qty');
    Route::get('/show-updated-qty-list', [ProductMasterController::class, 'showUpdatedQtyList']);

    //Stock Adjustment
    Route::get('/stock-adjustment-view', [StockAdjustmentController::class, 'index'])->name('stock.adjustment.view');
    Route::post('/stock-adjustment-store', [StockAdjustmentController::class, 'store'])->name('stock.adjustment.store');
    Route::get('/stock-adjustment-data-list', [StockAdjustmentController::class, 'list']);

    //Stock Transfer
    Route::get('/stock-transfer-view', [StockTransferController::class, 'index'])->name('stock.transfer.view');
    Route::post('/stock-transfer-store', [StockTransferController::class, 'store'])->name('stock.transfer.store');
    Route::get('/stock-transfer-data-list', [StockTransferController::class, 'list']);

    //Stock Balance
    Route::get('/stock-balance-view', [StockBalanceController::class, 'index'])->name('stock.balance.view');
    Route::post('/stock-balance-store', [StockBalanceController::class, 'store'])->name('stock.balance.store');
    Route::get('/stock-balance-data-list', [StockBalanceController::class, 'list']);

    //channel Movement Analysis
    Route::get('/channel-movement-analysis', [ChannelMovementAnalysisController::class, 'index'])->name('channel.movement.analysis');
    Route::get('/channel-analysis/{channel}', [ChannelMovementAnalysisController::class, 'show'])->name('channel.show');
    Route::post('/channel-analysis/update', [ChannelMovementAnalysisController::class, 'updateField'])->name('channel.updateField');
    Route::get('/channels/get-monthly-data/{channel}', [ChannelMovementAnalysisController::class, 'getMonthlyData'])->name('channels.getMonthlyData');

    Route::get('/master-pricing-inc-dsc', [MasterIncDscController::class, 'index'])->name('master.pricing.inc.dsc');
    Route::get('/master-pricing/{channel}', [MasterIncDscController::class, 'show'])->name('channel.show');
    Route::post('/master-pricing/update', [MasterIncDscController::class, 'updateField'])->name('channel.updateField');
    Route::get('/master-pricing/get-monthly-data/{channel}', [MasterIncDscController::class, 'getMonthlyData'])->name('channels.getMonthlyData');


    //New Marketplaces
    Route::get('/new-marketplaces-dashboard', [NewMarketplaceController::class, 'index'])->name('new.marketplaces.dashboard');
    Route::get('/channels/fetch', [NewMarketplaceController::class, 'getChannelsFromGoogleSheet'])->name('channels.fetch');
    Route::post('/new-marketplaces-store', [NewMarketplaceController::class, 'store'])->name('new.marketplaces.store');
    Route::get('/new-marketplaces-by-status', [NewMarketplaceController::class, 'getMarketplacesByStatus'])->name('new.marketplaces.byStatus');
    Route::post('/new-marketplaces/import', [NewMarketplaceController::class, 'import'])->name('new-marketplaces.import');
    Route::get('/new-marketplaces/export', [NewMarketplaceController::class, 'export'])->name('new-marketplaces.export');
    Route::get('/edit-new-marketplaces/{id}', [NewMarketplaceController::class, 'edit']);
    Route::get('/new-marketplaces/{id}', [NewMarketplaceController::class, 'show']);
    Route::post('/new-marketplaces/{id}', [NewMarketplaceController::class, 'update']);
    Route::post('/marketplaces-update-status', [NewMarketplaceController::class, 'updateStatus'])->name('marketplaces.updateStatus');


    // Route::post('/new-marketplaces/update/{id}', [NewMarketplaceController::class, 'update']);

    //Warehouse
    Route::get('/list_all_warehouses', [WarehouseController::class, 'index'])->name('list_all_warehouses');
    Route::post('/warehouses/store', [WarehouseController::class, 'store'])->name('warehouses.store');
    Route::get('/warehouses/list', [WarehouseController::class, 'list']);
    Route::get('/warehouses/{id}/edit', [WarehouseController::class, 'edit']);
    Route::post('/warehouses/update/{id}', [WarehouseController::class, 'update'])->name('warehouses.update');
    Route::delete('/warehouses/{id}', [WarehouseController::class, 'destroy']);

    Route::get('/main-godown', [WarehouseController::class, 'mainGodown'])->name('main.godown');
    Route::get('/return-godown', [WarehouseController::class, 'returnGodown'])->name('returns.godown');
    Route::get('/openbox-godown', [WarehouseController::class, 'openBoxGodown'])->name('openbox.godown');
    Route::get('/showroom-godown', [WarehouseController::class, 'showroomGodown'])->name('showroom.godown');
    Route::get('/useditem-godown', [WarehouseController::class, 'usedItemGodown'])->name('useditem.godown');
    Route::get('/trash-godown', [WarehouseController::class, 'trashGodown'])->name('trash.godown');



    //Purchase Order
    Route::controller(PurchaseOrderController::class)->group(function () {
        Route::get('/list-all-purchase-orders', 'index')->name('list-all-purchase-orders');
        Route::post('/store-purchase-orders', 'store')->name('purchase-orders.store');
        Route::get('/purchase-orders/list', 'getPurchaseOrdersData')->name('purchase-orders.data');
        Route::get('/purchase-orders/convert', 'convert')->name('purchase-orders.convert');
        Route::get('/purchase-order/{id}/generate-pdf', 'generatePdf')->name('generate-pdf');
        Route::post('/purchase-orders/delete', 'deletePurchaseOrders');
        Route::get('/purchase-orders/{id}', 'showPurchaseOrders');
        Route::post('/purchase-orders/{id}', 'updatePurchaseOrder');
    });

    //Purchase
    Route::controller(PurchaseController::class)->group(function () {
        Route::get('/purchase/list', 'index')->name('purchase.index');
        Route::get('/purchase-orders/items-by-supplier/{supplier_id}', 'getItemsBySupplier');
        Route::get('/product-master/get-parent/{sku}', 'getParentBySku');
        Route::post('/purchase/save', 'store')->name('purchase.store');
        Route::get('/purchase-data/list', 'getPurchaseSummary');
        Route::post('/purchase/delete', 'deletePurchase');
    });

    //RFQ Form
    Route::controller(RFQController::class)->group(function () {
        Route::get('/rfq-form/list', 'index')->name('rfq-form.index');
        Route::get('rfq-form/data', 'getRfqFormsData');
        Route::post('/rfq-form/store', 'storeRFQForm')->name('rfq-form.store');
        Route::get('/rfq-form/edit/{id}', 'edit')->name('rfq-form.edit');
        Route::post('/rfq-form/update/{id}', 'update')->name('rfq-form.update');
        Route::delete('/rfq-form/delete/{id}', 'destroy')->name('rfq-form.destroy');

        //form reports
        Route::get('/rfq-form/reports/{id}', 'rfqReports')->name('rfq-form.reports');
        Route::get('/rfq-form/reports-data/{id}', 'getRfqReportsData')->name('rfq-form.reports.data');
    });

    //Sourcingƒvies
    Route::controller(SourcingController::class)->group(function () {
        Route::get('/sourcing/list', 'index')->name('sourcing.index');
        Route::get('/sourcing-data/list', 'getSourcingData')->name('sourcing.list');
        Route::post('/sourcing/save', 'storeSourcing')->name('sourcing.save');
        Route::post('/sourcing/update/{id}', 'updateSourcing')->name('sourcing.update');
        Route::post('/sourcing/delete', 'deleteSourcing')->name('sourcing.delete');
        Route::get('/get-parent-by-sku/{sku}', 'getParentBySku')->name('getParentBySku');
    });

    //Review Master
    Route::controller(AmazonReviewController::class)->group(function () {
        Route::get('/review-master/amazon-product-reviews/', 'index')->name('review.masters.amazon');
        Route::get('/amazon-product-review-data', 'fetchAmazonProductReview');
        Route::post('/amazon-product-review-import', 'importProductReview')->name('amazon.product.review.import');
        Route::post('/amazon-product-reviews/save', 'createUpdateProductReview');
    });

    //Review Dashboard
    Route::controller(ReviewDashboardController::class)->group(function () {
        Route::get('/review-master/daboard', 'index')->name('review.master.dashboard');
        Route::get('/review-dashboard-data', 'getReviewDataChannelBased');
    });

    //LedgerMaster
    Route::controller(LedgerMasterController::class)->group(function () {
        Route::get('/ledger-master/advance-and-payments/', 'advanceAndPayments')->name('ledger.advance.payments');
        Route::get('/ledger-master/supplier-ledger/', 'supplierLedger')->name('supplier.ledger');
        Route::post('/ledger-master/supplier-ledger-save', 'supplierStore')->name('supplier.ledger.save');
        Route::post('/supplier-ledger/update', 'updateSupplierLedger')->name('supplier.ledger.update');
        Route::get('/supplier-ledger/get-balance', 'getSupplierBalance')->name('supplier.ledger.get-balance');
        Route::get('/supplier-ledger/list', 'fetchSupplierLedgerData');
        Route::post('/advance-payments/save', 'saveAdvancePayments')->name('advance.payments.save');
        Route::get('/advance-and-payments/data', 'getAdvancePaymentsData');
        Route::post('/advance-payments/delete', 'deleteAdvancePayments');
        Route::post('/supplier-ledger/delete', 'deleteSupplierLedger');
    });



    // Doba Routes
    Route::get('/zero-doba', [DobaZeroController::class, 'dobaZeroview'])->name('zero.doba');
    Route::get('/zero_doba/view-data', [DobaZeroController::class, 'getViewDobaZeroData']);
    Route::post('/zero_doba/reason-action/update-data', [DobaZeroController::class, 'updateReasonAction']);

    Route::get('/listing-doba', [ListingDobaController::class, 'listingDoba'])->name('listing.doba');
    Route::get('/listing_doba/view-data', [ListingDobaController::class, 'getViewListingDobaData']);
    Route::post('/listing_doba/save-status', [ListingDobaController::class, 'saveStatus']);
    Route::post('/listing_doba/import', [ListingDobaController::class, 'import'])->name('listing_doba.import');
    Route::get('/listing_doba/export', [ListingDobaController::class, 'export'])->name('listing_doba.export');

    Route::get('/doba-data-view', [DobaController::class, 'getViewDobaData']);
    Route::get('/doba', [DobaController::class, 'dobaView'])->name('doba');
    Route::post('/doba/save-nr', [DobaController::class, 'saveNrToDatabase']);
    Route::post('/doba/update-listed-live', [DobaController::class, 'updateListedLive']);
    Route::post('/doba/saveLowProfit', [DobaController::class, 'saveLowProfit']);
    Route::post('/update-doba-pricing', [DobaController::class, 'updatePrice']);
    Route::get('/doba-pricing-cvr', [DobaController::class, 'dobaPricingCVR']);
    Route::post('/doba/save-sprice', [DobaController::class, 'saveSpriceToDatabase'])->name('doba.save-sprice');
    Route::post('/update-all-doba-skus', [DobaController::class, 'updateAllDobaSkus']);
    Route::post('/doba-analytics/import', [DobaController::class, 'importDobaAnalytics'])->name('doba.analytics.import');
    Route::get('/doba-analytics/export', [DobaController::class, 'exportDobaAnalytics'])->name('doba.analytics.export');
    Route::get('/doba-analytics/sample', [DobaController::class, 'downloadSample'])->name('doba.analytics.sample');


    //update sku inv and l30
    Route::post('/update-all-amazon-skus', [OverallAmazonController::class, 'updateAllAmazonSkus']);
    Route::post('/update-all-amazon-fba-skus', [OverallAmazonFbaController::class, 'updateAllAmazonfbaSkus']);
    Route::post('/update-all-ebay-skus', [EbayController::class, 'updateAllEbaySkus']);
    Route::post('/update-all-ebay-skus', [EbayTwoController::class, 'updateAllEbay2Skus']);

    Route::post('/update-all-shopifyB2C-skus', [Shopifyb2cController::class, 'updateAllShopifyB2CSkus']);
    Route::post('/update-all-macy-skus', [MacyController::class, 'updateAllMacySkus']);
    Route::post('/update-all-neweggb2c-skus', [Neweggb2cController::class, 'updateAllNeweggB2CSkus']);
    Route::post('/update-all-reverb-skus', [ReverbController::class, 'updateAllAReverbSkus']);
    Route::post('/update-all-wayfair-skus', [WayfairController::class, 'updateAllWayfairSkus']);
    Route::post('/update-all-reverb-skus', [ReverbController::class, 'updateAllReverbSkus']);
    Route::post('/update-all-temu-skus', [TemuController::class, 'updateAllTemuSkus']);
    Route::post('/update-amazon-price', action: [OverallAmazonController::class, 'updatePrice'])->name('amazon.priceChange');

    //ajax routes
    Route::get('/amazon/all-data', [OverallAmazonController::class, 'getAllData'])->name('amazon.allData');
    Route::get('/channel/all-data', [ChannelMasterController::class, 'getAllData'])->name('channel.allData');
    Route::get('/amazon/view-data', [OverallAmazonController::class, 'getViewAmazonData'])->name('amazon.viewData');
    Route::post('/update-fba-status', [OverallAmazonController::class, 'updateFbaStatus'])
        ->name('update.fba.status');
    Route::get('/listing_audit_amazon/view-data', [ListingAuditAmazonController::class, 'getViewListingAuditAmazonData']);
    Route::get('/listing_audit_ebay/view-data', [ListingAuditEbayController::class, 'getViewListingAuditEbayData']);
    Route::get('/listing_ebay/view-data', [ListingEbayController::class, 'getViewListingEbayData']);
    Route::get('/amazon/zero/view-data', [AmazonZeroController::class, 'getViewAmazonZeroData'])->name('amazon.zero.viewData');
    Route::get('/amazon/low-visibility/view-data', [AmazonLowVisibilityController::class, 'getViewAmazonLowVisibilityData']);
    Route::get('/amazon/low-visibility/view-data-fba', [AmazonLowVisibilityController::class, 'getViewAmazonLowVisibilityDataFba']);
    Route::get('/amazon/low-visibility/view-data-fbm', [AmazonLowVisibilityController::class, 'getViewAmazonLowVisibilityDataFbm']);
    Route::get('/amazon/low-visibility/view-data-both', [AmazonLowVisibilityController::class, 'getViewAmazonLowVisibilityDataBoth']);


    Route::get('/ebay/zero/view-data', [EbayZeroController::class, 'getVieweBayZeroData'])->name('ebay.zero.viewData');
    Route::get('/ebay/low-visibility/view-data', [EbayLowVisibilityController::class, 'getVieweBayLowVisibilityData']);
    Route::get('/ebay2/low-visibility/view-data', [Ebay2LowVisibilityController::class, 'getVieweBay2LowVisibilityData']);
    Route::get('/ebay3/low-visibility/view-data', [Ebay3LowVisibilityController::class, 'getVieweBay3LowVisibilityData']);
    Route::get('/reverb/view-data', [ReverbController::class, 'getViewReverbData']);
    Route::get('/shopifyB2C/view-data', [Shopifyb2cZeroController::class, 'getViewShopifyB2CZeroData']);
    Route::get('/shopifyB2C/low-visibility/view-data', [Shopifyb2cLowVisibilityController::class, 'getViewShopifyB2CLowVisibilityData']);
    Route::get('/Macy/view-data', [MacyZeroController::class, 'getViewMacyZeroData']);
    Route::get('/Macy/low-visibility/view-data', [MacyLowVisibilityController::class, 'getViewMacyLowVisibilityData']);
    Route::get('/Neweggb2c/view-data', [Neweggb2cZeroController::class, 'getViewNeweggB2CZeroData']);
    Route::get('/Neweggb2c/low-visiblity/view-data', [Neweggb2cLowVisibilityController::class, 'getViewNeweggB2CLowVisibilityData']);
    Route::get('/Wayfaire/view-data', [WayfairZeroController::class, 'getViewWayfairZeroData']);
    Route::get('/Wayfaire/low-visibility/view-data', [WayfairLowVisibilityController::class, 'getViewWayfairLowVisibilityData']);
    Route::get('/Temu/view-data', [TemuZeroController::class, 'getViewTemuZeroData']);
    Route::get('/reverb/zero/view', [ReverbZeroController::class, 'index'])->name('reverb.zero.view');
    Route::get('/reverb/low-visibility/view', [ReverbLowVisibilityController::class, 'reverbLowVisibilityview'])->name('reverb.low.visibility.view');
    Route::get('/Temu/low-visibility/view-data', [TemuLowVisibilityController::class, 'getViewTemuLowVisibilityData']);
    Route::get('/reverb/zero/view-data', [ReverbZeroController::class, 'getZeroViewData']);
    Route::get('/zero-reverb/view-data', [ReverbZeroController::class, 'getViewReverbZeroData']);
    Route::get('/reverb/zero-low-visibility/view-data', [ReverbLowVisibilityController::class, 'getViewReverbLowVisibilityData']);
    Route::get('/temu/view-data', [TemuController::class, 'getViewTemuData']);
    Route::get('/amazonfba/view-data', [OverallAmazonFbaController::class, 'getViewAmazonFbaData'])->name('amazonfba.viewData');
    Route::get('/fbainv/view-data', [AmazonFbaInvController::class, 'getViewAmazonfbaInvData'])->name('fbainv.viewData');
    Route::get('/product-master-data', [ProductMasterController::class, 'product_master_data']);

    Route::get('/reverb-pricing-cvr', [ReverbController::class, 'reverbPricingCvr'])->name('reverb');
    Route::get('/reverb-pricing-increase-cvr', [ReverbController::class, 'reverbPricingIncreaseCvr'])->name('reverb');
    Route::get('/reverb-pricing-decrease-cvr', [ReverbController::class, 'reverbPricingDecreaseCvr'])->name('reverb');

    Route::post('/reverb/save-sprice', [ReverbController::class, 'saveSpriceToDatabase'])->name('reverb.save-sprice');


    // routes/web.php or routes/api.php
    Route::get('/channel-counts', [ChannelMasterController::class, 'getChannelCounts']);

    Route::get('/home', fn() => view('index'))->name('home');
    Route::get('/product-master', [ProductMasterController::class, 'product_master_index'])
        ->name('product.master');
    Route::get('/catalogue/{first?}/{second?}', [CatalougeManagerController::class, 'catalouge_manager_index'])
        ->name('catalogue.manager');
    //channel index
    Route::get('/channel/promotion-master', [ChannelPromotionMasterController::class, 'channel_promotion_master_index'])
        ->name('promotion.master');

    Route::get('/channel/{firstChannel?}/{secondChannel?}', [ChannelMasterController::class, 'channel_master_index'])
        ->name('channel.master');
    Route::get('/channel-wise/{firstChannelWise?}/{secondChannelWise?}', [ChannelWiseController::class, 'channel_wise_index'])
        ->name('channel.wise');




    //Marketplace index view routes/
    Route::get('/overall-amazon', action: [OverallAmazonController::class, 'overallAmazon'])->name('overall.amazon');
    Route::post('/overallAmazon/saveLowProfit', action: [OverallAmazonController::class, 'saveLowProfit']);
    Route::get('/amazon-pricing-cvr', action: [OverallAmazonController::class, 'amazonPricingCVR'])->name('amazon.pricing.cvr');
    Route::get('/amazon-pricing-increase-decrease', action: [OverallAmazonController::class, 'amazonPriceIncreaseDecrease'])->name('amazon.pricing.increase');
    Route::post('/amazon/save-manual-link', [OverallAmazonController::class, 'saveManualLink'])->name('amazon.saveManualLink');
    Route::get('/amazon-pricing-increase', action: [OverallAmazonController::class, 'amazonPriceIncrease'])->name('amazon.pricing.inc');
    Route::post('/amazon/save-manual-link', [OverallAmazonController::class, 'saveManualLink'])->name('amazon.saveManualLink');
    Route::get('/getFilteredAmazonData', [OverallAmazonController::class, 'getFilteredAmazonData']);
    Route::post('/amazon-analytics/import', [OverallAmazonController::class, 'importAmazonAnalytics'])->name('amazon.analytics.import');
    Route::get('/amazon-analytics/export', [OverallAmazonController::class, 'exportAmazonAnalytics'])->name('amazon.analytics.export');
    Route::get('/amazon-analytics/sample', [OverallAmazonController::class, 'downloadSample'])->name('amazon.analytics.sample');


    //ebay 2 
    Route::get('/zero-ebay2', [Ebay2ZeroController::class, 'ebay2Zeroview'])->name('zero.ebay2');
    Route::get('/zero_ebay2/view-data', [Ebay2ZeroController::class, 'getViewEbay2ZeroData']);
    Route::post('/zero_ebay2/reason-action/update-data', [Ebay2ZeroController::class, 'updateReasonAction']);
    Route::post('/zero_ebay2/save-nr', [Ebay2ZeroController::class, 'saveEbayTwoZeroNR']);
    Route::get('/listing-ebaytwo', [ListingEbayTwoController::class, 'listingEbayTwo'])->name('listing.ebayTwo');
    Route::get('/listing_ebaytwo/view-data', [ListingEbayTwoController::class, 'getViewListingEbayTwoData']);
    Route::post('/listing_ebaytwo/save-status', [ListingEbayTwoController::class, 'saveStatus']);
    Route::post('/listing_ebaytwo/import', [ListingEbayTwoController::class, 'import'])->name('listing_ebaytwo.import');
    Route::get('/listing_ebaytwo/export', [ListingEbayTwoController::class, 'export'])->name('listing_ebaytwo.export');

    Route::get('ebayTwoAnalysis', action: [EbayTwoController::class, 'overallEbay']);
    Route::get('/ebay2/view-data', [EbayTwoController::class, 'getViewEbay2Data']);
    Route::get('ebayTwoPricingCVR', [EbayTwoController::class, 'ebayTwoPricingCVR'])->name('ebayTwo.pricing.cvr');
    Route::post('/update-all-ebay2-skus', [EbayTwoController::class, 'updateAllEbay2Skus']);
    Route::post('/ebay2/save-nr', [EbayTwoController::class, 'saveNrToDatabase']);
    Route::post('/ebay2/update-listed-live', [EbayTwoController::class, 'updateListedLive']);
    Route::post('/ebay2/save-low-profit-count', [EbayTwoController::class, 'saveLowProfit']);
    Route::post('/ebay2-analytics/import', [EbayTwoController::class, 'importEbayTwoAnalytics'])->name('ebay2.analytics.import');
    Route::get('/ebay2-analytics/export', [EbayTwoController::class, 'exportEbayTwoAnalytics'])->name('ebay2.analytics.export');
    Route::get('/ebay2-analytics/sample', [EbayTwoController::class, 'downloadSample'])->name('ebay2.analytics.sample');

    //ebay 3
    Route::get('/zero-ebay3', [Ebay3ZeroController::class, 'ebay3Zeroview'])->name('zero.ebay3');
    Route::get('/zero_ebay3/view-data', [Ebay3ZeroController::class, 'getViewEbay3ZeroData']);
    Route::post('/zero_ebay3/reason-action/update-data', [Ebay3ZeroController::class, 'updateReasonAction']);
    Route::post('/zero_ebay3/save-nr', [Ebay3ZeroController::class, 'saveEbayThreeZeroNR']);
    Route::get('/listing-ebaythree', [ListingEbayThreeController::class, 'listingEbayThree'])->name('listing.ebayThree');
    Route::get('/listing_ebaythree/view-data', [ListingEbayThreeController::class, 'getViewListingEbayThreeData']);
    Route::post('/listing_ebaythree/save-status', [ListingEbayThreeController::class, 'saveStatus']);
    Route::post('/listing_ebaythree/import', [ListingEbayThreeController::class, 'import'])->name('listing_ebaythree.import');
    Route::get('/listing_ebaythree/export', [ListingEbayThreeController::class, 'export'])->name('listing_ebaythree.export');

    Route::get('ebayThreeAnalysis', action: [EbayThreeController::class, 'overallthreeEbay']);
    Route::get('/ebay3/view-data', [EbayThreeController::class, 'getViewEbay3Data']);
    Route::get('ebayThreePricingCVR', [EbayThreeController::class, 'ebayThreePricingCVR'])->name('ebayThree.pricing.cvr');
    Route::post('/update-all-ebay3-skus', [EbayThreeController::class, 'updateAllEbay3Skus']);
    Route::post('/ebay3/save-nr', [EbayThreeController::class, 'saveNrToDatabase']);
    Route::post('/ebay3/update-listed-live', [EbayThreeController::class, 'updateListedLive']);
    Route::post('/ebay3-analytics/import', [EbayThreeController::class, 'importEbayThreeAnalytics'])->name('ebay3.analytics.import');
    Route::get('/ebay3-analytics/export', [EbayThreeController::class, 'exportEbayThreeAnalytics'])->name('ebay3.analytics.export');
    Route::get('/ebay3-analytics/sample', [EbayThreeController::class, 'downloadSample'])->name('ebay3.analytics.sample');

    //walmart
    Route::get('/zero-walmart', [WalmartZeroController::class, 'walmartZeroview'])->name('zero.walmart');
    Route::get('/zero_walmart/view-data', [WalmartZeroController::class, 'getViewWalmartZeroData']);
    Route::post('/zero_walmart/reason-action/update-data', [WalmartZeroController::class, 'updateReasonAction']);
    Route::get('/listing-walmart', [ListingWalmartController::class, 'listingWalmart'])->name('listing.walmart');
    Route::get('/listing_walmart/view-data', [ListingWalmartController::class, 'getViewListingWalmartData']);
    Route::post('/listing_walmart/save-status', [ListingWalmartController::class, 'saveStatus']);
    Route::post('/listing_walmart/import', [ListingWalmartController::class, 'import'])->name('listing_walmart.import');
    Route::get('/listing_walmart/export', [ListingWalmartController::class, 'export'])->name('listing_walmart.export');

    Route::get('walmartAnalysis', action: [WalmartControllerMarket::class, 'overallWalmart']);
    Route::get('/walmart/view-data', [WalmartControllerMarket::class, 'getViewWalmartData']);
    Route::get('walmartPricingCVR', [WalmartControllerMarket::class, 'walmartPricingCVR'])->name('walmart.pricing.cvr');
    Route::post('/update-all-walmart-skus', [WalmartControllerMarket::class, 'updateAllWalmartSkus']);
    Route::post('/walmart/save-nr', [WalmartControllerMarket::class, 'saveNrToDatabase']);
    Route::post('/walmart/update-listed-live', [WalmartControllerMarket::class, 'updateListedLive']);
    Route::post('/walmart-analytics/import', [WalmartControllerMarket::class, 'importWalmartAnalytics'])->name('walmart.analytics.import');
    Route::get('/walmart-analytics/export', [WalmartControllerMarket::class, 'exportWalmartAnalytics'])->name('walmart.analytics.export');
    Route::get('/walmart-analytics/sample', [WalmartControllerMarket::class, 'downloadSample'])->name('walmart.analytics.sample');

    //Listing Audit amazon
    Route::get('/listing-audit-amazon', action: [ListingAuditAmazonController::class, 'listingAuditAmazon'])->name('listing.audit.amazon');
    Route::get('/listing-amazon', [ListingAmazonController::class, 'listingAmazon'])->name('listing.amazon');
    Route::get('/listing_amazon/view-data', [ListingAmazonController::class, 'getViewListingAmazonData']);
    Route::post('/listing_amazon/save-status', [ListingAmazonController::class, 'saveStatus']);
    Route::post('/listing_amazon/import', [ListingAmazonController::class, 'import'])->name('listing_amazon.import');
    Route::get('/listing_amazon/export', [ListingAmazonController::class, 'export'])->name('listing_amazon.export');


    Route::get('/listing-audit-ebay', [ListingAuditEbayController::class, 'listingAuditEbay'])->name('listing.audit.ebay');
    Route::get('/listing-ebay', [ListingEbayController::class, 'listingEbay'])->name('listing.ebay');
    Route::post('/listing_ebay/import', [ListingEbayController::class, 'import'])->name('listing_ebay.import');
    Route::get('/listing_ebay/export', [ListingEbayController::class, 'export'])->name('listing_ebay.export');


    Route::get('/amazon-zero-view', action: [AmazonZeroController::class, 'amazonZero'])->name('amazon.zero.view');
    Route::get('/amazon-low-visibility-view', action: [AmazonLowVisibilityController::class, 'amazonLowVisibility'])->name('amazon.low.visibility.view');
    Route::get('/amazon-low-visibility-view-fba', action: [AmazonLowVisibilityController::class, 'amazonLowVisibilityFba'])->name('amazon.low.visibility.view.fba');
    Route::get('/amazon-low-visibility-view-fbm', action: [AmazonLowVisibilityController::class, 'amazonLowVisibilityFbm'])->name('amazon.low.visibility.view.fbm');
    Route::get('/amazon-low-visibility-view-both', action: [AmazonLowVisibilityController::class, 'amazonLowVisibilityBoth'])->name('amazon.low.visibility.view.both');


    Route::get('/overall-amazon-fba', action: [OverallAmazonFbaController::class, 'overallAmazonFBA'])->name('overall.amazon.fba');
    Route::get('/overall-amazon-fbainv', action: [AmazonFbaInvController::class, 'amazonFbaInv'])->name('overall.amazon.fbainv');

    //Listing Audit ebay
    Route::get('/ebay', [EbayController::class, 'ebayView'])->name('ebay');
    Route::post('/ebay/saveLowProfit', [EbayController::class, 'saveLowProfit']);
    Route::post('/ebay-analytics/import', [EbayController::class, 'importEbayAnalytics'])->name('ebay.analytics.import');
    Route::get('/ebay-analytics/export', [EbayController::class, 'exportEbayAnalytics'])->name('ebay.analytics.export');
    Route::get('/ebay-analytics/sample', [EbayController::class, 'downloadSample'])->name('ebay.analytics.sample');

    Route::any('/update-ebay-sku-pricing', [EbayController::class, 'updateEbayPricing'])->name('ebay.priceUpdate');
    Route::any('/update-ebay2-sku-pricing', [EbayTwoController::class, 'updateEbayPricing'])->name('ebay2.priceUpdate');
    // Route::post('/update-amazon-pricing', [OverallAmazonController::class, 'updatePrice'])->name('amazon.priceUpdate');
    Route::get('/check-amazon-auth', [OverallAmazonController::class, 'checkAmazonAuth']);

    Route::post('/update-fba-status-ebay', [EbayController::class, 'updateFbaStatusEbay'])
        ->name('update.fba.status-ebay');
    Route::get('/ebay-pricing-cvr', [EbayController::class, 'ebayPricingCVR'])->name('ebay.pricing.cvr');


    Route::get('/ebay-pricing-decrease', [EbayController::class, 'ebayPricingIncreaseDecrease'])->name('ebay.pricing.decrease');
    Route::get('/ebay-pricing-increase', action: [EbayController::class, 'ebayPricingIncrease'])->name('ebay.pricing.inc');

    Route::get('/ebay-zero-view', action: [EbayZeroController::class, 'ebayZero'])->name('ebay.zero.view');
    Route::get('/ebay-low-visibility-view', action: [EbayLowVisibilityController::class, 'ebayLowVisibility'])->name('ebay.low.visibility.view');
    Route::get('/ebay2-low-visibility-view', action: [Ebay2LowVisibilityController::class, 'ebay2LowVisibility'])->name('ebay2.low.visibility.view');
    Route::get('/ebay3-low-visibility-view', action: [Ebay3LowVisibilityController::class, 'ebay3LowVisibility'])->name('ebay3.low.visibility.view');


    //Listing Audit Macy
    Route::get('/listing-macys', [ListingMacysController::class, 'listingMacys'])->name('listing.macys');
    Route::get('/listing_macys/view-data', [ListingMacysController::class, 'getViewListingMacysData']);
    Route::post('/listing_macys/save-status', [ListingMacysController::class, 'saveStatus']);
    Route::post('/listing_macys/import', [ListingMacysController::class, 'import'])->name('listing_macys.import');
    Route::get('/listing_macys/export', [ListingMacysController::class, 'export'])->name('listing_macys.export');
    Route::get('/macys', [MacyController::class, 'macyView'])->name('macys');
    Route::post('/macys/saveLowProfit', [MacyController::class, 'saveLowProfit']);
    Route::get('/macys-zero-view', action: [MacyZeroController::class, 'macyZeroView'])->name('macy.zero.view');
    Route::get('/macys-low-visibility-view', action: [MacyLowVisibilityController::class, 'macyLowVisibilityView'])->name('macy.low.visibility.view');
    Route::post('/macys-analytics/import', [MacyController::class, 'importMacysAnalytics'])->name('macys.analytics.import');
    Route::get('/macys-analytics/export', [MacyController::class, 'exportMacysAnalytics'])->name('macys.analytics.export');
    Route::get('/macys-analytics/sample', [MacyController::class, 'downloadSample'])->name('macys.analytics.sample');

    //Listing Audit shopifyB2C
    Route::get('/listing-shopifyb2c', [ListingShopifyB2CController::class, 'listingShopifyB2C'])->name('listing.shopifyb2c');
    Route::get('/listing_shopifyb2c/view-data', [ListingShopifyB2CController::class, 'getViewListingShopifyB2CData']);
    Route::post('/listing_shopifyb2c/save-status', [ListingShopifyB2CController::class, 'saveStatus']);
    Route::post('/listing_shopifyb2c/import', [ListingShopifyB2CController::class, 'import'])->name('listing_shopifyb2c.import');
    Route::get('/listing_shopifyb2c/export', [ListingShopifyB2CController::class, 'export'])->name('listing_shopifyb2c.export');
    Route::get('/shopifyB2C-zero-view', action: [Shopifyb2cZeroController::class, 'shopifyb2cZeroView'])->name('shopifyB2C.zero.view');
    Route::get('/shopifyB2C-low-visibility-view', action: [Shopifyb2cLowVisibilityController::class, 'shopifyb2cLowVisibilityView'])->name('shopifyB2C.low.visibility.view');
    Route::get('/shopifyB2C', [Shopifyb2cController::class, 'shopifyb2cView'])->name('shopifyB2C');
    Route::post('/shopifyb2c/saveLowProfit', [Shopifyb2cController::class, 'saveLowProfit']);
    Route::post('/shopifyb2c-analytics/import', [Shopifyb2cController::class, 'importShopifyB2CAnalytics'])->name('shopifyb2c.analytics.import');
    Route::get('/shopifyb2c-analytics/export', [Shopifyb2cController::class, 'exportShopifyB2CAnalytics'])->name('shopifyb2c.analytics.export');
    Route::get('/shopifyb2c-analytics/sample', [Shopifyb2cController::class, 'downloadSample'])->name('shopifyb2c.analytics.sample');

    //listing Audit Wayfair
    Route::any('/update-wayfair-sku-pricing', [ListingWayfairController::class, 'updatePricing'])->name('wayfair.priceUpdate');
    Route::get('/listing-wayfair', [ListingWayfairController::class, 'listingWayfair'])->name('listing.wayfair');
    Route::get('/listing_wayfair/view-data', [ListingWayfairController::class, 'getViewListingWayfairData']);
    Route::post('/listing_wayfair/save-status', [ListingWayfairController::class, 'saveStatus']);
    Route::post('/listing_wayfair/import', [ListingWayfairController::class, 'import'])->name('listing_wayfair.import');
    Route::get('/listing_wayfair/export', [ListingWayfairController::class, 'export'])->name('listing_wayfair.export');
    Route::get('/Wayfair-zero-view', action: [WayfairZeroController::class, 'wayfairZeroView'])->name('wayfair.zero.view');
    Route::get('/Wayfair-low-visibility-view', action: [WayfairLowVisibilityController::class, 'wayfairLowVisibilityView'])->name('wayfair.low.visibility.view');
    Route::get('/Wayfair', [WayfairController::class, 'wayfairView'])->name('Wayfair');
    Route::post('/wayfair/saveLowProfit', [WayfairController::class, 'saveLowProfit']);
    Route::post('/wayfair-analytics/import', [WayfairController::class, 'importWayfairAnalytics'])->name('wayfair.analytics.import');
    Route::get('/wayfair-analytics/export', [WayfairController::class, 'exportWayfairAnalytics'])->name('wayfair.analytics.export');
    Route::get('/wayfair-analytics/sample', [WayfairController::class, 'downloadSample'])->name('wayfair.analytics.sample');


    //listing Audit Neweggb2c
    Route::get('/neweggB2C', [Neweggb2cController::class, 'neweggB2CView'])->name('neweggB2C');
    Route::post('/neweggB2C/saveLowProfit', [Neweggb2cController::class, 'saveLowProfit']);
    Route::get('/Neweggb2c-zero-view', action: [Neweggb2cZeroController::class, 'neweggB2CZeroView'])->name('neweggb2c.zero.view');
    Route::get('/Neweggb2c-low-visibility-view', action: [Neweggb2cLowVisibilityController::class, 'neweggB2CLowVisibilityView'])->name('neweggb2c.low.visibility.view');
    Route::get('/zero-neweggb2b', [NeweggB2BZeroController::class, 'neweggB2BZeroview'])->name('zero.neweggb2b');
    Route::get('/zero_neweggb2b/view-data', [NeweggB2BZeroController::class, 'getViewNeweggB2BZeroData']);
    Route::post('/zero_neweggb2b/reason-action/update-data', [NeweggB2BZeroController::class, 'updateReasonAction']);
    Route::get('/listing-neweggb2c', [ListingNeweggB2CController::class, 'listingNeweggB2C'])->name('listing.neweggb2c');
    Route::get('/listing_neweggb2c/view-data', [ListingNeweggB2CController::class, 'getViewListingNeweggB2CData']);
    Route::post('/listing_neweggb2c/save-status', [ListingNeweggB2CController::class, 'saveStatus']);

    //listing audit reverb
    Route::get('/listing-reverb', [ListingReverbController::class, 'listingReverb'])->name('listing.reverb');
    Route::get('/listing_reverb/view-data', [ListingReverbController::class, 'getViewListingReverbData']);
    Route::post('/listing_reverb/save-status', [ListingReverbController::class, 'saveStatus']);
    Route::get('/reverb', [ReverbController::class, 'reverbView'])->name('reverb');
    Route::post('/listing_reverb/import', [ListingReverbController::class, 'import'])->name('listing_reverb.import');
    Route::get('/listing_reverb/export', [ListingReverbController::class, 'export'])->name('listing_reverb.export');
    Route::post('/reverb/saveLowProfit', [ReverbController::class, 'saveLowProfit']);
    Route::get('/reverb/zero/view', [ReverbZeroController::class, 'index'])->name('reverb.zero.view');
    Route::get('/reverb-low-visiblity-view', [ReverbLowVisibilityController::class, 'reverbLowVisibilityview'])->name('reverb.low.visibility.view');

    //listing temu
    Route::get('/listing-temu', [ListingTemuController::class, 'listingTemu'])->name('listing.temu');
    Route::get('/listing_temu/view-data', [ListingTemuController::class, 'getViewListingTemuData']);
    Route::post('/listing_temu/save-status', [ListingTemuController::class, 'saveStatus']);
    Route::post('/listing_temu/import', [ListingTemuController::class, 'import'])->name('listing_temu.import');
    Route::get('/listing_temu/export', [ListingTemuController::class, 'export'])->name('listing_temu.export');
    Route::get('/temu', [TemuController::class, 'temuView'])->name('temu');
    Route::get('/temu-pricing-cvr', [TemuController::class, 'temuPricingCVR'])->name('temu.pricing');
    Route::get('/temu-pricing-inc', [TemuController::class, 'temuPricingCVRinc'])->name('temu.pricing.inc');
    Route::get('/temu-pricing-dsc', [TemuController::class, 'temuPricingCVRdsc'])->name('temu.pricing.dsc');

    Route::post('/temu/save-sprice', [TemuController::class, 'saveSpriceToDatabase'])->name('temu.save-sprice');
    Route::post('/temu/save-ship', [TemuController::class, 'saveShipToDatabase']);

    Route::get('/temu-zero-view', [TemuZeroController::class, 'temuZeroView'])->name('temu.zero.view');
    Route::get('/temu-low-visiblity-view', [TemuLowVisibilityController::class, 'temuLowVisibilityView'])->name('temu.low.visibility.view');
    Route::post('/temu-analytics/import', [TemuController::class, 'importTemuAnalytics'])->name('temu.analytics.import');
    Route::get('/temu-analytics/export', [TemuController::class, 'exportTemuAnalytics'])->name('temu.analytics.export');
    Route::get('/temu-analytics/sample', [TemuController::class, 'downloadSample'])->name('temu.analytics.sample');


    // Advertisement Master view routes
    Route::get('/kw-amazon', [KwAmazonController::class, 'Amazon'])->name('advertisment.kw.amazon');
    Route::post('/update-checkbox-flag', [KwAmazonController::class, 'updateCheckboxes']);
    Route::get('/kw-ebay', [KwEbayController::class, 'Ebay'])->name('advertisment.kw.eBay');
    Route::get('/kw-walmart', [WalmartController::class, 'Walmart'])->name('advertisment.kw.walmart');
    Route::get('/prod-target-amazon', [ProdTargetAmazonController::class, 'Amazon'])->name('advertisment.prod.target.Amazon');
    Route::post('/update-all-checkbox', [ProdTargetAmazonController::class, 'updateCheckbox']);
    Route::get('/headline-amazon', [HeadlineAmazonController::class, 'Amazon'])->name('advertisment.headline.Amazon');
    Route::post('/update-checkbox', [HeadlineAmazonController::class, 'update']);
    Route::get('/promoted-ebay', [PromotedEbayController::class, 'Ebay'])->name('advertisment.promoted.eBay');
    Route::get('/google-shopping', [GoogleShoppingController::class, 'GoogleShopping'])->name('advertisment.shopping.google');
    Route::get('/demand-gen-googleNetworks', [GoogleNetworksController::class, 'GoogleNetworks'])->name('advertisment.demand.gen.googleNetworks');
    Route::get('/productwise-fb-img', [ProductWiseMetaParentController::class, 'FacebookImage'])->name('advertisment.demand.productWise.metaParent.img.facebook');
    Route::get('/productwise-insta-img', [ProductWiseMetaParentController::class, 'InstagramImage'])->name('advertisment.demand.productWise.metaParent.img.instagram');
    Route::get('/productwise-fb-video', [ProductWiseMetaParentController::class, 'FacebookVideo'])->name('advertisment.demand.productWise.metaParent.video.facebook');
    Route::get('/productwise-insta-video', [ProductWiseMetaParentController::class, 'InstagramVideo'])->name('advertisment.demand.productWise.metaParent.video.instagram');

    // Ajax Advertisement Master view routes
    Route::get('/kw-ebay-get-data', [KwEbayController::class, 'getViewKwEbayData'])->name('kwEbay.getData');
    Route::post('/update-checkbox-flag', [KwEbayController::class, 'updateCheckboxes']);
    Route::get('/kw-walmart-get-data', [WalmartController::class, 'getViewKwWalmartData'])->name('kwWalmart.getData');
    Route::post('/update-checkbox-flag', [WalmartController::class, 'updateCheckboxes']);
    Route::get('/google-shopping-get-data', [GoogleShoppingController::class, 'getViewGoogleShoppingData'])->name('googleShopping.getData');


    //channel master index view routes
    Route::get('/return-analysis', [ReturnController::class, 'return_master_index'])->name('return.master');
    Route::get('/expenses-analysis', [ExpensesController::class, 'expenses_master_index'])->name('expenses.master');
    Route::get('/review-analysis', [ReviewController::class, 'review_master_index'])->name('review.master');
    Route::get('/health-analysis', [HealthController::class, 'health_master_index'])->name('health.master');

    //product master index view routes
    Route::get('/review.analysis', action: [ReviewAnalysisController::class, 'reviewAnalysis'])->name('review.analysis');
    Route::get('/pricing.analysis', action: [PricingAnalysisController::class, 'pricingAnalysis'])->name('pricing.analysis');
    Route::get('/pRoi.analysis', action: [PrAnalysisController::class, 'pRoiAnalysis'])->name('pRoi.analysis');
    Route::get('/return.analysis', action: [ReturnAnalysisController::class, 'returnAnalysis'])->name('return.analysis');
    Route::get('/stock.analysis', action: [StockAnalysisController::class, 'stockAnalysis'])->name('stock.analysis');
    Route::get('/shortfall.analysis', action: [ShortFallAnalysisController::class, 'shortFallAnalysis'])->name('shortfall.analysis');
    Route::get('/costprice.analysis', action: [CostpriceAnalysisController::class, 'costpriceAnalysis'])->name('costprice.analysis');
    Route::get('/forecast.analysis', action: [ForecastAnalysisController::class, 'forecastAnalysis'])->name('forecast.analysis');

    Route::get('/listing-master', action: [ListingManagerController::class, 'listingmaster'])->name('listing');

    //marketing master index view routes
    Route::get('/listingLQS.master', action: [ListingLQSMasterController::class, 'listingLQSMaster'])->name('listingLQS.master');
    Route::get('/listingLQS/view-data', [ListingLQSMasterController::class, 'getViewListingData'])->name('listingLQS.viewData');
    Route::post('/listing-lqs/save-action', [ListingLQSMasterController::class, 'saveAction']);
    Route::get('/cvrLQS.master', action: [CvrLQSMasterController::class, 'cvrLQSMaster'])->name('cvrLQS.master');
    Route::get('/cvrLQS/view-data', [CvrLQSMasterController::class, 'getViewCvrData'])->name('cvrLQS.viewData');
    Route::post('/cvr-lqs/save-action', [CvrLQSMasterController::class, 'saveAction']);

    Route::post('/import-cvr-data', [CvrLQSMasterController::class, 'importCVRData'])->name('import.cvr');

    Route::get('/lqs-from-sheet', [ListingLQSMasterController::class, 'getLqsFromGoogleSheet']);

    //ebay lqs cvr
    Route::get('/ebaycvrLQS.master', action: [EbayCvrLqsController::class, 'cvrLQSMaster'])->name('ebaycvrLQS.master');
    Route::get('/ebaycvrLQS/view-data', [EbayCvrLqsController::class, 'getViewEbayCvrData'])->name('ebaycvrLQS.viewData');
    Route::post('/ebay-cvr-lqs/save-action', [EbayCvrLqsController::class, 'saveEbayAction']);

    Route::post('/import-ebay-cvr-data', [EbayCvrLqsController::class, 'importEbayCVRData'])->name('import.ebay.cvr');



    //To Be DC routes
    Route::get('/tobedc_list', [ToBeDCController::class, 'index'])->name('tobedc.list');

    //Supplier routes
    Route::get('/supplier.list', [SupplierController::class, 'supplierList'])->name('supplier.list');
    Route::post('/supplier.create', [SupplierController::class, 'postSupplier'])->name('supplier.create');
    Route::delete('/supplier/delete/{id}', [SupplierController::class, 'deleteSupplier'])->name('supplier.delete');
    Route::post('/supplier/import', [SupplierController::class, 'bulkImport'])->name('supplier.import');
    Route::post('/supplier-rating', [SupplierController::class, 'storeRating'])->name('supplier.rating.save');

    //Catategory routes
    Route::get('/category.list', [CategoryController::class, 'categoryList'])->name('category.list');
    Route::post('/category.create', [CategoryController::class, 'postCategory'])->name('category.create');
    Route::delete('/category/delete/{id}', [CategoryController::class, 'destroy'])->name('category.delete');
    Route::post('/category/bulk-delete', [CategoryController::class, 'bulkDelete'])->name('category.bulk-delete');

    //To Order Analysis routes
    Route::controller(ToOrderAnalysisController::class)->group(function () {
        Route::get('/test', 'test')->name('test');
        Route::get('/to-order-analysis', 'index')->name('to.order.analysis');
        Route::get('/to-order-analysis-new', 'toOrderAnalysisNew')->name('to.order.analysis.new');
        Route::get('/to-order-analysis/data', 'getToOrderAnalysis')->name('to.order.analysis.data');
        Route::post('/update-link', 'updateLink')->name('update.rfq.link');
        Route::post('/mfrg-progresses/insert', 'storeMFRG')->name('mfrg.progresses.insert');
        Route::post('/save-to-order-review', 'storeToOrderReview')->name('save.to_order_review');
    });

    //Movement Analysis
    Route::get('/movement.analysis', action: [MovementAnalysisController::class, 'movementAnalysis'])->name('movement.analysis');
    Route::post('/update-smsl', [MovementAnalysisController::class, 'updateSmsl'])->name('update-smsl');

    //Update Forecast Sheet
    Route::post('/update-forecast-data', [ForecastAnalysisController::class, 'updateForcastSheet'])->name('update.forecast.data');
    Route::get('/inventory-stages', [ForecastAnalysisController::class, 'invetoryStagesView'])->name('inventory.stages');
    Route::get('/inventory-stages/data', [ForecastAnalysisController::class, 'invetoryStagesData']);

    //MFRG In Progress
    Route::controller(MFRGInProgressController::class)->group(function () {
        Route::get('/mfrg-in-progress', 'index')->name('mfrg.in.progress');
        Route::post('/mfrg-progresses/inline-update-by-sku', 'inlineUpdateBySku');
        Route::get('/convert-currency', 'convert');
        Route::post('/ready-to-ship/insert', 'storeDataReadyToShip')->name('ready.to.ship.insert');

        Route::get('/mfrg-in-progress/new', 'newMfrgView')->name('mfrg.in.progress.new');
        Route::get('/mfrg-in-progress/data', 'getMfrgProgressData')->name('mfrg.in.progress.data');
    });

    //Ready To Ship
    Route::get('/ready-to-ship', [ReadyToShipController::class, 'index'])->name('ready.to.ship');
    Route::post('/ready-to-ship/inline-update-by-sku', [ReadyToShipController::class, 'inlineUpdateBySku']);
    Route::post('/ready-to-ship/revert-back-mfrg', [ReadyToShipController::class, 'revertBackMfrg']);
    Route::post('/ready-to-ship/move-to-transit', [ReadyToShipController::class, 'moveToTransit']);


    //China Load
    Route::get('/china-load', [ChinaLoadController::class, 'index'])->name('china.load');
    Route::post('/china-load/inline-update-by-sl', [ChinaLoadController::class, 'inlineUpdateBySl']);

    //On Sea Transit
    Route::get('/on-sea-transit', [OnSeaTransitController::class, 'index'])->name('on.sea.transit');
    Route::post('/on-sea-transit/inline-update-or-create', [OnSeaTransitController::class, 'inlineUpdateOrCreate']);

    //On Road Transit
    Route::get('/on-road-transit', [OnRoadTransitController::class, 'index'])->name('on.road.transit');
    Route::post('/on-road-transit/inline-update-or-create', [OnRoadTransitController::class, 'inlineUpdateOrCreate']);

    //Transit Container Details
    Route::get('/transit-container-details', [TransitContainerDetailsController::class, 'index'])->name('transit.container.details');
    Route::post('/transit-container/add-tab', [TransitContainerDetailsController::class, 'addTab']);
    Route::post('/transit-container/save-row', [TransitContainerDetailsController::class, 'saveRow']);
    Route::post('/upload-image', [TransitContainerDetailsController::class, 'uploadImage'])->name('transit.upload-image');
    Route::get('/transit-container-changes', [TransitContainerDetailsController::class, 'transitContainerChanges'])->name('transit.container.changes');
    Route::get('/transit-container-new', [TransitContainerDetailsController::class, 'transitContainerNew'])->name('transit.container.new');
    Route::post('/transit-container/save', [TransitContainerDetailsController::class, 'transitContainerStoreItems']);
    Route::post('/transit-container/delete', [TransitContainerDetailsController::class, 'deleteTransitItem']);

    Route::post('/inventory-warehouse/push', [InventoryWarehouseController::class, 'pushInventory'])->name('inventory.push');
    Route::get('/inventory-warehouse', [InventoryWarehouseController::class, 'index'])->name('inventory.index');


    Route::controller(ArrivedContainerController::class)->group(function () {
        Route::get('/arrived/container', 'index')->name('arrived.container');
        Route::post('/arrived/container/push', 'pushArrivedContainer');
    });


    Route::controller(QualityEnhanceController::class)->group(function () {
        Route::get('/quality-enhance/list', 'index')->name('quality.enhance');
        Route::post('/quality-enhance/get-parent', 'getParentFromSKU')->name('quality.enhance.getParent');
        Route::get('/quality-enhance/data', 'getData')->name('quality.enhance.data');
        Route::post('/quality-enhance/save', 'saveQualityEnhance')->name('quality.enhance.save');
        Route::post('/quality-enhance/update', 'update')->name('quality.enhance.update');
    });

    Route::controller(ContainerPlanningController::class)->group(function () {
        Route::get('/container-planning', 'index')->name('container.planning');
        Route::get('/container-planning/data', 'getContainerPlannings')->name('container.planning.data');
        Route::get('/container-planning/po-details/{id}', 'getPoDetails');
        Route::post('/container-planning/save', 'saveContainerPlanning')->name('container.planning.save');
        Route::post('/container-planning/delete', 'deleteContainerPlanning')->name('container.planning.delete');
    });

    //api data view routes
    Route::get('/shopify/products', [ShopifyController::class, 'getProducts']);

    //data save routes
    Route::post('/product_master/store', [ProductMasterController::class, 'store'])->name('product_master.store');
    Route::post('/product-master/batch-update', [ProductMasterController::class, 'batchUpdate']);
    Route::post('/channel_master/store', [ChannelMasterController::class, 'store'])->name('channel_master.store');
    Route::post('/channel-master/update-sheet-link', [ChannelMasterController::class, 'updateSheetLink']);
    Route::post('/channels-master/toggle-flag', [ChannelMasterController::class, 'toggleCheckboxFlag']);
    Route::post('/update-channel-type', [ChannelMasterController::class, 'updateType']);
    Route::post('/update-channel-percentage', [ChannelMasterController::class, 'updatePercentage']);



    //data update routes
    Route::post('/channel_master/update', [ChannelMasterController::class, 'update']);

    //data delete routes
    Route::delete('/product_master/delete', [ProductMasterController::class, 'destroy'])->name('product_master.destroy');

    //reverb update
    Route::post('/update-reverb-column', [ReverbController::class, 'updateReverbColumn']);

    Route::post('/product-master/import-from-sheet', [ProductMasterController::class, 'importFromSheet']);

    //amazon db save routes
    Route::post('/amazon/save-nr', [OverallAmazonController::class, 'saveNrToDatabase']);
    Route::post('/amazon/update-listed-live', [OverallAmazonController::class, 'updateListedLive']);

    Route::post('/amazon/save-sprice', [OverallAmazonController::class, 'saveSpriceToDatabase'])->name('amazon.save-sprice');

    Route::post('/listing_audit_amazon/save-na', [ListingAuditAmazonController::class, 'saveAuditToDatabase']);
    Route::post('/amazon-zero/reason-action/update', [AmazonZeroController::class, 'updateReasonAction']);
    Route::post('/amazon-low-visibility/reason-action/update', [AmazonLowVisibilityController::class, 'updateReasonAction']);


    // Route::get('/pricing-master.pricing_master', [PricingMasterController::class, 'pricingMaster']);
    // Route::get('/pricing-analysis-data-view', [PricingMasterController::class, 'getViewPricingAnalysisData']);


    Route::get('/pricing-analysis-data-view', [PricingMasterViewsController::class, 'getViewPricingAnalysisData']);
    Route::post('/update-amazon-price', action: [PricingMasterViewsController::class, 'updatePrice'])->name('amazon.priceChange');
    Route::post('/push-shopify-price', action: [PricingMasterViewsController::class, 'pushShopifyPriceBySku'])->name('shopify.priceChange');
    Route::post('/push-ebay-price', action: [PricingMasterViewsController::class, 'pushEbayPriceBySku'])->name('ebay.priceChange');
    Route::post('/push-ebay2-price', action: [PricingMasterViewsController::class, 'pushEbayTwoPriceBySku'])->name('ebay2.priceChange');
    Route::post('/push-ebay3-price', action: [PricingMasterViewsController::class, 'pushEbayThreePriceBySku'])->name('ebay3.priceChange');
    Route::post('/pricing-master/save', [PricingMasterController::class, 'save']);
    Route::post('/pricing-master/save-sprice', [PricingMasterViewsController::class, 'saveSprice']);
    Route::post('/pricing-master/save-remark', [PricingMasterViewsController::class, 'saveRemark']);
    Route::post('/push-walmart-price', [PricingMasterViewsController::class, 'pushPricewalmart']);
    // Route::post('/push-doba-price', [PricingMasterViewsController::class, 'pushdobaPriceBySku']);
    Route::post('/update-doba-price', [PricingMasterViewsController::class, 'pushdobaPriceBySku']); // Added for compatibility
    Route::get('/test-doba-connection', [PricingMasterViewsController::class, 'testDobaConnection']); // Debug route
    Route::post('/update-reverb-price', [PricingMasterViewsController::class, 'updateReverbPrice'])->name('reverb.priceChange');
    Route::post('/update-macy-price', [PricingMasterViewsController::class, 'updateMacyPrice'])->name('macy.priceChange');
    // Route::post('/update-reverb-price', [PricingMasterViewsController::class, 'updateReverbPrice'])->name('reverb.priceChange');




    // Pricing Master Views Roi Dashboard

    Route::get('/pricing-masters.pricing_masters', [PricingMasterViewsController::class, 'pricingMaster']);
    Route::get('/inventory-by-sales-value', [PricingMasterViewsController::class, 'inventoryBySalesValue'])->name('inventory.by.sales.value');
    Route::get('/pricing-master-data-views', [PricingMasterViewsController::class, 'getViewPricingAnalysisData']);
    Route::get('/pricing-master/roi-dashboard', [PricingMasterViewsController::class, 'getViewPricingAnalysisROIDashboardData']);
    Route::post('/pricing-master/save', [PricingMasterViewsController::class, 'save']);
    Route::get('/parent.pricing-masters', [PricingMasterViewsController::class, 'pricingMasterCopy']);
    Route::get('/calculate-cvr-masters', [PricingMasterViewsController::class, 'calculateCVRMasters']);





    Route::get('/movement-pricing-master', [MovementPricingMaster::class, 'MovementPricingMaster']);
    Route::get('/pricing-analysis-data-views', [MovementPricingMaster::class, 'getViewPricingAnalysisData']);
    Route::post('/pricing-master/save', [MovementPricingMaster::class, 'save']);



    Route::get('/ads-pricing-master', [AdsMasterController::class, 'adsMaster']);
    Route::get('/ads-pricing-analysis-data-views', [AdsMasterController::class, 'getViewPricingAnalysisData']);
    Route::post('/pricing-master/save', [AdsMasterController::class, 'save']);




    // Analysis routes
    Route::get('/pricing-master/l30-analysis', [PricingMasterViewsController::class, 'getL30Analysis']);
    Route::get('/pricing-master/site-analysis', [PricingMasterViewsController::class, 'getSiteAnalysis']);
    Route::get('/pricing-master/profit-analysis', [PricingMasterViewsController::class, 'getProfitAnalysis']);
    Route::get('/pricing-master/roi-analysis', [PricingMasterViewsController::class, 'getRoiAnalysis']);

    //ebay db save routes
    Route::post('/ebay/save-nr', [EbayController::class, 'saveNrToDatabase']);
    Route::post('/ebay/update-listed-live', [EbayController::class, 'updateListedLive']);
    Route::post('/ebay/save-sprice', [EbayController::class, 'saveSpriceToDatabase'])->name('ebay.save-sprice');
    Route::post('/ebay/save-sprice', [EbayTwoController::class, 'saveSpriceToDatabase'])->name('ebay.save-sprice');

    Route::post('/listing_ebay/save-status', [ListingEbayController::class, 'saveStatus']);
    Route::post('/listing_audit_ebay/save-na', [ListingAuditEbayController::class, 'saveAuditToDatabase']);
    Route::post('/ebay-zero/reason-action/update', [EbayZeroController::class, 'updateReasonAction']);
    Route::post('/ebay-low-visibility/reason-action/update', [EbayLowVisibilityController::class, 'updateReasonAction']);
    Route::post('/ebay2-low-visibility/reason-action/update', [Ebay2LowVisibilityController::class, 'updateReasonAction']);
    Route::post('/ebay3-low-visibility/reason-action/update', [Ebay3LowVisibilityController::class, 'updateReasonAction']);


    // Shopify B2C route
    Route::get('/listing-audit-shopifyb2c', [ListingAuditShopifyb2cController::class, 'listingAuditShopifyb2c'])->name('listing.audit.shopifyb2c');
    Route::get('/listing_audit_shopifyb2c/view-data', [ListingAuditShopifyb2cController::class, 'getViewListingAuditShopifyb2cData']);
    Route::post('/shopifyb2c/save-nr', [Shopifyb2cController::class, 'saveNrToDatabase']);
    Route::post('/shopifyb2c/update-listed-live', [Shopifyb2cController::class, 'updateListedLive']);
    Route::post('/listing_audit_shopifyb2c/save-na', [ListingAuditShopifyb2cController::class, 'saveAuditToDatabase']);
    Route::post('/shopify/save-sprice', [Shopifyb2cController::class, 'saveSpriceToDatabase']);
    Route::get('/shopify-pricing-cvr', [Shopifyb2cController::class, 'shopifyPricingCvr']);

    Route::get('/shopify-pricing-increase-decrease', [Shopifyb2cController::class, 'shopifyb2cViewPricingIncreaseDecrease']);

    Route::post('/shopifyb2c-zero/reason-action/update', [Shopifyb2cZeroController::class, 'updateReasonAction']);
    Route::post('/shopifyb2c-low-visibility/reason-action/update', [Shopifyb2cLowVisibilityController::class, 'updateReasonAction']);

    // Macy route
    Route::get('/listing-audit-macy', [ListingAuditMacyController::class, 'listingAuditMacy'])->name('listing.audit.macy');
    Route::get('/listing_audit_macy/view-data', [ListingAuditMacyController::class, 'getViewListingAuditMacyData']);
    Route::post('/macy/save-nr', [MacyController::class, 'saveNrToDatabase']);
    Route::post('/macys/save-sprice', [MacyController::class, 'saveSpriceToDatabase'])->name('macy.save-sprice');
    Route::post('/macy/update-listed-live', [MacyController::class, 'updateListedLive']);
    Route::post('/listing_audit_macy/save-na', [ListingAuditMacyController::class, 'saveAuditToDatabase']);
    Route::post('/macy-zero/reason-action/update', [MacyZeroController::class, 'updateReasonAction']);
    Route::post('/macy-low-visibility/reason-action/update', [MacyLowVisibilityController::class, 'updateReasonAction']);

    // Newegg B2C route
    Route::get('/listing-audit-neweggb2c', [ListingAuditNeweggb2cController::class, 'listingAuditNeweggb2c'])->name('listing.audit.neweggb2c');
    Route::get('/listing_audit_neweggb2c/view-data', [ListingAuditNeweggb2cController::class, 'getViewListingAuditNeweggb2cData']);
    Route::post('/neweggb2c/save-nr', [Neweggb2cController::class, 'saveNrToDatabase']);
    Route::post('/listing_audit_neweggb2c/save-na', [ListingAuditNeweggb2cController::class, 'saveAuditToDatabase']);
    Route::post('/neweggb2c-zero/reason-action/update', [Neweggb2cZeroController::class, 'updateReasonAction']);
    Route::post('/neweggb2c-low-visibility/reason-action/update', [Neweggb2cLowVisibilityController::class, 'updateReasonAction']);

    // Wayfaire route 
    Route::get('/listing-audit-wayfair', [ListingAuditWayfairController::class, 'listingAuditWayfair'])->name('listing.audit.wayfair');
    Route::get('/listing_audit_wayfair/view-data', [ListingAuditWayfairController::class, 'getViewListingAuditWayfairData']);
    Route::post('/wayfair/save-nr', [WayfairController::class, 'saveNrToDatabase']);
    Route::post('/wayfair/update-listed-live', [WayfairController::class, 'updateListedLive']);
    Route::post('/listing_audit_wayfair/save-na', [ListingAuditWayfairController::class, 'saveAuditToDatabase']);
    Route::post('/wayfair-zero/reason-action/update', [WayfairZeroController::class, 'updateReasonAction']);
    Route::post('/wayfair-low-visibility/reason-action/update', [WayfairLowVisibilityController::class, 'updateReasonAction']);

    // Reverb route
    Route::get('/listing-audit-reverb', [ListingAuditReverbController::class, 'listingAuditReverb'])->name('listing.audit.reverb');
    Route::get('/listing_audit_reverb/view-data', [ListingAuditReverbController::class, 'getViewListingAuditReverbData']);
    Route::post('/reverb/save-nr', [ReverbController::class, 'saveNrToDatabase']);
    Route::post('/reverb/update-listed-live', [ReverbController::class, 'updateListedLive']);
    Route::post('/listing_audit_reverb/save-na', [ListingAuditReverbController::class, 'saveAuditToDatabase']);
    Route::post('/reverb-zero/reason-action/update', [ReverbZeroController::class, 'updateReasonAction']);
    Route::post('/reverb-low-visibility/reason-action/update', [ReverbLowVisibilityController::class, 'updateReasonAction']);
    Route::post('/reverb-data/import', [ReverbController::class, 'importReverbAnalytics'])->name('reverb.analytics.import');
    Route::get('/reverb-data/export', [ReverbController::class, 'exportReverbAnalytics'])->name('reverb.analytics.export');
    Route::get('/reverb-data/sample', [ReverbController::class, 'downloadSample'])->name('reverb.analytics.sample');


    // Temu route
    Route::get('/listing-audit-temu', [ListingAuditTemuController::class, 'listingAuditTemu'])->name('listing.audit.temu');
    Route::get('/listing_audit_temu/view-data', [ListingAuditTemuController::class, 'getViewListingAuditTemuData']);
    Route::post('/temu/save-nr', [TemuController::class, 'saveNrToDatabase']);
    Route::post('/temu/update-listed-live', [TemuController::class, 'updateListedLive']);
    Route::post('/listing_audit_temu/save-na', [ListingAuditTemuController::class, 'saveAuditToDatabase']);
    Route::post('/temu-zero/reason-action/update', [TemuZeroController::class, 'updateReasonAction']);
    Route::post('/temu-low-visibility/reason-action/update', [TemuLowVisibilityController::class, 'updateReasonAction']);

    // aliExpress route
    Route::get('/zero-aliexpress', [AliexpressZeroController::class, 'aliexpressZeroview'])->name('zero.aliexpress');
    Route::get('/zero_aliexpress/view-data', [AliexpressZeroController::class, 'getViewAliexpressZeroData']);
    Route::post('/zero_aliexpress/reason-action/update-data', [AliexpressZeroController::class, 'updateReasonAction']);
    Route::get('/listing-aliexpress', [ListingAliexpressController::class, 'listingAliexpress'])->name('listing.aliexpress');
    Route::get('/listing_aliexpress/view-data', [ListingAliexpressController::class, 'getViewListingAliexpressData']);
    Route::post('/listing_aliexpress/save-status', [ListingAliexpressController::class, 'saveStatus']);
    Route::post('/listing_aliexpress/import', [ListingAliexpressController::class, 'import'])->name('listing_aliexpress.import');
    Route::get('/listing_aliexpress/export', [ListingAliexpressController::class, 'export'])->name('listing_aliexpress.export');

    Route::get('aliexpressAnalysis', action: [AliexpressController::class, 'overallAliexpress']);
    Route::get('/aliexpress/view-data', [AliexpressController::class, 'getViewAliexpressData']);
    // Route::post('/update-all-aliexpress-skus', [AliexpressController::class, 'updateAllaliexpressSkus']);
    Route::post('/aliexpress/save-nr', [AliexpressController::class, 'saveNrToDatabase']);
    Route::post('/aliexpress/update-listed-live', [AliexpressController::class, 'updateListedLive']);
    Route::post('/aliexpress-analytics/import', [AliexpressController::class, 'importAliexpressAnalytics'])->name('aliexpress.analytics.import');
    Route::get('/aliexpress-analytics/export', [AliexpressController::class, 'exportAliexpressAnalytics'])->name('aliexpress.analytics.export');
    Route::get('/aliexpress-analytics/sample', [AliexpressController::class, 'downloadSample'])->name('aliexpress.analytics.sample');


    // ebay variation
    Route::get('/zero-ebayvariation', [EbayVariationZeroController::class, 'ebayVariationZeroview'])->name('zero.ebayvariation');
    Route::get('/zero_ebayvariation/view-data', [EbayVariationZeroController::class, 'getViewEbayVariationZeroData']);
    Route::post('/zero_ebayvariation/reason-action/update-data', [EbayVariationZeroController::class, 'updateReasonAction']);
    Route::get('/listing-ebayvariation', [ListingEbayVariationController::class, 'listingEbayVariation'])->name('listing.ebayvariation');
    Route::get('/listing_ebayvariation/view-data', [ListingEbayVariationController::class, 'getViewListingEbayVariationData']);
    Route::post('/listing_ebayvariation/save-status', [ListingEbayVariationController::class, 'saveStatus']);
    Route::post('/listing_ebayvariation/import', [ListingEbayVariationController::class, 'import'])->name('listing_ebayvariation.import');
    Route::get('/listing_ebayvariation/export', [ListingEbayVariationController::class, 'export'])->name('listing_ebayvariation.export');

    // shopify wholesale
    Route::get('/zero-shopifywholesale', [ShopifyWholesaleZeroController::class, 'shopifyWholesaleZeroview'])->name('zero.shopifywholesale');
    Route::get('/zero_shopifywholesale/view-data', [ShopifyWholesaleZeroController::class, 'getViewShopifyWholesaleZeroData']);
    Route::post('/zero_shopifywholesale/reason-action/update-data', [ShopifyWholesaleZeroController::class, 'updateReasonAction']);
    Route::get('/listing-shopifywholesale', [ListingShopifyWholesaleController::class, 'listingShopifyWholesale'])->name('listing.shopifywholesale');
    Route::get('/listing_shopifywholesale/view-data', [ListingShopifyWholesaleController::class, 'getViewListingShopifyWholesaleData']);
    Route::post('/listing_shopifywholesale/save-status', [ListingShopifyWholesaleController::class, 'saveStatus']);
    Route::post('/listing_shopifywholesale/import', [ListingShopifyWholesaleController::class, 'import'])->name('listing_shopifywholesale.import');
    Route::get('/listing_shopifywholesale/export', [ListingShopifyWholesaleController::class, 'export'])->name('listing_shopifywholesale.export');

    Route::post('/shopifywholesale/save-nr', [ShopifyWholesaleZeroController::class, 'saveNrToDatabase'])->name('zero.shopifywholesale.save-nr');


    //listing Faire
    Route::get('/zero-faire', [FaireZeroController::class, 'faireZeroview'])->name('zero.faire');
    Route::get('/zero_faire/view-data', [FaireZeroController::class, 'getViewFaireZeroData']);
    Route::post('/zero_faire/reason-action/update-data', [FaireZeroController::class, 'updateReasonAction']);
    Route::get('/listing-faire', [ListingFaireController::class, 'listingFaire'])->name('listing.faire');
    Route::get('/listing_faire/view-data', [ListingFaireController::class, 'getViewListingFaireData']);
    Route::post('/listing_faire/save-status', [ListingFaireController::class, 'saveStatus']);
    Route::post('/listing_faire/import', [ListingFaireController::class, 'import'])->name('listing_faire.import');
    Route::get('/listing_faire/export', [ListingFaireController::class, 'export'])->name('listing_faire.export');


    // listing TiktokShop
    Route::get('/zero-tiktokshop', [TiktokShopZeroController::class, 'tiktokShopZeroview'])->name('zero.tiktokshop');
    Route::get('/zero_tiktokshop/view-data', [TiktokShopZeroController::class, 'getViewTiktokShopZeroData']);
    Route::post('/zero_tiktokshop/reason-action/update-data', [TiktokShopZeroController::class, 'updateReasonAction']);
    Route::get('/listing-tiktokshop', [ListingTiktokShopController::class, 'listingTiktokShop'])->name('listing.tiktokshop');
    Route::get('/listing_tiktokshop/view-data', [ListingTiktokShopController::class, 'getViewListingTiktokShopData']);
    Route::post('/listing_tiktokshop/save-status', [ListingTiktokShopController::class, 'saveStatus']);
    Route::post('/listing_tiktokshop/import', [ListingTiktokShopController::class, 'import'])->name('listing_tiktokshop.import');
    Route::get('/listing_tiktokshop/export', [ListingTiktokShopController::class, 'export'])->name('listing_tiktokshop.export');

    Route::get('tiktokAnalysis', action: [TiktokShopController::class, 'overallTiktok']);
    Route::get('/tiktok/view-data', [TiktokShopController::class, 'getViewTiktokData']);
    Route::get('walmartPricingCVR', [TiktokShopController::class, 'tiktokPricingCVR'])->name('tiktok.pricing.cvr');
    Route::post('/update-all-tiktok-skus', [TiktokShopController::class, 'updateAllTiktokSkus']);
    Route::post('/tiktok/save-nr', [TiktokShopController::class, 'saveNrToDatabase']);
    Route::post('/tiktok/update-listed-live', [TiktokShopController::class, 'updateListedLive']);
    Route::post('/tiktok-analytics/import', [TiktokShopController::class, 'importTiktokAnalytics'])->name('tiktok.analytics.import');
    Route::get('/tiktok-analytics/export', [TiktokShopController::class, 'exportTiktokAnalytics'])->name('tiktok.analytics.export');
    Route::get('/tiktok-analytics/sample', [TiktokShopController::class, 'downloadSample'])->name('tiktok.analytics.sample');

    // listing MercariWShip
    Route::get('/zero-mercariwship', [MercariWShipZeroController::class, 'mercariWShipZeroview'])->name('zero.mercariwship');
    Route::get('/zero_mercariwship/view-data', [MercariWShipZeroController::class, 'getViewMercariWShipZeroData']);
    Route::post('/zero_mercariwship/reason-action/update-data', [MercariWShipZeroController::class, 'updateReasonAction']);
    Route::get('/listing-mercariwship', [ListingMercariWShipController::class, 'listingMercariWShip'])->name('listing.mercariwship');
    Route::get('/listing_mercariwship/view-data', [ListingMercariWShipController::class, 'getViewListingMercariWShipData']);
    Route::post('/listing_mercariwship/save-status', [ListingMercariWShipController::class, 'saveStatus']);
    Route::post('/listing_mercariwship/import', [ListingMercariWShipController::class, 'import'])->name('listing_mercariwship.import');
    Route::get('/listing_mercariwship/export', [ListingMercariWShipController::class, 'export'])->name('listing_mercariwship.export');


    // FBMarketplace
    Route::get('/zero-fbmarketplace', [FBMarketplaceZeroController::class, 'fbMarketplaceZeroview'])->name('zero.fbmarketplace');
    Route::get('/zero_fbmarketplace/view-data', [FBMarketplaceZeroController::class, 'getViewFBMarketplaceZeroData']);
    Route::post('/zero_fbmarketplace/reason-action/update-data', [FBMarketplaceZeroController::class, 'updateReasonAction']);
    Route::get('/listing-fbmarketplace', [ListingFBMarketplaceController::class, 'listingFBMarketplace'])->name('listing.fbmarketplace');
    Route::get('/listing_fbmarketplace/view-data', [ListingFBMarketplaceController::class, 'getViewListingFBMarketplaceData']);
    Route::post('/listing_fbmarketplace/save-status', [ListingFBMarketplaceController::class, 'saveStatus']);
    Route::post('/listing_fbmarketplace/import', [ListingFBMarketplaceController::class, 'import'])->name('listing_fbmarketplace.import');
    Route::get('/listing_fbmarketplace/export', [ListingFBMarketplaceController::class, 'export'])->name('listing_fbmarketplace.export');


    // Business5Core
    Route::get('/zero-business5core', [Business5CoreZeroController::class, 'business5CoreZeroview'])->name('zero.business5core');
    Route::get('/zero_business5core/view-data', [Business5CoreZeroController::class, 'getViewBusiness5CoreZeroData']);
    Route::post('/zero_business5core/reason-action/update-data', [Business5CoreZeroController::class, 'updateReasonAction']);
    Route::get('/listing-business5core', [ListingBusiness5CoreController::class, 'listingBusiness5Core'])->name('listing.business5core');
    Route::get('/listing_business5core/view-data', [ListingBusiness5CoreController::class, 'getViewListingBusiness5CoreData']);
    Route::post('/listing_business5core/save-status', [ListingBusiness5CoreController::class, 'saveStatus']);
    Route::post('/listing_business5core/import', [ListingBusiness5CoreController::class, 'import'])->name('listing_business5core.import');
    Route::get('/listing_business5core/export', [ListingBusiness5CoreController::class, 'export'])->name('listing_business5core.export');


    //  Pls
    Route::get('/zero-pls', [PLSZeroController::class, 'plsZeroview'])->name('zero.pls');
    Route::get('/zero_pls/view-data', [PLSZeroController::class, 'getViewPLSZeroData']);
    Route::post('/zero_pls/reason-action/update-data', [PLSZeroController::class, 'updateReasonAction']);
    Route::get('/listing-pls', [ListingPlsController::class, 'listingPls'])->name('listing.pls');
    Route::get('/listing_pls/view-data', [ListingPlsController::class, 'getViewListingPlsData']);
    Route::post('/listing_pls/save-status', [ListingPlsController::class, 'saveStatus']);
    Route::post('/listing_pls/import', [ListingPlsController::class, 'import'])->name('listing_pls.import');
    Route::get('/listing_pls/export', [ListingPlsController::class, 'export'])->name('listing_pls.export');

    //  AutoDS
    Route::get('/zero-autods', [AutoDSZeroController::class, 'autoDSZeroview'])->name('zero.autods');
    Route::get('/zero_autods/view-data', [AutoDSZeroController::class, 'getViewAutoDSZeroData']);
    Route::post('/zero_autods/reason-action/update-data', [AutoDSZeroController::class, 'updateReasonAction']);
    Route::get('/listing-autods', [ListingAutoDSController::class, 'listingAutoDS'])->name('listing.autods');
    Route::get('/listing_autods/view-data', [ListingAutoDSController::class, 'getViewListingAutoDSData']);
    Route::post('/listing_autods/save-status', [ListingAutoDSController::class, 'saveStatus']);
    Route::post('/listing_autods/import', [ListingAutoDSController::class, 'import'])->name('listing_autods.import');
    Route::get('/listing_autods/export', [ListingAutoDSController::class, 'export'])->name('listing_autods.export');

    // MercariWoShip
    Route::get('/zero-mercariwoship', [MercariWoShipZeroController::class, 'mercariWoShipZeroview'])->name('zero.mercariwoship');
    Route::get('/zero_mercariwoship/view-data', [MercariWoShipZeroController::class, 'getViewMercariWoShipZeroData']);
    Route::post('/zero_mercariwoship/reason-action/update-data', [MercariWoShipZeroController::class, 'updateReasonAction']);
    Route::get('/listing-mercariwoship', [ListingMercariWoShipController::class, 'listingMercariWoShip'])->name('listing.mercariwoship');
    Route::get('/listing_mercariwoship/view-data', [ListingMercariWoShipController::class, 'getViewListingMercariWoShipData']);
    Route::post('/listing_mercariwoship/save-status', [ListingMercariWoShipController::class, 'saveStatus']);
    Route::post('/listing_mercariwoship/import', [ListingMercariWoShipController::class, 'import'])->name('listing_mercariwoship.import');
    Route::get('/listing_mercariwoship/export', [ListingMercariWoShipController::class, 'export'])->name('listing_mercariwoship.export');

    // Poshmark
    Route::get('/zero-poshmark', [PoshmarkZeroController::class, 'poshmarkZeroview'])->name('zero.poshmark');
    Route::get('/zero_poshmark/view-data', [PoshmarkZeroController::class, 'getViewPoshmarkZeroData']);
    Route::post('/zero_poshmark/reason-action/update-data', [PoshmarkZeroController::class, 'updateReasonAction']);
    Route::get('/listing-poshmark', [ListingPoshmarkController::class, 'listingPoshmark'])->name('listing.poshmark');
    Route::get('/listing_poshmark/view-data', [ListingPoshmarkController::class, 'getViewListingPoshmarkData']);
    Route::post('/listing_poshmark/save-status', [ListingPoshmarkController::class, 'saveStatus']);
    Route::post('/listing_poshmark/import', [ListingPoshmarkController::class, 'import'])->name('listing_poshmark.import');
    Route::get('/listing_poshmark/export', [ListingPoshmarkController::class, 'export'])->name('listing_poshmark.export');



    // Tiendamia
    Route::get('/zero-tiendamia', [TiendamiaZeroController::class, 'tiendamiaZeroview'])->name('zero.tiendamia');
    Route::get('/zero_tiendamia/view-data', [TiendamiaZeroController::class, 'getViewTiendamiaZeroData']);
    Route::post('/zero_tiendamia/reason-action/update-data', [TiendamiaZeroController::class, 'updateReasonAction']);
    Route::get('/listing-tiendamia', [ListingTiendamiaController::class, 'listingTiendamia'])->name('listing.tiendamia');
    Route::get('/listing_tiendamia/view-data', [ListingTiendamiaController::class, 'getViewListingTiendamiaData']);
    Route::post('/listing_tiendamia/save-status', [ListingTiendamiaController::class, 'saveStatus']);
    Route::post('/listing_tiendamia/import', [ListingTiendamiaController::class, 'import'])->name('listing_tiendamia.import');
    Route::get('/listing_tiendamia/export', [ListingTiendamiaController::class, 'export'])->name('listing_tiendamia.export');



    // Shein
    Route::get('/zero-shein', [SheinZeroController::class, 'sheinZeroview'])->name('zero.shein');
    Route::get('/zero_shein/view-data', [SheinZeroController::class, 'getViewSheinZeroData']);
    Route::post('/zero_shein/reason-action/update-data', [SheinZeroController::class, 'updateReasonAction']);
    Route::get('/listing-shein', [ListingSheinController::class, 'listingShein'])->name('listing.shein');
    Route::get('/listing_shein/view-data', [ListingSheinController::class, 'getViewListingSheinData']);
    Route::post('/listing_shein/save-status', [ListingSheinController::class, 'saveStatus']);
    Route::post('/listing_shein/import', [ListingSheinController::class, 'import'])->name('listing_shein.import');
    Route::get('/listing_shein/export', [ListingSheinController::class, 'export'])->name('listing_shein.export');

    Route::get('sheinAnalysis', action: [SheinController::class, 'overallShein']);
    Route::get('/shein/view-data', [SheinController::class, 'getViewSheinData']);
    Route::get('sheinPricingCVR', [SheinController::class, 'sheinPricingCVR'])->name('shein.pricing.cvr');
    Route::post('/update-all-shein-skus', [SheinController::class, 'updateAllSheinSkus']);
    Route::post('/shein/save-nr', [SheinController::class, 'saveNrToDatabase']);
    Route::post('/shein/update-listed-live', [SheinController::class, 'updateListedLive']);
    Route::post('/shein-analytics/import', [SheinController::class, 'importSheinAnalytics'])->name('shein.analytics.import');
    Route::get('/shein-analytics/export', [SheinController::class, 'exportSheinAnalytics'])->name('shein.analytics.export');
    Route::get('/shein-analytics/sample', [SheinController::class, 'downloadSample'])->name('shein.analytics.sample');

    //  Spocket
    Route::get('/zero-spocket', [SpocketZeroController::class, 'spocketZeroview'])->name('zero.spocket');
    Route::get('/zero_spocket/view-data', [SpocketZeroController::class, 'getViewSpocketZeroData']);
    Route::post('/zero_spocket/reason-action/update-data', [SpocketZeroController::class, 'updateReasonAction']);
    Route::get('/listing-spocket', [ListingSpocketController::class, 'listingSpocket'])->name('listing.spocket');
    Route::get('/listing_spocket/view-data', [ListingSpocketController::class, 'getViewListingSpocketData']);
    Route::post('/listing_spocket/save-status', [ListingSpocketController::class, 'saveStatus']);

    // Zendrop
    Route::get('/zero-zendrop', [ZendropZeroController::class, 'zendropZeroview'])->name('zero.zendrop');
    Route::get('/zero_zendrop/view-data', [ZendropZeroController::class, 'getViewZendropZeroData']);
    Route::post('/zero_zendrop/reason-action/update-data', [ZendropZeroController::class, 'updateReasonAction']);
    Route::get('/listing-zendrop', [ListingZendropController::class, 'listingZendrop'])->name('listing.zendrop');
    Route::get('/listing_zendrop/view-data', [ListingZendropController::class, 'getViewListingZendropData']);
    Route::post('/listing_zendrop/save-status', [ListingZendropController::class, 'saveStatus']);

    // Syncee
    Route::get('/zero-syncee', [SynceeZeroController::class, 'synceeZeroview'])->name('zero.syncee');
    Route::get('/zero_syncee/view-data', [SynceeZeroController::class, 'getViewSynceeZeroData']);
    Route::post('/zero_syncee/reason-action/update-data', [SynceeZeroController::class, 'updateReasonAction']);
    Route::get('/listing-syncee', [ListingSynceeController::class, 'listingSyncee'])->name('listing.syncee');
    Route::get('/listing_syncee/view-data', [ListingSynceeController::class, 'getViewListingSynceeData']);
    Route::post('/listing_syncee/save-status', [ListingSynceeController::class, 'saveStatus']);
    Route::post('/listing_syncee/import', [ListingSynceeController::class, 'import'])->name('listing_syncee.import');
    Route::get('/listing_syncee/export', [ListingSynceeController::class, 'export'])->name('listing_syncee.export');


    // Offerup
    Route::get('/zero-offerup', [OfferupZeroController::class, 'offerupZeroview'])->name('zero.offerup');
    Route::get('/zero_offerup/view-data', [OfferupZeroController::class, 'getViewOfferupZeroData']);
    Route::post('/zero_offerup/reason-action/update-data', [OfferupZeroController::class, 'updateReasonAction']);
    Route::get('/listing-offerup', [ListingOfferupController::class, 'listingOfferup'])->name('listing.offerup');
    Route::get('/listing_offerup/view-data', [ListingOfferupController::class, 'getViewListingOfferupData']);
    Route::post('/listing_offerup/save-status', [ListingOfferupController::class, 'saveStatus']);

    // listing Newegg B2B
    Route::get('/listing-neweggb2b', [ListingNeweggB2BController::class, 'listingNeweggB2B'])->name('listing.neweggb2b');
    Route::get('/listing_neweggb2b/view-data', [ListingNeweggB2BController::class, 'getViewListingNeweggB2BData']);
    Route::post('/listing_neweggb2b/save-status', [ListingNeweggB2BController::class, 'saveStatus']);

    // Appscenic
    Route::get('/zero-appscenic', [AppscenicZeroController::class, 'appscenicZeroview'])->name('zero.appscenic');
    Route::get('/zero_appscenic/view-data', [AppscenicZeroController::class, 'getViewAppscenicZeroData']);
    Route::post('/zero_appscenic/reason-action/update-data', [AppscenicZeroController::class, 'updateReasonAction']);
    Route::get('/listing-appscenic', [ListingAppscenicController::class, 'listingAppscenic'])->name('listing.appscenic');
    Route::get('/listing_appscenic/view-data', [ListingAppscenicController::class, 'getViewListingAppscenicData']);
    Route::post('/listing_appscenic/save-status', [ListingAppscenicController::class, 'saveStatus']);

    // listing fbshop
    Route::get('/zero-fbshop', [FBShopZeroController::class, 'fbShopZeroview'])->name('zero.fbshop');
    Route::get('/zero_fbshop/view-data', [FBShopZeroController::class, 'getViewFBShopZeroData']);
    Route::post('/zero_fbshop/reason-action/update-data', [FBShopZeroController::class, 'updateReasonAction']);
    Route::get('/listing-fbshop', [ListingFBShopController::class, 'listingFBShop'])->name('listing.fbshop');
    Route::get('/listing_fbshop/view-data', [ListingFBShopController::class, 'getViewListingFBShopData']);
    Route::post('/listing_fbshop/save-status', [ListingFBShopController::class, 'saveStatus']);
    Route::post('/listing_fbshop/import', [ListingFBShopController::class, 'import'])->name('listing_fbshop.import');
    Route::get('/listing_fbshop/export', [ListingFBShopController::class, 'export'])->name('listing_fbshop.export');


    // Instagram Shop
    Route::get('/zero-instagramshop', [InstagramShopZeroController::class, 'instagramShopZeroview'])->name('zero.instagramshop');
    Route::get('/zero_instagramshop/view-data', [InstagramShopZeroController::class, 'getViewInstagramShopZeroData']);
    Route::post('/zero_instagramshop/reason-action/update-data', [InstagramShopZeroController::class, 'updateReasonAction']);
    Route::get('/listing-instagramshop', [ListingInstagramShopController::class, 'listingInstagramShop'])->name('listing.instagramshop');
    Route::get('/listing_instagramshop/view-data', [ListingInstagramShopController::class, 'getViewListingInstagramShopData']);
    Route::post('/listing_instagramshop/save-status', [ListingInstagramShopController::class, 'saveStatus']);
    Route::post('/listing_instagramshop/import', [ListingInstagramShopController::class, 'import'])->name('listing_instagramshop.import');
    Route::get('/listing_instagramshop/export', [ListingInstagramShopController::class, 'export'])->name('listing_instagramshop.export');



    // listing Yamibuy
    Route::get('/zero-yamibuy', [YamibuyZeroController::class, 'yamibuyZeroview'])->name('zero.yamibuy');
    Route::get('/zero_yamibuy/view-data', [YamibuyZeroController::class, 'getViewYamibuyZeroData']);
    Route::post('/zero_yamibuy/reason-action/update-data', [YamibuyZeroController::class, 'updateReasonAction']);
    Route::get('/listing-yamibuy', [ListingYamibuyController::class, 'listingYamibuy'])->name('listing.yamibuy');
    Route::get('/listing_yamibuy/view-data', [ListingYamibuyController::class, 'getViewListingYamibuyData']);
    Route::post('/listing_yamibuy/save-status', [ListingYamibuyController::class, 'saveStatus']);
    Route::post('/listing_yamibuy/import', [ListingYamibuyController::class, 'import'])->name('listing_yamibuy.import');
    Route::get('/listing_yamibuy/export', [ListingYamibuyController::class, 'export'])->name('listing_yamibuy.export');



    // listing DHGate
    Route::get('/zero-dhgate', [DHGateZeroController::class, 'dhgateZeroview'])->name('zero.dhgate');
    Route::get('/zero_dhgate/view-data', [DHGateZeroController::class, 'getViewDHGateZeroData']);
    Route::post('/zero_dhgate/reason-action/update-data', [DHGateZeroController::class, 'updateReasonAction']);
    Route::get('/listing-dhgate', [ListingDHGateController::class, 'listingDHGate'])->name('listing.dhgate');
    Route::get('/listing_dhgate/view-data', [ListingDHGateController::class, 'getViewListingDHGateData']);
    Route::post('/listing_dhgate/save-status', [ListingDHGateController::class, 'saveStatus']);
    Route::post('/listing_dhgate/import', [ListingDHGateController::class, 'import'])->name('listing_dhgate.import');
    Route::get('/listing_dhgate/export', [ListingDHGateController::class, 'export'])->name('listing_dhgate.export');



    // listing Walmart Canada
    Route::get('/zero-swgearexchange', [SWGearExchangeZeroController::class, 'swGearExchangeZeroview'])->name('zero.swgearexchange');
    Route::get('/zero_swgearexchange/view-data', [SWGearExchangeZeroController::class, 'getViewSWGearExchangeZeroData']);
    Route::post('/zero_swgearexchange/reason-action/update-data', [SWGearExchangeZeroController::class, 'updateReasonAction']);
    Route::get('/listing-swgearexchange', [ListingSWGearExchangeController::class, 'listingSWGearExchange'])->name('listing.swgearexchange');
    Route::get('/listing_swgearexchange/view-data', [ListingSWGearExchangeController::class, 'getViewListingSWGearExchangeData']);
    Route::post('/listing_swgearexchange/save-status', [ListingSWGearExchangeController::class, 'saveStatus']);
    Route::post('/listing_swgearexchange/import', [ListingSWGearExchangeController::class, 'import'])->name('listing_swgearexchange.import');
    Route::get('/listing_swgearexchange/export', [ListingSWGearExchangeController::class, 'export'])->name('listing_swgearexchange.export');

    // Permissions
    // Route::get('/permissions', [NewPermissionController::class, 'index'])->name('permissions');
    // Route::post('/permissions/store', [NewPermissionController::class, 'store'])->name('permissions.store');

    // listing Bestbuy USA
    Route::get('/zero-bestbuyusa', [BestbuyUSAZeroController::class, 'bestbuyUSAZeroview'])->name('zero.bestbuyusa');
    Route::get('/bestbuyusa-analytics', [BestbuyUSAZeroController::class, 'bestbuyUSAZeroAnalytics'])->name('zero.bestbuyusa.analytics');
    Route::get('/zero_bestbuyusa/view-data', [BestbuyUSAZeroController::class, 'getViewBestbuyUSAZeroData']);
    Route::post('/zero_bestbuyusa/update-listed-live', [BestbuyUSAZeroController::class, 'updateListedLive']);
    Route::post('/zero_bestbuyusa/reason-action/update-data', [BestbuyUSAZeroController::class, 'updateReasonAction']);
    Route::get('/listing-bestbuyusa', [ListingBestbuyUSAController::class, 'listingBestbuyUSA'])->name('listing.bestbuyusa');
    Route::get('/listing_bestbuyusa/view-data', [ListingBestbuyUSAController::class, 'getViewListingBestbuyUSAData']);
    Route::post('/listing_bestbuyusa/save-status', [ListingBestbuyUSAController::class, 'saveStatus']);
    Route::post('/listing_bestbuyusa/import', [ListingBestbuyUSAController::class, 'import'])->name('listing_bestbuyusa.import');
    Route::get('/listing_bestbuyusa/export', [ListingBestbuyUSAController::class, 'export'])->name('listing_bestbuyusa.export');
    Route::post('/bestbuyusa-analytics/import', [BestbuyUSAZeroController::class, 'importBestBuyUsaAnalytics'])->name('bestbuyusa.analytics.import');
    Route::get('/bestbuyusa-analytics/export', [BestbuyUSAZeroController::class, 'exportBestBuyUsaAnalytics'])->name('bestbuyusa.analytics.export');
    Route::get('/bestbuyusa-analytics/sample', [BestbuyUSAZeroController::class, 'downloadSample'])->name('bestbuyusa.analytics.sample');


    //listing Master
    Route::get('/listing-master', [ListingMasterController::class, 'index'])->name('listingMaster');
    Route::get('/listing-master-data', [ListingMasterController::class, 'getListingMasterData']);

    Route::get('/listing-master-counts', [ListingMasterController::class, 'getListingMasterCountsViews']);
    Route::post('/listing-master-counts-data', [ListingMasterController::class, 'getMarketplacesData']);
    Route::delete('/listing-master/{marketplace}', [ListingMasterController::class, 'destroy'])->name('listing-master.destroy');


    //overall cvr-lqs
    Route::get('/overall-lqs-cvr', [OverallCvrLqsController::class, 'index'])->name('overallLqsCvr');
    Route::get('/lqs-cvr-data', [OverallCvrLqsController::class, 'getCvrLqsData']);

    // Route::get('/listing-master-counts', [OverallCvrLqsController::class, 'getListingMasterCountsViews']);
    // Route::post('/listing-master-counts-data', [OverallCvrLqsController::class, 'getMarketplacesData']);
    Route::delete('/listing-master/{marketplace}', [OverallCvrLqsController::class, 'destroy'])->name('listing-master.destroy');

    // MM video posted route
    Route::controller(VideoPostedController::class)->group(function () {
        Route::get('/markrting-master/video-posted', 'videoPostedView')->name('mm.video.posted');
        Route::get('/videoPosted/view-data', 'getViewVideoPostedData');
        Route::post('/video-posted/save', 'storeOrUpdate')->name('video_posted_value.store_or_update');

        Route::get('/marketing-master/product-video-upload', 'productVideoUploadView')->name('mm.product.video.upload');
        Route::get('/product-video-upload/view-data', 'getProductVideoUploadData');
        Route::post('/product-video-upload/save', 'productVideoUploadUpdate')->name('mm.product.video.upload.save');

        Route::get('/marketing-master/assembly-video-req', 'assemblyVideoReq')->name('mm.assembly.video.posted');
        Route::get('/assembly-video-req/view-data', 'getAssemblyVideoPostedData');
        Route::post('/assembly-video-req/save', 'asseblyStoreOrUpdate')->name('assembly_video_req.store_or_update');

        Route::get('/marketing-master/assembly-video-upload', 'assemblyVideoUploadView')->name('mm.assembly.video.upload');
        Route::get('/assembly-video-upload/view-data', 'getAssemblyVideoUploadData');
        Route::post('/assembly-video-upload/save', 'assemblyVideoUploadUpdate')->name('assembly_video_upload.store_or_update');

        Route::get('/marketing-master/3d-video-req', 'threeDVideoReq')->name('mm.3d.video.posted');
        Route::get('/3d-video-req/view-data', 'getThreeDVideoPostedData');
        Route::post('/3d-video-req/save', 'threeDStoreOrUpdate')->name('3d_video_req.store_or_update');

        Route::get('/marketing-master/3d-video-upload', 'threeDVideoUploadView')->name('mm.3d.video.upload');
        Route::get('/3d-video-upload/view-data', 'getThreeDVideoUploadData');
        Route::post('/3d-video-upload/save', 'threeDVideoUploadUpdate')->name('3d_video_upload.store_or_update');

        Route::get('/marketing-master/360-video-req', 'three60VideoReq')->name('mm.360.video.posted');
        Route::get('/360-video-req/view-data', 'getThree60VideoPostedData');
        Route::post('/360-video-req/save', 'three60StoreOrUpdate')->name('360_video_req.store_or_update');

        Route::get('/marketing-master/360-video-upload', 'three60VideoUploadView')->name('mm.360.video.upload');
        Route::get('/360-video-upload/view-data', 'getThree60VideoUploadData');
        Route::post('/360-video-upload/save', 'three60VideoUploadUpdate')->name('360_video_upload.store_or_update');

        Route::get('/marketing-master/benefits-video-req', 'benefitsVideoReq')->name('mm.benefits.video.posted');
        Route::get('/benefits-video-req/view-data', 'getBenefitsVideoPostedData');
        Route::post('/benefits-video-req/save', 'benefitsStoreOrUpdate')->name('benefits_video_req.store_or_update');

        Route::get('/marketing-master/benefits-video-upload', 'benefitsVideoUploadView')->name('mm.benefits.video.upload');
        Route::get('/benefits-video-upload/view-data', 'getBenefitsVideoUploadData');
        Route::post('/benefits-video-upload/save', 'benefitsVideoUploadUpdate')->name('benefits_video_upload.store_or_update');

        Route::get('/marketing-master/diy-video-req', 'diyVideoReq')->name('mm.diy.video.posted');
        Route::get('/diy-video-req/view-data', 'getDiyVideoPostedData');
        Route::post('/diy-video-req/save', 'diyStoreOrUpdate')->name('diy_video_req.store_or_update');

        Route::get('/marketing-master/diy-video-upload', 'diyVideoUploadView')->name('mm.diy.video.upload');
        Route::get('/diy-video-upload/view-data', 'getDiyVideoUploadData');
        Route::post('/diy-video-upload/save', 'diyVideoUploadUpdate')->name('diy_video_upload.store_or_update');

        Route::get('/marketing-master/shoppable-video-req', 'shoppableVideoReq')->name('mm.shoppable.video.posted');
        Route::get('/shoppable-video-req/view-data', 'getShoppableVideoPostedData');
        Route::post('/shoppable-video-req/save', 'shoppableStoreOrUpdate')->name('shoppable_video_req.store_or_update');

        Route::post('/video-import', 'import')->name('video.import');
    });

    Route::controller(ClaimReimbursementController::class)->group(function () {
        Route::get('/claim-reimbursement', 'index')->name('claim.reimbursement');
        Route::get('/claim-reimbursement/view-data', 'getViewClaimReimbursementData');
        Route::post('/claim-reimbursement/save', 'saveClaimReimbursement')->name('claim.reimbursement.save');
    });

    Route::controller(VideoAdsMasterController::class)->group(function () {
        Route::get('/tiktok-video-ad', 'tiktokIndex')->name('tiktok.ads.master');
        Route::get('/tikotok-video-ads', 'getTikTokVideoAdsData');
        Route::post('/tiktok-video-ads/save', 'saveTiktokVideoAds')->name('tiktok_video_ads.save');

        Route::get('/facebook-video-ad', 'facebookVideoAdView')->name('facebook.ads.master');
        Route::get('/facebook-video-ads', 'getFacebookVideoAdsData');
        Route::post('/facebook-video-ads/save', 'saveFacebookVideoAds')->name('facebook_video_ads.save');

        Route::get('/facebook-feed-ad', 'facebookFeedAdView')->name('facebook.feed.ads.master');
        Route::get('/facebook-feed-ads', 'getFacebookFeedAdsData');
        Route::post('/facebook-feed-ads/save', 'saveFacebookFeedAds')->name('facebook_feed_ads.save');

        Route::get('/facebook-reel-ad', 'facebookReelAdView')->name('facebook.reel.ads.master');
        Route::get('/facebook-reel-ads', 'getFacebookReelAdsData');
        Route::post('/facebook-reel-ads/save', 'saveFacebookReelAds')->name('facebook_reel_ads.save');

        Route::get('/instagram-video-ad', 'InstagramVideoAdView')->name('instagram.ads.master');
        Route::get('/instagram-video-ads', 'getInstagramVideoAdsData');
        Route::post('/instagram-video-ads/save', 'saveInstagramVideoAds')->name('instagram_video_ads.save');

        Route::get('/instagram-feed-ad', 'instagramFeedAdView')->name('instagram.feed.ads.master');


        Route::get('/instagram-feed-ads', 'getInstagramFeedAdsData');
        Route::post('/instagram-feed-ads/save', 'saveInstagramFeedAds')->name('instagram_feed_ads.save');

        Route::get('/instagram-reel-ad', 'instagramReelAdView')->name('instagram.reel.ads.master');
        Route::get('/instagram-reel-ads', 'getInstagramReelAdsData');
        Route::post('/instagram-reel-ads/save', 'saveInstagramReelAds')->name('instagram_reel_ads.save');

        Route::get('/youtube-video-ad', 'youtubeVideoAdView')->name('youtube.ads.master');
        Route::get('/youtube-video-ads', 'getYoutubeVideoAdsData');
        Route::post('/youtube-video-ads/save', 'saveYoutubeVideoAds')->name('youtube_video_ads.save');

        Route::get('/youtube-shorts-ad', 'youtubeShortsAdView')->name('youtube.shorts.ads.master');
        Route::get('/youtube-shorts-ads', 'getYoutubeShortsAdsData');
        Route::post('/youtube-shorts-ads/save', 'saveYoutubeShortsAds')->name('youtube_shorts_ads.save');
    });

    Route::controller(ShoppableVideoController::class)->group(function () {
        Route::get('/shoppable-video/one-ration', 'oneRation')->name('one.ration');
        Route::get('/one-ration-video/view-data', 'getOneRatioVideoData');
        Route::post('/one-ration-video/save', 'saveOneRationVideo');

        Route::get('/shoppable-video/four-ration', 'fourRation')->name('four.ration');
        Route::get('/four-ration-video/view-data', 'getFourRatioVideoData');
        Route::post('/four-ration-video/save', 'saveFourRationVideo');

        Route::get('/shoppable-video/nine-ration', 'nineRation')->name('nine.ration');
        Route::get('/nine-ration-video/view-data', 'getNineRatioVideoData');
        Route::post('/nine-ration-video/save', 'saveNineRationVideo');

        Route::get('/shoppable-video/sixteen-ration', 'sixteenRation')->name('sixteen.ration');
        Route::get('/sixteen-ration-video/view-data', 'getSixteenRatioVideoData');
        Route::post('/sixteen-ration-video/save', 'savesixteenRationVideo');
    });

    Route::controller(CampaignImportController::class)->group(function () {
        Route::get('campaign', 'index')->name('campaign');
        Route::get('/campaign/under-utilised/', 'budgetUnderUtilised')->name('campaign.under');
        Route::get('/campaign/over-utilised/', 'budgetOverUtilised')->name('campaign.over');
        Route::post('/upload-csv', 'upload');
        Route::post('campaigns/update-note', 'updateField')->name('campaigns.update-note');
        Route::post('/campaigns/data', 'getCampaigns')->name('campaigns.data');
        Route::get('/campaigns/list', 'getCampaignsData')->name('campaigns.list');
        // Route::post('/campaign/save', 'storeOrUpdateCampaign')->name('campaign.save');
    });

    Route::controller(AmazonSpBudgetController::class)->group(function () {
        Route::get('/amazon-sp/amz-utilized-bgt-kw', 'amzUtilizedBgtKw')->name('amazon-sp.amz-utilized-bgt-kw');
        Route::get('/amazon-sp/get-amz-utilized-bgt-kw', 'getAmzUtilizedBgtKw');
        Route::post('/update-amazon-sp-bid-price', 'updateAmazonSpBidPrice');
        Route::put('/update-keywords-bid-price', 'updateCampaignKeywordsBid');

        Route::get('/amazon-sp/amz-utilized-bgt-pt', 'amzUtilizedBgtPt')->name('amazon-sp.amz-utilized-bgt-pt');
        Route::get('/amazon-sp/get-amz-utilized-bgt-pt', 'getAmzUtilizedBgtPt');
        Route::put('/update-amazon-sp-targets-bid-price', 'updateCampaignTargetsBid');
        Route::post('/update-amazon-nr-nrl-fba', 'updateNrNRLFba');
    });

    Route::controller(AmazonSbBudgetController::class)->group(function () {
        Route::get('/amazon-sb/amz-utilized-bgt-hl', 'amzUtilizedBgtHl')->name('amazon-sb.amz-utilized-bgt-hl');
        Route::get('/amazon-sb/get-amz-utilized-bgt-hl', 'getAmzUtilizedBgtHl');
        Route::post('/update-amazon-sb-bid-price', 'updateAmazonSbBidPrice');
        Route::put('/amazon-sb/update-keywords-bid-price', 'updateCampaignKeywordsBid');

        Route::get('/amazon-sb/amz-under-utilized-bgt-hl', 'amzUnderUtilizedBgtHl')->name('amazon-sb.amz-under-utilized-bgt-hl');
        Route::get('/amazon-sb/get-amz-under-utilized-bgt-hl', 'getAmzUnderUtilizedBgtHl');
        Route::post('/update-amazon-under-sb-bid-price', 'updateUnderAmazonSbBidPrice');
    });

    Route::controller(AmzUnderUtilizedBgtController::class)->group(function () {
        Route::get('/amazon-sp/amz-under-utilized-bgt-kw', 'amzUnderUtilizedBgtKw')->name('amazon-sp.amz-under-utilized-bgt-kw');
        Route::get('/amazon-sp/get-amz-under-utilized-bgt-kw', 'getAmzUnderUtilizedBgtKw');
        Route::post('/update-amazon-under-utilized-sp-bid-price', 'updateAmazonSpBidPrice');
        Route::put('/update-keywords-bid-price', 'updateCampaignKeywordsBid');
        Route::put('/update-amz-under-targets-bid-price', 'updateCampaignTargetsBid');

        Route::get('/amazon-sp/amz-under-utilized-bgt-pt', 'amzUnderUtilizedBgtPt')->name('amazon-sp.amz-under-utilized-bgt-pt');
        Route::get('/amazon-sp/get-amz-under-utilized-bgt-pt', 'getAmzUnderUtilizedBgtPt');
    });

    Route::controller(AmzCorrectlyUtilizedController::class)->group(function () {
        Route::get('/amazon/correctly-utilized-bgt-kw', 'correctlyUtilizedKw')->name('amazon.amz-correctly-utilized-bgt-kw');
        Route::get('/get-amz-correctly-utilized-bgt-kw', 'getAmzCorrectlyUtilizedBgtKw');

        Route::get('/amazon/correctly-utilized-bgt-hl', 'correctlyUtilizedHl')->name('amazon.amz-correctly-utilized-bgt-hl');
        Route::get('/get-amz-correctly-utilized-bgt-hl', 'getAmzCorrectlyUtilizedBgtHl');

        Route::get('/amazon/correctly-utilized-bgt-pt', 'correctlyUtilizedPt')->name('amazon.amz-correctly-utilized-bgt-pt');
        Route::get('/get-amz-correctly-utilized-bgt-pt', 'getAmzCorrectlyUtilizedBgtPt');
    });

    Route::controller(AmazonAdRunningController::class)->group(function () {
        Route::get('/amazon/ad-running/list', 'index')->name('amazon.ad-running.list');
        Route::get('/amazon/ad-running/data', 'getAmazonAdRunningData');
    });

    Route::controller(AmazonPinkDilAdController::class)->group(function () {
        Route::get('/amazon/pink-dil/kw/ads', 'amazonPinkDilKwAds')->name('amazon.pink.dil.kw.ads');
        Route::get('/amazon/pink-dil/kw/ads/data', 'getAmazonPinkDilKwAdsData');

        Route::get('/amazon/pink-dil/pt/ads', 'amazonPinkDilPtAds')->name('amazon.pink.dil.pt.ads');
        Route::get('/amazon/pink-dil/pt/ads/data', 'getAmazonPinkDilPtAdsData');

        Route::get('/amazon/pink-dil/hl/ads', 'amazonPinkDilHlAds')->name('amazon.pink.dil.hl.ads');
        Route::get('/amazon/pink-dil/hl/ads/data', 'getAmazonPinkDilHlAdsData');
    });

    Route::controller(EbayOverUtilizedBgtController::class)->group(function () {
        Route::get('/ebay-over-uti', 'ebayOverUtilisation')->name('ebay-over-uti');
        Route::get('/ebay/under/utilized', 'ebayOverUtilized')->name('ebay-under-utilize');
        Route::get('/ebay/correctly/utlized', 'ebayCorrectlyUtilized')->name('ebay-correctly-utilize');
        Route::get('/ebay/make-new/campaign/kw', 'ebayMakeCampaignKw')->name('ebay-make-new-campaign-kw');
        Route::get('/ebay/make-new/campaign/kw/data', 'getEbayMakeNewCampaignKw');

        Route::get('/ebay-over-uti/data', 'getEbayOverUtiData')->name('ebay-over-uti-data');
        Route::post('/update-ebay-nr-data', 'updateNrData');
        Route::put('/update-ebay-keywords-bid-price', 'updateKeywordsBidDynamic');
    });

    //FaceBook Adds Manager 
    Route::controller(FacebookAddsManagerController::class)->group(function () {
        Route::get('/facebook-ads-control/data', 'index')->name('facebook.ads.index');
        Route::get('/facebook-web-to-video', 'facebookWebToVideo')->name('facebook.web.to.video');
        Route::get('/facebook-web-to-video-data', 'facebookWebToVideoData')->name('facebook.web.to.video.data');
        Route::get('/fb-img-caraousal-to-web', 'FbImgCaraousalToWeb')->name('fb.img.caraousal.to.web');
        Route::get('/fb-img-caraousal-to-web-data', 'FbImgCaraousalToWebData')->name('fb.img.caraousal.to.web.data');
    });

    Route::controller(AmazonACOSController::class)->group(function () {
        Route::get('/amazon-acos-control/data', 'index')->name('amazon.acos.index');
        Route::get('/amazon-acos-data', 'getAmzonAcOSData');

        Route::get('/amazon-acos-kw-control', 'amazonAcosKwControl')->name('amazon.acos.kw.control');
        Route::get('/amazon-acos-kw-control-data', 'amazonAcosKwControlData')->name('amazon.acos.kw.control.data');
        Route::get('/amazon-acos-hl-control', 'amazonAcosHlControl')->name('amazon.acos.hl.control');
        Route::get('/amazon-acos-hl-control-data', 'amazonAcosHlControlData')->name('amazon.acos.hl.control.data');
        Route::get('/amazon-acos-pt-control', 'amazonAcosPtControl')->name('amazon.acos.pt.control');
        Route::get('/amazon-acos-pt-control-data', 'amazonAcosPtControlData')->name('amazon.acos.pt.control.data');

        Route::put('/update-amazon-campaign-bgt-price', 'updateAmazonCampaignBgt');
        Route::put('/update-amazon-sb-campaign-bgt-price', 'updateAmazonSbCampaignBgt');
    });

    Route::controller(AmazonCampaignReportsController::class)->group(function () {
        Route::get('/amazon/campaign/reports', 'index')->name('amazon.campaign.reports');
        Route::get('/amazon/kw/ads', 'amazonKwAdsView')->name('amazon.kw.ads');
        Route::get('/amazon/kw/ads/data', 'getAmazonKwAdsData');
        Route::get('/amazon-kw-ads/filter', 'filterKwAds')->name('amazonKwAds.filter');
        Route::get('/amazon/pt/ads', 'amazonPtAdsView')->name('amazon.pt.ads');
        
        Route::get('/amazon/pt/ads/data', 'getAmazonPtAdsData');
        Route::get('/amazon/hl/ads', 'amazonHlAdsView')->name('amazon.hl.ads');
        Route::get('/amazon/hl/ads/data', 'getAmazonHlAdsData');

        Route::get('/amazon/campaign/reports/data', 'getAmazonCampaignsData');
    });

    Route::controller(AmazonFbaAdsController::class)->group(function(){
        Route::get('/amazon/fba/over/kw/ads', 'amzFbaUtilizedBgtKw')->name('amazon.fba.over.kw.ads');
        Route::get('/amazon/fba/over/pt/ads', 'amzFbaUtilizedBgtPt')->name('amazon.fba.over.pt.ads');
        Route::get('/amazon/fba/under/kw/ads', 'amzFbaUnderUtilizedBgtKw')->name('amazon.fba.under.kw.ads');
        Route::get('/amazon/fba/under/pt/ads', 'amzFbaUnderUtilizedBgtPt')->name('amazon.fba.under.pt.ads');
        Route::get('/amazon/fba/correct/kw/ads', 'amzFbaCorrectlyUtilizedBgtKw')->name('amazon.fba.correct.kw.ads');
        Route::get('/amazon/fba/correct/pt/ads', 'amzFbaCorrectlyUtilizedBgtPt')->name('amazon.fba.correct.pt.ads');

        Route::get('/amazon/fba/kw/ads/data', 'getAmazonFbaKwAdsData');
        Route::get('/amazon/fba/pt/ads/data', 'getAmazonFbaPtAdsData');
    });

    Route::controller(EbayACOSController::class)->group(function () {
        Route::get('/ebay-over-uti-acos-pink', 'ebayOverUtiAcosPink')->name('ebay-over-uti-acos-pink');
        Route::get('/ebay-over-uti-acos-green', 'ebayOverUtiAcosGreen')->name('ebay-over-uti-acos-green');
        Route::get('/ebay-over-uti-acos-red', 'ebayOverUtiAcosRed')->name('ebay-over-uti-acos-red');

        Route::get('/ebay-under-uti-acos-pink', 'ebayUnderUtiAcosPink')->name('ebay-under-uti-acos-pink');
        Route::get('/ebay-under-uti-acos-green', 'ebayUnderUtiAcosGreen')->name('ebay-under-uti-acos-green');
        Route::get('/ebay-under-uti-acos-red', 'ebayUnderUtiAcosRed')->name('ebay-under-uti-acos-red');

        Route::get('/ebay-uti-acos/data', 'getEbayUtilisationAcosData');
    });

    Route::controller(EbayPinkDilAdController::class)->group(function () {
        Route::get('/ebay/pink-dil/ads', 'index')->name('ebay.pink.dil.ads');
        Route::get('/ebay/pink-dil/ads/data', 'getEbayPinkDilAdsData');
    });

    Route::controller(EbayPMPAdsController::class)->group(function () {
        Route::get('/ebay/pmp/ads', 'index')->name('ebay.pmp.ads');
        Route::get('/ebay/pmp/ads/data', 'getEbayPmpAdsData');
        Route::post('/update-ebay-pmt-percenatge', 'updateEbayPercentage');
        Route::post('/update-ebay-pmt-sprice', 'saveEbayPMTSpriceToDatabase');
    });

    Route::controller(EbayKwAdsController::class)->group(function () {
        Route::get('/ebay/keywords/ads', 'index')->name('ebay.keywords.ads');
        Route::get('/ebay/keywords/ads/data', 'getEbayKwAdsData');

        Route::get('/ebay/keywords/ads/less-than-twenty', 'ebayPriceLessThanTwentyAdsView')->name('ebay.keywords.ads.less-than-twenty');
        Route::get('/ebay/keywords/ads/less-than-twenty/data', 'ebayPriceLessThanTwentyAdsData');
    });

    Route::controller(WalmartUtilisationController::class)->group(function () {
        Route::get('/walmart/utilized/kw', 'index')->name('walmart.utilized.kw');
        Route::get('/walmart/over/utilized', 'overUtilisedView')->name('walmart.over.utilized');
        Route::get('/walmart/under/utilized', 'underUtilisedView')->name('walmart.under.utilized');
        Route::get('/walmart/correctly/utilized', 'correctlyUtilisedView')->name('walmart.correctly.utilized');
        Route::get('/walmart/utilized/kw/data', 'getWalmartAdsData');
    });

    // shopify amazon stock mapping
     Route::controller(StockMappingController::class)->group(function () {
        Route::get('/stock/mapping/view', 'index')->name('view.stock.mapping');
        Route::get('/stock/mapping/inventory/data', 'getShopifyAmazonInventoryStock')->name('stock.mapping.inventory');
        Route::get('/stock/mapping/shopify/data', 'getShopifyStock')->name('stock.mapping.shopify');
        Route::get('/stock/mapping/amazon/data', 'getAmazonStock')->name('stock.mapping.amazon');
        
    });
    // shopify amazon stock mapping

    Route::controller(GoogleShoppingAdsController::class)->group(function () {
        Route::get('/google/shopping', 'index')->name('google.shopping');
        Route::get('/google/shopping/serp', 'googleShoppingSerp')->name('google.shopping.serp');
        Route::get('/google/shopping/pmax', 'googleShoppingPmax')->name('google.shopping.pmax');
        Route::get('/google/shopping/over/utilize', 'googleOverUtilizeView')->name('google.shopping.over.utilize');
        Route::get('/google/shopping/under/utilize', 'googleUnderUtilizeView')->name('google.shopping.under.utilize');
        Route::get('/google/shopping/running', 'googleShoppingAdsRunning')->name('google.shopping.running');
        Route::get('/google/shopping/data', 'getGoogleShoppingAdsData');
    });

    
    Route::post('/channel-promotion/store', [ChannelPromotionMasterController::class, 'storeOrUpdatePromotion']);

  


    Route::get('', [RoutingController::class, 'index'])->name('root');
    Route::get('{firstShop}/{secondShop}', [ShopifyController::class, 'shopifyView'])->name('shopify');
    Route::get('{first}/{second}', [RoutingController::class, 'secondLevel'])->name('second');
    Route::get('{first}/{second}/{third}', [RoutingController::class, 'thirdLevel'])->name('third');
    Route::post('/ebay-product-price-update', [EbayDataUpdateController::class, 'updatePrice'])->name('ebay_product_price_update');
    
    Route::get('{any}', [RoutingController::class, 'root'])->name('any');
        
    
    // Route::post('/auto-stock-balance-store', [AutoStockBalanceController::class, 'store'])->name('autostock.balance.store');
    // Route::get('/auto-stock-balance-data-list', [AutoStockBalanceController::class, 'list']);
});
