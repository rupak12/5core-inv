<?php

namespace App\Http\Controllers\PricingMaster;

use App\Http\Controllers\ApiController;
use App\Http\Controllers\Controller;
use App\Http\Controllers\UpdatePriceApiController;
use App\Models\AliexpressDataView;
use App\Models\AliexpressListingStatus;
use App\Models\AliExpressSheetData;
use App\Models\AmazonListingStatus;
use App\Models\EbayListingStatus;
use App\Models\EbayTwoListingStatus;
use App\Models\EbayThreeListingStatus;
use App\Models\ShopifyB2CListingStatus;
use App\Models\DobaListingStatus;
use App\Models\AmazonDatasheet;
use App\Models\AmazonDataView;
use Exception;
use App\Models\DobaDataView;
use App\Models\DobaMetric;
use App\Models\EbayMetric;
use App\Models\PricingMaster;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\MacyProduct;
use App\Models\ReverbProduct;
use App\Models\TemuProductSheet;
use App\Models\WalmartDataView;
use App\Models\Ebay2Metric;
use App\Models\Ebay3Metric;
use App\Models\EbayDataView;
use App\Models\EbayThreeDataView;
use App\Models\EbayTwoDataView;
use App\Models\MacyDataView;
use App\Models\MacysListingStatus;
use App\Models\ReverbListingStatus;
use App\Models\ReverbViewData;
use App\Models\SheinDataView;
use App\Models\Shopifyb2cDataView;
use App\Models\TemuDataView;
use App\Models\TemuListingStatus;
use App\Models\WalmartListingStatus;
use App\Models\SheinSheetData;
use App\Models\SheinListingStatus;
use App\Models\BestbuyUsaProduct;
use App\Models\BestbuyUSADataView;
use App\Models\BestbuyUSAListingStatus;
use App\Models\CvrLqs;
use App\Models\TemuMetric;
use App\Models\TiendamiaProduct;
use App\Models\TiendamiaDataView;
use App\Models\TiendamiaListingStatus;
use App\Models\TiktokShopDataView;
use App\Models\TiktokShopListingStatus;
use App\Models\TiktokSheet;
use App\Services\AmazonSpApiService;
use App\Services\DobaApiService;
use App\Services\EbayApiService;
use App\Services\WalmartService;
use Illuminate\Contracts\Session\Session;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;
use Illuminate\View\ViewServiceProvider;
use Jasara\AmznSPA\AmznSPA;

class PricingMasterViewsController extends Controller
{
    protected $apiController;
    protected $walmart;
    protected $doba;
    protected $ebay;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
        $this->walmart = new WalmartService();
        $this->doba = new DobaApiService();
        $this->ebay = new EbayApiService();
    }


    public function pricingMaster(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        $processedData = $this->processPricingData();

        return view('pricing-master.pricing_masters_view', [
            'mode' => $mode,
            'demo' => $demo,
            'records' => $processedData,
        ]);
    }

    public function inventoryBySalesValue(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        $processedData = $this->processPricingData();

        return view('pricing-master.inventory_by_sales_value', [
            'mode' => $mode,
            'demo' => $demo,
            'records' => $processedData,
        ]);
    }


    public function calculateCVRMasters(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        $processedData = $this->processPricingData();

        return view('pricing-master.cvr_master', [
            'mode' => $mode,
            'demo' => $demo,

            'records' => $processedData,
        ]);
    }


    public function pricingMasterIncR(Request $request)
        {
            $mode = $request->query('mode');
            $demo = $request->query('demo');

            $processedData = $this->processPricingData();

            return view('pricing-master.pricing_master_incr', [
                'mode' => $mode,
                'demo' => $demo,
                'records' => $processedData,
            ]);
        }




    protected function processPricingData($searchTerm = '')
    {

        $productData = ProductMaster::whereNull('deleted_at')

            ->orderBy('id', 'asc')
            ->get();

        $skus = $productData
            ->pluck('sku')
            ->filter(function ($sku) {
                return stripos($sku, 'PARENT') === false;
            })
            ->unique()
            ->toArray();
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy(function ($item) {
            return trim(strtoupper($item->sku));
        });

        $amazonData  = AmazonDatasheet::whereIn('sku', $skus)->get()->keyBy('sku');
        $amazonListingData = AmazonListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayData    = EbayMetric::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayListingData = EbayListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $temuListingData = TemuListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayTwoListingData = EbayTwoListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayThreeListingData = EbayThreeListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $shopifyb2cListingData = Shopifyb2cListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $dobaListingData = DobaListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $macysListingStatus = MacysListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $reverbListingData = ReverbListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $walmartListingData = WalmartListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $sheinListingData = SheinListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $bestbuyUsaListingData = BestbuyUSAListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $dobaListingData = DobaListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $tiendamiaListingData = TiendamiaListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $tiktokListingData = TiktokShopListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $aliexpressListingData = AliexpressListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');



        // $dobaData    = DobaMetric::whereIn('sku', $skus)->get()->keyBy('sku');
        $pricingData = PricingMaster::whereIn('sku', $skus)->get()->keyBy('sku');
        $macyData    = MacyProduct::whereIn('sku', $skus)->get()->keyBy('sku');
        $reverbData  = ReverbProduct::whereIn('sku', $skus)->get()->keyBy('sku');
        $tiktokLookup = TiktokSheet::whereIn('sku', $skus)->get()->keyBy('sku');
        $walmartLookup = DB::connection('apicentral')
            ->table('walmart_api_data as api')
            ->select(
                'api.sku',
                'api.price',
                DB::raw('COALESCE(m.l30, 0) as l30'),
                DB::raw('COALESCE(m.l60, 0) as l60')
            )
            ->leftJoin('walmart_metrics as m', 'api.sku', '=', 'm.sku')
            ->whereIn('api.sku', $skus)
            ->get()
            ->keyBy('sku');


        $dobaData = DB::connection('apicentral')
            ->table('doba_api_data as api_doba')
            ->select(
                'api_doba.spu as sku',
                'api_doba.sellPrice as doba_price',
                DB::raw('COALESCE(doba_m.l30, 0) as l30'),
                DB::raw('COALESCE(doba_m.l60, 0) as l60')
            )
            ->leftJoin('doba_metrics as doba_m', 'api_doba.spu', '=', 'doba_m.sku')
            ->whereIn('api_doba.spu', $skus)
            ->get()
            ->keyBy('sku');


        // $ebay2Lookup = Ebay2Metric::whereIn('sku', $skus)->get()->keyBy('sku');

        $ebay2Lookup = DB::connection('apicentral')
            ->table('ebay2_metrics')
            ->select('sku', 'ebay_price', 'ebay_l30', 'ebay_l60', 'views')
            ->whereIn('sku', $skus)
            ->get()
            ->keyBy('sku');



        $ebay3Lookup = Ebay3Metric::whereIn('sku', $skus)->get()->keyBy('sku');
        $temuMetricLookup = TemuMetric::whereIn('sku', $skus)->get()->keyBy('sku');
        $amazonDataView = AmazonDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayDataView = EbayDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $shopifyb2cDataView = Shopifyb2cDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $dobaDataView = DobaDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $temuDataView = TemuDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $reverbDataView = ReverbViewData::whereIn('sku', $skus)->get()->keyBy('sku');
        $macyDataView = MacyDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $sheinDataView = SheinDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $sheinData = SheinSheetData::whereIn('sku', $skus)->get()->keyBy('sku');
        $bestbuyUsaLookup = BestbuyUsaProduct::whereIn('sku', $skus)->get()->keyBy('sku');
        $bestbuyUsaDataView = BestbuyUSADataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $tiendamiaLookup = TiendamiaProduct::whereIn('sku', $skus)->get()->keyBy('sku');
        $tiendamiaDataView = TiendamiaDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $aliexpressDataView = AliexpressDataView::whereIn('sku', $skus)->get()->keyBy('sku');
        $aliexpressLookup = AliExpressSheetData::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch LMPA data from 5core_repricer database - get lowest price per SKU (excluding 0 prices)
        $lmpaLookup = collect();
        try {
            $lmpaLookup = DB::connection('repricer')
                ->table('lmpa_data')
                ->select('sku', DB::raw('MIN(price) as lowest_price'))
                ->where('price', '>', 0)
                ->whereIn('sku', $skus)
                ->groupBy('sku')
                ->get()
                ->keyBy('sku');
        } catch (Exception $e) {
            Log::warning('Could not fetch LMPA data from repricer database: ' . $e->getMessage());
            // Fallback to empty collection - will use Amazon's price_lmpa instead
        }

        // Fetch LMP data from 5core_repricer database for eBay - get lowest price per SKU (excluding 0 prices)
        $lmpLookup = collect();
        try {
            $lmpLookup = DB::connection('repricer')
                ->table('lmp_data')
                ->select('sku', DB::raw('MIN(price) as lowest_price'))
                ->where('price', '>', 0)
                ->whereIn('sku', $skus)
                ->groupBy('sku')
                ->get()
                ->keyBy('sku');
        } catch (Exception $e) {
            Log::warning('Could not fetch LMP data from repricer database: ' . $e->getMessage());
            // Fallback to empty collection - will use eBay's price_lmpa instead
        }



        $processedData = [];

        foreach ($productData as $product) {
            $sku = $product->sku;

            if (!empty($searchTerm) && stripos($sku, $searchTerm) === false && stripos($product->parent, $searchTerm) === false) {
                continue;
            }

            $isParent = stripos($sku, 'PARENT') !== false;
            $values = is_string($product->Values) ? json_decode($product->Values, true) : $product->Values;
            if (!is_array($values)) {
                $values = [];
            }

            $msrp = (float) ($values['msrp'] ?? 0);
            $map  = (float) ($values['map'] ?? 0);
            $lp   = (float) ($values['lp'] ?? 0);
            $ship = (float) ($values['ship'] ?? 0);
            $temuship = (float) ($values['temu_ship'] ?? 0);
            $ebay2ship = (float) ($values['ebay2_ship'] ?? 0);
            $initialQuantity = (float) ($values['initial_quantity'] ?? 0);

            $amazon  = $amazonData[$sku] ?? null;
            $ebay    = $ebayData[$sku] ?? null;
            $doba    = $dobaData[$sku] ?? null;
            $pricing = $pricingData[$sku] ?? null;
            $macy    = $macyData[$sku] ?? null;
            $reverb  = $reverbData[$sku] ?? null;
            $temuMetric = $temuMetricLookup[$sku] ?? null;
            $walmart = $walmartLookup[$sku] ?? null;
            $ebay2   = $ebay2Lookup[$sku] ?? null;
            $ebay3   = $ebay3Lookup[$sku] ?? null;
            $lmpa    = $lmpaLookup[$sku] ?? null;
            $lmp     = $lmpLookup[$sku] ?? null;
            $shein   = $sheinData[$sku] ?? null;
            $bestbuyUsa = $bestbuyUsaLookup[$sku] ?? null;
            $tiendamia = $tiendamiaLookup[$sku] ?? null;
            $tiktok = $tiktokLookup[$sku] ?? null;
            $aliexpress = $aliexpressLookup[$sku] ?? null;


            // Get Shopify data for L30 and INV
            $shopifyItem = $shopifyData[trim(strtoupper($sku))] ?? null;
            $inv = $shopifyItem ? ($shopifyItem->inv ?? 0) : 0;
            $l30 = $shopifyItem ? ($shopifyItem->quantity ?? 0) : 0;
            $shopify_l30 = $shopifyItem ? ($shopifyItem->shopify_l30 ?? 0) : 0;

            $total_views = ($amazon ? ($amazon->sessions_l30 ?? 0) : 0) +
                ($ebay ? ($ebay->views ?? 0) : 0) +
                ($ebay2 ? ($ebay2->views ?? 0) : 0) +
                ($ebay3 ? ($ebay3->views ?? 0) : 0) +
                ($shein ? ($shein->views_clicks ?? 0) : 0) +
                ($reverb ? ($reverb->views ?? 0) : 0) +
                ($temuMetric ? ($temuMetric->product_clicks_l30 ?? 0) : 0) +
                ($tiktok ? ($tiktok->views ?? 0) : 0);

            // Calculate total L30 and L60 counts
            $total_l30_count = ($tiktok ? ($tiktok->shopify_tiktokl30 ?? 0) : 0) +
                ($shein ? ($shein->shopify_sheinl30 ?? 0) : 0) +
                ($amazon ? ($amazon->units_ordered_l30 ?? 0) : 0) +
                ($ebay ? ($ebay->ebay_l30 ?? 0) : 0) +
                ($ebay2 ? ($ebay2->ebay_l30 ?? 0) : 0) +
                ($ebay3 ? ($ebay3->ebay_l30 ?? 0) : 0) +
                ($temuMetric ? ($temuMetric->quantity_purchased_l30 ?? 0) : 0) +
                ($reverb ? ($reverb->r_l30 ?? 0) : 0) +
                ($walmart ? ($walmart->l30 ?? 0) : 0) +
                ($macy ? ($macy->m_l30 ?? 0) : 0) +
                ($bestbuyUsa ? ($bestbuyUsa->m_l30 ?? 0) : 0) +
                ($tiendamia ? ($tiendamia->m_l30 ?? 0) : 0) +
                ($doba ? ($doba->l30 ?? 0) : 0) +
                ($aliexpress ? ($aliexpress->aliexpress_l30 ?? 0) : 0);




            $total_l60_count =
                ($tiktok ? ($tiktok->shopify_tiktokl60 ?? 0) : 0) +
                ($shein ? ($shein->shopify_sheinl60 ?? 0) : 0) +
                ($amazon ? ($amazon->units_ordered_l60 ?? 0) : 0) +
                ($ebay ? ($ebay->ebay_l60 ?? 0) : 0) +
                ($ebay2 ? ($ebay2->ebay_l60 ?? 0) : 0) +
                ($ebay3 ? ($ebay3->ebay_l60 ?? 0) : 0) +
                ($temuMetric ? ($temuMetric->quantity_purchased_l60 ?? 0) : 0) +
                ($reverb ? ($reverb->r_l60 ?? 0) : 0) +
                ($walmart ? ($walmart->l60 ?? 0) : 0) +
                ($macy ? ($macy->m_l60 ?? 0) : 0) +
                ($bestbuyUsa ? ($bestbuyUsa->m_l60 ?? 0) : 0) +
                ($tiendamia ? ($tiendamia->m_l60 ?? 0) : 0) +
                ($doba ? ($doba->l60 ?? 0) : 0) +
                ($aliexpress ? ($aliexpress->aliexpress_l60 ?? 0) : 0);



            // Calculate avg CVR only for channels where L30 > 0 and views > 0
            $channels = [
                ['data' => $amazon,     'l30' => 'units_ordered_l30',      'views' => 'sessions_l30'],
                ['data' => $ebay,       'l30' => 'ebay_l30',               'views' => 'views'],
                ['data' => $ebay2,      'l30' => 'ebay_l30',               'views' => 'views'],
                ['data' => $ebay3,      'l30' => 'ebay_l30',               'views' => 'views'],
                ['data' => $temuMetric, 'l30' => 'quantity_purchased_l30', 'views' => 'product_clicks_l30'],
                ['data' => $reverb,     'l30' => 'r_l30',                  'views' => 'views'],
                ['data' => $walmart,    'l30' => 'l30',                    'views' => 'views'],
                ['data' => $tiktok,     'l30' => 'shopify_tiktokl30',                    'views' => 'views'],
                ['data' => $shein,      'l30' => 'shopify_sheinl30',       'views' => 'views_clicks'],
            ];

            $l30_count = 0;        // Count of channels where L30 > 0 and views > 0
            $views_sum = 0;        // Sum of views for those channels

            foreach ($channels as $channel) {
                $obj = $channel['data'];
                if ($obj && is_object($obj)) {
                    $l30 = $obj->{$channel['l30']} ?? 0;
                    $views = $obj->{$channel['views']} ?? 0;

                    if ($l30 > 0 && $views > 0) {
                        $l30_count++;           // count channels
                        $views_sum += $views;   // sum views
                    }
                }
            }



            $total_l30_count_data = 0;

            // Count channels where L30 > 0, ignore views
            if ($amazon && ($amazon->units_ordered_l30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($ebay && ($ebay->ebay_l30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($ebay2 && ($ebay2->ebay_l30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($ebay3 && ($ebay3->ebay_l30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($temuMetric && ($temuMetric->{'quantity_purchased_l30'} ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($reverb && ($reverb->r_l30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            // Walmart excluded as it doesn't have views data
            if ($tiktok && ($tiktok->shopify_tiktokl30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($shein && ($shein->shopify_sheinl30 ?? 0) > 0) {
                $total_l30_count_data++;
            }
            if ($aliexpress && ($aliexpress->aliexpress_l30 ?? 0) > 0) {
                $total_l30_count_data++;
            }

            // For $avgCvr, use all views
            $views_sum =
                ($amazon->sessions_l30 ?? 0) +
                ($ebay->views ?? 0) +
                ($ebay2->views ?? 0) +
                ($ebay3->views ?? 0) +
                ($temuMetric->{'product_clicks_l30'} ?? 0) +
                ($reverb->views ?? 0) +
                ($tiktok->views ?? 0) +
                ($shein->views_clicks ?? 0);

            $avgCvr = $views_sum > 0
                ? number_format(($l30_count / $views_sum) * 100, 1) . ' %'
                : '0.0 %';

            $item = (object) [
                'SKU'     => $sku,
                'Parent'  => $product->parent,
                'remark'  => $product->remark,
                'L30'     => $l30,
                'shopify_l30' => $shopify_l30,
                'total_views' => $total_views,
                'INV'     => $inv,
                'Dil%'    => $inv > 0 ? round((($shopifyItem->quantity ?? 0) / $inv) * 100) : 0,
                'MSRP'    => $msrp,
                'MAP'     => $map,
                'LP'      => $lp,
                'SHIP'    => $ship,
                'temu_ship' => $temuship,
                'ebay2_ship' => $ebay2ship,
                'initial_quantity' => $initialQuantity,
                'is_parent' => $isParent,
                'inv' => $shopifyData[trim(strtoupper($sku))]->inv ?? 0,
                'avgCvr' => $avgCvr,
                'total_l30_count' => $total_l30_count,
                'total_l30_count_data' => $total_l30_count_data,
                'total_l60_count' => $total_l60_count,

                'initial_cogs' => $lp != 0 ? $initialQuantity * $lp : 0,
                'current_cogs' => $lp != 0 ? $inv * $lp : 0,
                // 'avg_inventory' will be set after $item is created

                // Amazon
                'amz_price' => $amazon ? ($amazon->price ?? 0) : 0,
                'amz_l30'   => $amazon ? ($amazon->units_ordered_l30 ?? 0) : 0,
                'amz_l60'   => $amazon ? ($amazon->units_ordered_l60 ?? 0) : 0,
                'sessions_l30' => $amazon ? ($amazon->sessions_l30 ?? 0) : 0,
                'amz_cvr'   => $amazon ? $this->calculateCVR((float)($amazon->units_ordered_l30 ?? 0), (float)($amazon->sessions_l30 ?? 0)) : null,
                'amz_buyer_link' => isset($amazonListingData[$sku]) ? ($amazonListingData[$sku]->value['buyer_link'] ?? null) : null,
                'amz_seller_link' => isset($amazonListingData[$sku]) ? ($amazonListingData[$sku]->value['seller_link'] ?? null) : null,

                'price_lmpa' => $lmpa ? ($lmpa->lowest_price ?? 0) : ($amazon ? ($amazon->price_lmpa ?? 0) : 0),
                'amz_pft'   => $amazon && ($amazon->price ?? 0) > 0 ? (($amazon->price * 0.70 - $lp - $ship) / $amazon->price) : 0,
                'amz_roi'   => $amazon && $lp > 0 && ($amazon->price ?? 0) > 0 ? (($amazon->price * 0.70 - $lp - $ship) / $lp) : 0,
                'amz_req_view' => $amazon && $amazon->sessions_l30 > 0 && $amazon->units_ordered_l30 > 0
                    ? (($inv / 90) * 30) / (($amazon->units_ordered_l30 / $amazon->sessions_l30))
                    : 0,


                // eBay
                'ebay_price' => $ebay ? ($ebay->ebay_price ?? 0) : 0,
                'ebay_l30'   => $ebay ? ($ebay->ebay_l30 ?? 0) : 0,
                'ebay_l60'   => $ebay ? ($ebay->ebay_l60 ?? 0) : 0,
                'ebay_views' => $ebay ? ($ebay->views ?? 0) : 0,
                'ebay_price_lmpa' => $lmp ? ($lmp->lowest_price ?? 0) : ($ebay ? ($ebay->price_lmpa ?? 0) : 0),
                'ebay_cvr'   => $ebay ? $this->calculateCVR($ebay->ebay_l30 ?? 0, $ebay->views ?? 0) : null,
                'ebay_pft'   => $ebay && ($ebay->ebay_price ?? 0) > 0 ? (($ebay->ebay_price * 0.72 - $lp - $ship) / $ebay->ebay_price) : 0,
                'ebay_roi'   => $ebay && $lp > 0 && ($ebay->ebay_price ?? 0) > 0 ? (($ebay->ebay_price * 0.72 - $lp - $ship) / $lp) : 0,

                'ebay_req_view' => $ebay && $ebay->views > 0 && $ebay->ebay_l30 > 0
                    ? (($inv / 90) * 30) / (($ebay->ebay_l30 / $ebay->views))
                    : 0,
                'ebay_buyer_link' => isset($ebayListingData[$sku]) ? ($ebayListingData[$sku]->value['buyer_link'] ?? null) : null,
                'ebay_seller_link' => isset($ebayListingData[$sku]) ? ($ebayListingData[$sku]->value['seller_link'] ?? null) : null,
                // 'ebay_buyer_link' --- IGNORE ---


                // Doba
                'doba_price' => $doba ? ($doba->doba_price ?? 0) : 0,
                'doba_l30'   => $doba ? ($doba->l30 ?? 0) : 0,
                'doba_l60'   => $doba ? ($doba->l60 ?? 0) : 0,

                'doba_pft'   => $doba && ($doba->doba_price ?? 0) > 0 ? (($doba->doba_price * 0.95 - $lp - $ship) / $doba->doba_price) : 0,
                'doba_roi'   => $doba && $lp > 0 && ($doba->doba_price ?? 0) > 0 ? (($doba->doba_price * 0.95 - $lp - $ship) / $lp) : 0,
                'doba_buyer_link' => isset($dobaListingData[$sku]) ? ($dobaListingData[$sku]->value['buyer_link'] ?? null) : null,
                'doba_seller_link' => isset($dobaListingData[$sku]) ? ($dobaListingData[$sku]->value['seller_link'] ?? null) : null,

                // Macy
                'macy_price' => $macy ? ($macy->price ?? 0) : 0,
                'macy_l30'   => $macy ? ($macy->m_l30 ?? 0) : 0,
                'macy_pft'   => $macy && $macy->price > 0 ? (($macy->price * 0.77 - $lp - $ship) / $macy->price) : 0,
                'macy_roi'   => $macy && $lp > 0 && $macy->price > 0 ? (($macy->price * 0.77 - $lp - $ship) / $lp) : 0,
                'macy_buyer_link' => isset($macysListingStatus[$sku]) ? ($macysListingStatus[$sku]->value['buyer_link'] ?? null) : null,
                'macy_seller_link' => isset($macysListingStatus[$sku]) ? ($macysListingStatus[$sku]->value['seller_link'] ?? null) : null,

                // Reverb
                'reverb_price' => $reverb ? ($reverb->price ?? 0) : 0,
                'reverb_l30'   => $reverb ? ($reverb->r_l30 ?? 0) : 0,
                'reverb_l60'   => $reverb ? ($reverb->r_l60 ?? 0) : 0,
                'reverb_views' => $reverb ? ($reverb->views ?? 0) : 0,
                'reverb_pft'   => $reverb && $reverb->price > 0 ? (($reverb->price * 0.77 - $lp - $ship) / $reverb->price) : 0,
                'reverb_roi'   => $reverb && $lp > 0 && $reverb->price > 0 ? (($reverb->price * 0.77 - $lp - $ship) / $lp) : 0,
                'reverb_req_view' => $reverb && $reverb->views > 0 && $reverb->r_l30 > 0 ? (($inv / 90) * 30) / (($reverb->r_l30 / $reverb->views)) : 0,
                'reverb_cvr' => $reverb ? $this->calculateCVR($reverb->r_l30 ?? 0, $reverb->views ?? 0) : null,

                'reverb_buyer_link' => isset($reverbListingData[$sku]) ? ($reverbListingData[$sku]->value['buyer_link'] ?? null) : null,
                'reverb_seller_link' => isset($reverbListingData[$sku]) ? ($reverbListingData[$sku]->value['seller_link'] ?? null) : null,

                // Temu
                'temu_price' => $temuMetric ? (float) ($temuMetric->{'temu_sheet_price'} ?? 0) : 0,
                'temu_l30'   => $temuMetric ? (float) ($temuMetric->{'quantity_purchased_l30'} ?? 0) : 0,
                'temu_l60'   => $temuMetric ? (float) ($temuMetric->{'quantity_purchased_l60'} ?? 0) : 0,
                'temu_dil'   => $temuMetric ? (float) ($temuMetric->{'dil'} ?? 0) : 0,
                'temu_views' => $temuMetric ? (float) ($temuMetric->{'product_clicks_l30'} ?? 0) : 0,

                'temu_pft'   => $temuMetric && ($temuMetric->temu_sheet_price ?? 0) > 0 ? (($temuMetric->temu_sheet_price * 0.87 - $lp - $temuship) / $temuMetric->temu_sheet_price) : 0,
                'temu_roi'   => $temuMetric && $lp > 0 && ($temuMetric->temu_sheet_price ?? 0) > 0 ? (($temuMetric->temu_sheet_price * 0.87 - $lp - $temuship) / $lp) : 0,
                'temu_cvr' => $temuMetric
                    ? $this->calculateCVR($temuMetric->{'quantity_purchased_l30'} ?? 0, $temuMetric->{'product_clicks_l30'} ?? 0)
                    : null,

                'temu_req_view' => $temuMetric && ($temuMetric->{'quantity_purchased_l30'} ?? 0) > 0
                    ? ($inv * 20)
                    : 0,
                'temu_buyer_link' => isset($temuListingData[$sku]) ? ($temuListingData[$sku]->value['buyer_link'] ?? null) : null,
                'temu_seller_link' => isset($temuListingData[$sku]) ? ($temuListingData[$sku]->value['seller_link'] ?? null) : null,

                // Walmart
                'walmart_price' => $walmart ? ($walmart->price ?? 0) : 0,
                'walmart_l30'   => $walmart ?  ($walmart->l30 ?? 0) : 0,
                'walmart_l60'   => $walmart ? ($walmart->l60 ?? 0) : 0,
                'walmart_dil'   => $walmart ?    ($walmart->dil ?? 0) : 0,
                'walmart_pft'   => $walmart && ($walmart->price ?? 0) > 0 ? (($walmart->price * 0.80 - $lp - $ship) / $walmart->price) : 0,
                'walmart_roi'   => $walmart && $lp > 0 && ($walmart->price ?? 0) > 0 ? (($walmart->price * 0.80 - $lp - $ship) / $lp) : 0,
                'walmart_buyer_link' => isset($walmartListingData[$sku]) ? ($walmartListingData[$sku]->value['buyer_link'] ?? null) : null,
                'walmart_seller_link' => isset($walmartListingData[$sku]) ? ($walmartListingData[$sku]->value['seller_link'] ?? null) : null,

                // eBay2
                'ebay2_price' => $ebay2 ? ($ebay2->ebay_price ?? 0) : 0,
                'ebay2_l30'   => $ebay2 ? ($ebay2->ebay_l30 ?? 0) : 0,
                'ebay2_l60'   => $ebay2 ? ($ebay2->ebay_l60 ?? 0) : 0,
                'ebay2_views' => $ebay2 ? ($ebay2->views ?? 0) : 0,
                'ebay2_dil'   => $ebay2 ? (float) ($ebay2->{'dil'} ?? 0) : 0,
                'ebay2_pft'   => $ebay2 && ($ebay2->ebay_price ?? 0) > 0 ? (($ebay2->ebay_price * 0.80 - $lp - $ship) / $ebay2->ebay_price) : 0,
                'ebay2_roi'   => $ebay2 && $lp > 0 && ($ebay2->ebay_price ?? 0) > 0 ? (($ebay2->ebay_price * 0.80 - $lp - $ship) / $lp) : 0,
                'ebay2_req_view' => $ebay2 && $ebay2->views > 0 && $ebay2->ebay_l30
                    ? (($inv / 90) * 30) / (($ebay2->ebay_l30 / $ebay2->views))
                    : 0,



                'ebay2_buyer_link' => isset($ebayTwoListingData[$sku]) ? ($ebayTwoListingData[$sku]->value['buyer_link'] ?? null) : null,
                'ebay2_seller_link' => isset($ebayTwoListingData[$sku]) ? ($ebayTwoListingData[$sku]->value['seller_link'] ?? null) : null,

                // eBay3
                'ebay3_price' => $ebay3 ? ($ebay3->ebay_price ?? 0) : 0,
                'ebay3_l30'   => $ebay3 ? ($ebay3->ebay_l30 ?? 0) : 0,
                'ebay3_l60'   => $ebay3 ? ($ebay3->ebay_l60 ?? 0) : 0,
                'ebay3_views' => $ebay3 ? ($ebay3->views ?? 0) : 0,
                'ebay3_dil'   => $ebay3 ? (float) ($ebay3->{'dil'} ?? 0) : 0,
                'ebay3_cvr'   => $ebay3 ? $this->calculateCVR($ebay3->ebay_l30 ?? 0, $ebay3->views ?? 0) : null,
                'ebay3_pft'   => $ebay3 && ($ebay3->ebay_price ?? 0) > 0 ? (($ebay3->ebay_price * 0.71 - $lp - $ship) / $ebay3->ebay_price) : 0,
                'ebay3_roi'   => $ebay3 && $lp > 0 && ($ebay3->ebay_price ?? 0) > 0 ? (($ebay3->ebay_price * 0.71 - $lp - $ship) / $lp) : 0,
                'ebay3_req_view' => $ebay3 && $ebay3->views && $ebay3->ebay_l30
                    ? (($inv / 90) * 30) / (($ebay3->ebay_l30 / $ebay3->views))
                    : 0,

                'ebay3_buyer_link' => isset($ebayThreeListingData[$sku]) ? ($ebayThreeListingData[$sku]->value['buyer_link'] ?? null) : null,
                'ebay3_seller_link' => isset($ebayThreeListingData[$sku]) ? ($ebayThreeListingData[$sku]->value['seller_link'] ?? null) : null,

                'shopifyb2c_buyer_link' => isset($shopifyb2cListingData[$sku]) ? ($shopifyb2cListingData[$sku]->value['buyer_link'] ?? null) : null,
                'shopifyb2c_seller_link' => isset($shopifyb2cListingData[$sku]) ? ($shopifyb2cListingData[$sku]->value['seller_link'] ?? null) : null,

                // Shein
                'shein_price' => $shein ? ($shein->price ?? 0) : 0,
                'shein_l30'   => $shein ? ($shein->shopify_sheinl30 ?? $shein->l30 ?? 0) : 0,
                'shein_l60'   => $shein ? ($shein->shopify_sheinl60 ?? $shein->l60 ?? 0) : 0,
                'shein_dil'   => $shein ? ($shein->dil ?? 0) : 0,
                'shein_pft'   => $shein && ($shein->price ?? 0) > 0 ? (($shein->price * 0.89 - $lp - $ship) / $shein->price) : 0,
                'shein_roi'   => $shein && $lp > 0 && ($shein->price ?? 0) > 0 ? (($shein->price * 0.89 - $lp - $ship) / $lp) : 0,
                'shein_req_view' => $shein && $shein->views && $shein->l30 ? (($inv / 90) * 30) / (($shein->l30 / $shein->views)) : 0,
                'shein_buyer_link' => isset($sheinListingData[$sku]) ? ($sheinListingData[$sku]->value['buyer_link'] ?? null) : null,
                'shein_seller_link' => isset($sheinListingData[$sku]) ? ($sheinListingData[$sku]->value['seller_link'] ?? null) : null,
                'shein_link1' => $shein ? ($shein->link1 ?? null) : null,
                'shein_cvr' => $shein ? $this->calculateCVR($shein->shopify_sheinl30 ?? 0, ($shein->views_clicks ?? 0) * 3.7) : null,

                // Bestbuy
                'bestbuy_price' => $bestbuyUsa ? ($bestbuyUsa->price ?? 0) : 0,
                'bestbuy_l30'   => $bestbuyUsa ? ($bestbuyUsa->m_l30 ?? 0) : 0,
                'bestbuy_l60'   => $bestbuyUsa ? ($bestbuyUsa->m_l60 ?? 0) : 0,
                'bestbuy_pft'   => $bestbuyUsa && ($bestbuyUsa->price ?? 0) > 0 ? (($bestbuyUsa->price * 0.80 - $lp - $ship) / $bestbuyUsa->price) : 0,
                'bestbuy_roi'   => $bestbuyUsa && $lp > 0 && ($bestbuyUsa->price ?? 0) > 0 ? (($bestbuyUsa->price * 0.80 - $lp - $ship) / $lp) : 0,
                'bestbuy_req_view' => 0, // No views data
                'bestbuy_cvr' => null, // No views data
                'bestbuy_buyer_link' => isset($bestbuyUsaListingData[$sku]) ? ($bestbuyUsaListingData[$sku]->value['buyer_link'] ?? null) : null,
                'bestbuy_seller_link' => isset($bestbuyUsaListingData[$sku]) ? ($bestbuyUsaListingData[$sku]->value['seller_link'] ?? null) : null,

                // Tiendamia
                'tiendamia_price' => $tiendamia ? ($tiendamia->price ?? 0) : 0,
                'tiendamia_l30'   => $tiendamia ? ($tiendamia->m_l30 ?? 0) : 0,
                'tiendamia_l60'   => $tiendamia ? ($tiendamia->m_l60 ?? 0) : 0,
                'tiendamia_pft'   => $tiendamia && ($tiendamia->price ?? 0) > 0 ? (($tiendamia->price * 0.83 - $lp - $ship) / $tiendamia->price) : 0,
                'tiendamia_roi'   => $tiendamia && $lp > 0 && ($tiendamia->price ?? 0) > 0 ? (($tiendamia->price * 0.83 - $lp - $ship) / $lp) : 0,
                'tiendamia_req_view' => 0, // No views data
                'tiendamia_cvr' => null, // No views data
                'tiendamia_buyer_link' => isset($tiendamiaListingData[$sku]) ? ($tiendamiaListingData[$sku]->value['buyer_link'] ?? null) : null,
                'tiendamia_seller_link' => isset($tiendamiaListingData[$sku]) ? ($tiendamiaListingData[$sku]->value['seller_link'] ?? null) : null,

                // TikTok                          
                'tiktok_price' => $tiktok ? ($tiktok->price ?? 0) : 0,
                'tiktok_l30'   => $tiktok ? ($tiktok->shopify_tiktokl30 ?? 0) : 0,
                'tiktok_l60'   => $tiktok ? ($tiktok->shopify_tiktokl60 ?? 0) : 0,
                'tiktok_views' => $tiktok ? ($tiktok->views ?? 0) : 0,
                'tiktok_pft'   => $tiktok && ($tiktok->price ?? 0) > 0 ? (($tiktok->price * 0.64 - $lp - $ship) / $tiktok->price) : 0,
                'tiktok_roi'   => $tiktok && $lp > 0 && ($tiktok->price ?? 0) > 0 ? (($tiktok->price * 0.64 - $lp - $ship) / $lp) : 0,

                'tiktok_req_view' => $tiktok && $tiktok->views > 0 && $tiktok->shopify_tiktokl30
                    ? (($inv / 90) * 30) / (($tiktok->shopify_tiktokl30 / $tiktok->views))
                    : 0,
                'tiktok_cvr'   => $tiktok ? $this->calculateCVR($tiktok->shopify_tiktokl30 ?? 0, $tiktok->views ?? 0) : null,
                'tiktok_buyer_link' => isset($tiktokListingData[$sku]) ? ($tiktokListingData[$sku]->value['buyer_link'] ?? null) : null,
                'tiktok_seller_link' => isset($tiktokListingData[$sku]) ? ($tiktokListingData[$sku]->value['seller_link'] ?? null) : null,

                // AliExpress
                'aliexpress_price' => $aliexpress ? ($aliexpress->price ?? 0) : 0,
                'aliexpress_l30'   => $aliexpress ? ($aliexpress->aliexpress_l30 ?? 0) : 0,
                'aliexpress_l60'   => $aliexpress ? ($aliexpress->aliexpress_l60 ?? 0) : 0,
                'aliexpress_pft'   => $aliexpress && ($aliexpress->price ?? 0) > 0 ? (($aliexpress->price * 0.89 - $lp - $ship) / $aliexpress->price) : 0,
                'aliexpress_roi'   => $aliexpress && $lp > 0 && ($aliexpress->price ?? 0) > 0 ? (($aliexpress->price * 0.89 - $lp - $ship) / $lp) : 0,
                'aliexpress_buyer_link' => isset($aliexpressListingData[$sku]) ? ($aliexpressListingData[$sku]->value['buyer_link'] ?? null) : null,
                'aliexpress_seller_link' => isset($aliexpressListingData[$sku]) ? ($aliexpressListingData[$sku]->value['seller_link'] ?? null) : null,

                // Direct assignments for blade template
                'views_clicks' => $shein ? ($shein->views_clicks ?? 0) : 0,
                'lmp' => $shein ? ($shein->lmp ?? 0) : 0,
                'shopify_sheinl30' => $shein ? ($shein->shopify_sheinl30 ?? 0) : 0,

                'total_req_view' => (
                    ($ebay && $ebay->views && $ebay->ebay_l30 ? ($inv * 20) : 0) +
                    ($ebay2 && $ebay2->views && $ebay2->ebay_l30 ? ($inv * 20) : 0) +
                    ($ebay3 && $ebay3->views && $ebay3->ebay_l30 ? ($inv * 20) : 0) +
                    ($amazon && $amazon->sessions_l30 && $amazon->units_ordered_l30 ? ($inv * 20) : 0) +
                    ($shein && $shein->views_clicks && $shein->shopify_sheinl30 ? ($inv * 20) : 0) +
                    ($reverb && $reverb->views && $reverb->r_l30 ? ($inv * 20) : 0) +
                    ($temuMetric && ($temuMetric->{'product_clicks_l30'} ?? 0) && ($temuMetric->{'quantity_purchased_l30'} ?? 0) ? ($inv * 20) : 0)
                ),


                //  100 / cvr * inv not cvr percentage                             












                // Amazon DataView values
                'amz_sprice' => isset($amazonDataView[$sku]) ?
                    (is_array($amazonDataView[$sku]->value) ?
                        ($amazonDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($amazonDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'amz_spft' => isset($amazonDataView[$sku]) ?
                    (is_array($amazonDataView[$sku]->value) ?
                        ($amazonDataView[$sku]->value['SPFT'] ?? null) : (json_decode($amazonDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'amz_sroi' => isset($amazonDataView[$sku]) ?
                    (is_array($amazonDataView[$sku]->value) ?
                        ($amazonDataView[$sku]->value['SROI'] ?? null) : (json_decode($amazonDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'shopifyb2c_sprice' => isset($shopifyb2cDataView[$sku]) ?
                    (is_array($shopifyb2cDataView[$sku]->value) ?
                        ($shopifyb2cDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($shopifyb2cDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'shopifyb2c_spft' => isset($shopifyb2cDataView[$sku]) ?
                    (is_array($shopifyb2cDataView[$sku]->value) ?
                        ($shopifyb2cDataView[$sku]->value['SPFT'] ?? null) : (json_decode($shopifyb2cDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'shopifyb2c_sroi' => isset($shopifyb2cDataView[$sku]) ?
                    (is_array($shopifyb2cDataView[$sku]->value) ?
                        ($shopifyb2cDataView[$sku]->value['SROI'] ?? null) : (json_decode($shopifyb2cDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                // eBay DataView values
                'ebay_sprice' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'ebay_spft' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SPFT'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'ebay_sroi' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SROI'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'ebay2_sprice' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'ebay2_spft' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SPFT'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'ebay2_sroi' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SROI'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'ebay3_sprice' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'ebay3_spft' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SPFT'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'ebay3_sroi' => isset($ebayDataView[$sku]) ?
                    (is_array($ebayDataView[$sku]->value) ?
                        ($ebayDataView[$sku]->value['SROI'] ?? null) : (json_decode($ebayDataView[$sku]->value, true)['SROI'] ?? null)) : null,


                'doba_sprice' => isset($dobaDataView[$sku]) ?
                    (is_array($dobaDataView[$sku]->value) ?
                        ($dobaDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($dobaDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'doba_final_price' => isset($dobaDataView[$sku]) ?
                    (is_array($dobaDataView[$sku]->value) ?
                        ($dobaDataView[$sku]->value['FINAL_PRICE'] ?? null) : (json_decode($dobaDataView[$sku]->value, true)['FINAL_PRICE'] ?? null)) : null,
                'doba_spft' => isset($dobaDataView[$sku]) ? (is_array($dobaDataView[$sku]->value) ?
                    ($dobaDataView[$sku]->value['SPFT'] ?? null) : (json_decode($dobaDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'doba_sroi' => isset($dobaDataView[$sku]) ?
                    (is_array($dobaDataView[$sku]->value) ?
                        ($dobaDataView[$sku]->value['SROI'] ?? null) : (json_decode($dobaDataView[$sku]->value, true)['SROI'] ?? null)) : null,


                'temu_sprice' => isset($temuDataView[$sku]) ?
                    (is_array($temuDataView[$sku]->value) ?
                        ($temuDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($temuDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'temu_spft' => isset($temuDataView[$sku]) ? (is_array($temuDataView[$sku]->value) ?
                    ($temuDataView[$sku]->value['SPFT'] ?? null) : (json_decode($temuDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'temu_sroi' => isset($temuDataView[$sku]) ?
                    (is_array($temuDataView[$sku]->value) ?
                        ($temuDataView[$sku]->value['SROI'] ?? null) : (json_decode($temuDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'reverb_sprice' => isset($reverbDataView[$sku]) ?
                    (is_array($reverbDataView[$sku]->value) ?
                        ($reverbDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($reverbDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'reverb_spft' => isset($reverbDataView[$sku]) ? (is_array($reverbDataView[$sku]->value) ?
                    ($reverbDataView[$sku]->value['SPFT'] ?? null) : (json_decode($reverbDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'reverb_sroi' => isset($reverbDataView[$sku]) ? (is_array($reverbDataView[$sku]->value) ?
                    ($reverbDataView[$sku]->value['SROI'] ?? null) : (json_decode($reverbDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'macy_sprice' => isset($macyDataView[$sku]) ?
                    (is_array($macyDataView[$sku]->value) ?
                        ($macyDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($macyDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'macy_spft' => isset($macyDataView[$sku]) ? (is_array($macyDataView[$sku]->value) ?
                    ($macyDataView[$sku]->value['SPFT'] ?? null) : (json_decode($macyDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'macy_sroi' => isset($macyDataView[$sku]) ?
                    (is_array($macyDataView[$sku]->value) ?
                        ($macyDataView[$sku]->value['SROI'] ?? null) : (json_decode($macyDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'shein_sprice' => isset($sheinDataView[$sku]) ?
                    (is_array($sheinDataView[$sku]->value) ?
                        ($sheinDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($sheinDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'shein_spft' => isset($sheinDataView[$sku]) ? (is_array($sheinDataView[$sku]->value) ?
                    ($sheinDataView[$sku]->value['SPFT'] ?? null) : (json_decode($sheinDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'shein_sroi' => isset($sheinDataView[$sku]) ?
                    (is_array($sheinDataView[$sku]->value) ?
                        ($sheinDataView[$sku]->value['SROI'] ?? null) : (json_decode($sheinDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'bestbuy_sprice' => isset($bestbuyUsaDataView[$sku]) ?
                    (is_array($bestbuyUsaDataView[$sku]->value) ?
                        ($bestbuyUsaDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($bestbuyUsaDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'bestbuy_spft' => isset($bestbuyUsaDataView[$sku]) ? (is_array($bestbuyUsaDataView[$sku]->value) ?
                    ($bestbuyUsaDataView[$sku]->value['SPFT'] ?? null) : (json_decode($bestbuyUsaDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'bestbuy_sroi' => isset($bestbuyUsaDataView[$sku]) ?
                    (is_array($bestbuyUsaDataView[$sku]->value) ?
                        ($bestbuyUsaDataView[$sku]->value['SROI'] ?? null) : (json_decode($bestbuyUsaDataView[$sku]->value, true)['SROI'] ?? null)) : null,

                'tiendamia_sprice' => isset($tiendamiaDataView[$sku]) ?
                    (is_array($tiendamiaDataView[$sku]->value) ?
                        ($tiendamiaDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($tiendamiaDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'tiendamia_spft' => isset($tiendamiaDataView[$sku]) ? (is_array($tiendamiaDataView[$sku]->value) ?
                    ($tiendamiaDataView[$sku]->value['SPFT'] ?? null) : (json_decode($tiendamiaDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'tiendamia_sroi' => isset($tiendamiaDataView[$sku]) ?
                    (is_array($tiendamiaDataView[$sku]->value) ?
                        ($tiendamiaDataView[$sku]->value['SROI'] ?? null) : (json_decode($tiendamiaDataView[$sku]->value, true)['SROI'] ?? null)) : null,




                'tiktok_sprice' => isset($tiktokDataView[$sku]) ?
                    (is_array($tiktokDataView[$sku]->value) ?
                        ($tiktokDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($tiktokDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'tiktok_spft' => isset($tiktokDataView[$sku]) ? (is_array($tiktokDataView[$sku]->value) ?
                    ($tiktokDataView[$sku]->value['SPFT'] ?? null) : (json_decode($tiktokDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'tiktok_sroi' => isset($tiktokDataView[$sku]) ?
                    (is_array($tiktokDataView[$sku]->value) ?
                        ($tiktokDataView[$sku]->value['SROI'] ?? null) : (json_decode($tiktokDataView[$sku]->value, true)['SROI'] ?? null)) : null,



                'aliexpress_sprice' => isset($aliexpressDataView[$sku]) ?
                    (is_array($aliexpressDataView[$sku]->value) ?
                        ($aliexpressDataView[$sku]->value['SPRICE'] ?? null) : (json_decode($aliexpressDataView[$sku]->value, true)['SPRICE'] ?? null)) : null,
                'aliexpress_spft' => isset($aliexpressDataView[$sku]) ? (is_array($aliexpressDataView[$sku]->value) ?
                    ($aliexpressDataView[$sku]->value['SPFT'] ?? null) : (json_decode($aliexpressDataView[$sku]->value, true)['SPFT'] ?? null)) : null,
                'aliexpress_sroi' => isset($aliexpressDataView[$sku]) ?
                    (is_array($aliexpressDataView[$sku]->value) ?
                        ($aliexpressDataView[$sku]->value['SROI'] ?? null) : (json_decode($aliexpressDataView[$sku]->value, true)['SROI'] ?? null)) : null,



            ];



            // Set avg_inventory after $item is created calcution 
            $item->avg_inventory = $lp != 0 ? (($item->initial_cogs + $item->current_cogs) / 2) : 0;
            $item->initial_calculated_cogs = $item->initial_cogs - $item->current_cogs;
            $item->inventory_turnover_ratio = $item->initial_calculated_cogs != 0 ? ($item->initial_calculated_cogs / $item->avg_inventory) : 0;
            $item->stock_rotation_days = $item->inventory_turnover_ratio != 0 ? 365 / $item->inventory_turnover_ratio : 0;



            // Add shopifyb2c fields after $item is created
            $shopify = $shopifyData[trim(strtoupper($sku))] ?? null;
            $item->shopifyb2c_price = $shopify ? $shopify->price : 0;
            $item->shopifyb2c_l30 = $shopify ? $shopify->quantity : 0;
            $item->shopifyb2c_l30_data = $shopify ? $shopify->shopify_l30 : 0;
            $item->shopifyb2c_image = $shopify ? $shopify->image_src : null;
            $item->shopifyb2c_pft = $item->shopifyb2c_price > 0 ? (($item->shopifyb2c_price * 0.75 - $lp - $ship) / $item->shopifyb2c_price) : 0;
            $item->shopifyb2c_roi = ($lp > 0 && $item->shopifyb2c_price > 0) ? (($item->shopifyb2c_price * 0.75 - $lp - $ship) / $lp) : 0;

            // Add inv_value and COGS calculations
            // inv_value = INV × shopifyb2c_price (total inventory worth at selling price)
            $item->inv_value = $inv * $item->shopifyb2c_price;

            // COGS = LP × INV (cost of goods sold - total cost of current inventory)
            $item->COGS = $lp * $inv;

            // Add analysis action buttons
            $item->l30_analysis = '<button class="btn btn-sm btn-info" onclick="showL30Modal(this)" data-sku="' . $item->SKU . '">L30</button>';


            $processedData[] = $item;
        }

        return $processedData;
    }

    public function getViewPricingAnalysisData(Request $request)
    {
        $page = $request->input('page', 1);
        $perPage = $request->input('per_page', 'all');
        $dilFilter = $request->input('dil_filter', 'all');
        $dataType = $request->input('data_type', 'all');
        $searchTerm = $request->input('search', '');
        $parentFilter = $request->input('parent', '');
        $skuFilter = $request->input('sku', '');
        $distinctOnly = $request->input('distinct_only', false);

        if ($perPage === 'all') {
            $perPage = 1000000;
        } else {
            $perPage = (int) $perPage;
        }

        $processedData = $this->processPricingData($searchTerm);

        $filteredData = $this->applyFilters($processedData, $dilFilter, $dataType, $parentFilter, $skuFilter);

        if ($distinctOnly) {
            return response()->json([
                'distinct_values' => $this->getDistinctValues($processedData),
                'status' => 200,
            ]);
        }

        $total = count($filteredData);
        $totalPages = ceil($total / $perPage);
        $offset = ($page - 1) * $perPage;
        $paginatedData = array_slice($filteredData, $offset, $perPage);


        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $paginatedData,
            'distinct_values' => $this->getDistinctValues($processedData),
            'current_page' => (int) $page,
            'per_page' => $perPage,
            'total' => $total,
            'last_page' => $totalPages,
            'status' => 200,
        ]);
    }


    // Get Pricing ROI Dashboard Data 


    protected function applyFilters($data, $dilFilter, $dataType, $parentFilter, $skuFilter)
    {
        return array_filter($data, function ($item) use ($dilFilter, $dataType, $parentFilter, $skuFilter) {
            if ($dilFilter !== 'all') {
                $dilPercent = ($item->{'Dil%'} ?? 0) * 100;
                switch ($dilFilter) {
                    case 'yellow':
                        if ($dilPercent >= 16.66) {
                            return false;
                        }
                        break;
                    case 'yellow':
                        if ($dilPercent < 0 || $dilPercent >= 0) {
                            return false;
                        }
                        break;
                    case 'green':
                        if ($dilPercent < 0 || $dilPercent >= 0) {
                            return false;
                        }
                        break;
                    case 'blue':
                        if ($dilPercent < 0) {
                            return false;
                        }
                        break;
                }
            }

            if ($dataType !== 'all') {
                $isParent = stripos($item->SKU ?? '', 'PARENT') !== false;
                if ($dataType === 'parent' && !$isParent) {
                    return false;
                }
                if ($dataType === 'sku' && $isParent) {
                    return false;
                }
            }

            if ($parentFilter && $item->Parent !== $parentFilter) {
                return false;
            }
            if ($skuFilter && $item->SKU !== $skuFilter) {
                return false;
            }

            return true;
        });
    }


    protected function calculateCVR($l30, $views)
    {
        if (!$views) return null;
        $cvr = ($l30 / $views) * 100;
        return [
            'value' => number_format($cvr, 2),
            'color' => $cvr <= 7 ? 'blue' : ($cvr <= 13 ? 'green' : 'red')
        ];
    }


    protected function getDistinctValues($data)
    {
        $parents = [];
        $skus = [];

        foreach ($data as $item) {
            if (!empty($item->Parent)) {
                $parents[$item->Parent] = true;
            }
            if (!empty($item->SKU)) {
                $skus[$item->SKU] = true;
            }
        }

        return [
            'parents' => array_keys($parents),
            'skus' => array_keys($skus),
        ];
    }


    public function updatePrice(Request $request)
    {
        $sku = $request["sku"];
        $price = $request["price"];

        $price = app(AmazonSpApiService::class)->updateAmazonPriceUS($sku, $price);

        return response()->json(['status' => 200, 'data' => $price]);
    }

    public function saveSprice(Request $request)
    {
        $data = $request->validate([
            'sku' => 'required|string',
            'type' => 'required|string',
            'sprice' => 'required|numeric',
            'LP' => 'required|numeric',    // cost price
            'SHIP' => 'required|numeric',
            'temu_ship' => 'required|numeric',  // Temu shipping cost
        ]);

        $sku = $data['sku'];
        $type = $data['type'];
        $sprice = $data['sprice'];
        $lp = $data['LP'];
        $ship = $data['SHIP'];
        $temuship = $data['temu_ship'];

        switch ($type) {
            case 'shein':
                // Shein logic
                $sheinDataView = SheinDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($sheinDataView->value ?? null) ? $sheinDataView->value : (isset($sheinDataView->value) ? (json_decode($sheinDataView->value, true) ?: []) : []);

                $spft = $sprice > 0 ? round(((($sprice * 0.89) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.89) - $lp - $ship) / $lp) * 100 : 0;

                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $sheinDataView->value = json_encode($existing);
                $sheinDataView->save();
                break;
            case 'amz':
                // Amazon logic
                $amazonDataView = AmazonDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($amazonDataView->value) ? $amazonDataView->value : (json_decode($amazonDataView->value, true) ?: []);

                $spft = $sprice > 0 ? ((($sprice * 0.70) - $lp - $ship) / $sprice) * 100 : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.70) - $lp - $ship) / $lp) * 100 : 0;
                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');


                $amazonDataView->value = $existing;
                $amazonDataView->save();
                break;

            case 'ebay':
                // eBay logic
                $ebayDataView = EbayDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($ebayDataView->value) ? $ebayDataView->value : (json_decode($ebayDataView->value, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.72) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.72) - $lp - $ship) / $lp) * 100 : 0;

                // Round and store as string
                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');



                $ebayDataView->value = $existing;
                $ebayDataView->save();
                break;


            case 'shopifyb2c':
                try {
                    $shopifyDataView = Shopifyb2cDataView::firstOrNew(['sku' => $sku]);
                    $existing = is_array($shopifyDataView->value) ? $shopifyDataView->value : (json_decode($shopifyDataView->value, true) ?: []);

                    // Calculate values
                    $spft = $sprice > 0 ? ((($sprice * 0.75) - $lp - $ship) / $sprice) * 100 : 0;
                    $sroi = $lp > 0 ? ((($sprice * 0.75) - $lp - $ship) / $lp) * 100 : 0;

                    // Format and store values
                    $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                    $existing['SPFT'] = number_format($spft, 2, '.', '');
                    $existing['SROI'] = number_format($sroi, 2, '.', '');

                    // Convert to JSON if needed
                    $shopifyDataView->value = json_encode($existing);

                    // Save with error logging
                    if (!$shopifyDataView->save()) {
                        Log::error("Failed to save ShopifyB2C data for SKU: $sku");
                        throw new \Exception("Save failed");
                    }

                    // Update Shopify price
                    $request = new Request();
                    $request->merge(['sku' => $sku, 'price' => $sprice]);
                    $this->pushShopifyPriceBySku($request);
                } catch (\Exception $e) {
                    Log::error("Error saving ShopifyB2C price: " . $e->getMessage());
                    return response()->json([
                        'message' => 'Error saving ShopifyB2C price',
                        'error' => $e->getMessage(),
                        'status' => 500
                    ]);
                }
                break;


            case 'ebay2':
                // eBay2 logic
                $ebay2DataView = EbayTwoDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($ebay2DataView->value) ? $ebay2DataView->value : (json_decode($ebay2DataView->value, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.80) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.80) - $lp - $ship) / $lp) * 100 : 0;

                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $ebay2DataView->value = $existing;
                $ebay2DataView->save();
                break;


            case 'ebay3':
                // eBay3 logic
                $ebay3DataView = EbayThreeDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($ebay3DataView->value) ? $ebay3DataView->value : (json_decode($ebay3DataView->value, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.71) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.71) - $lp - $ship) / $lp) * 100 : 0;

                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $ebay3DataView->value = $existing;
                $ebay3DataView->save();
                break;

            case 'doba':
                // Doba logic
                $dobaDataView = DobaDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($dobaDataView->value) ? $dobaDataView->value : (json_decode($dobaDataView->value, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.95) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.95) - $lp - $ship) / $lp) * 100 : 0;

                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['FINAL_PRICE'] = number_format($sprice * 0.75, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $dobaDataView->value = $existing;
                $dobaDataView->save();

                // Update ProductMaster Values field with doba_final_price
                $product = ProductMaster::where('sku', $sku)->first();
                if ($product) {
                    $values = is_string($product->Values) ? json_decode($product->Values, true) : $product->Values;
                    if (!is_array($values)) {
                        $values = [];
                    }
                    $values['doba_final_price'] = number_format($sprice * 0.75, 2, '.', '');
                    $product->Values = json_encode($values);
                    $product->save();
                }
                break;

            case 'temu':
                // Temu logic
                $temuDataView = TemuDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($temuDataView->value) ? $temuDataView->value : (json_decode($temuDataView->value, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.87) - $lp - $temuship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.87) - $lp - $temuship) / $lp) * 100 : 0;

                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $temuDataView->value = $existing;
                $temuDataView->save();
                break;


            case 'reverb':
                // Reverb logic
                $reverbDataView = ReverbViewData::firstOrNew(['sku' => $sku]);
                $existing = is_array($reverbDataView->values) ? $reverbDataView->values : (json_decode($reverbDataView->values, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.84) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.84) - $lp - $ship) / $lp) * 100 : 0;


                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $reverbDataView->values = $existing;
                $reverbDataView->save();
                break;


            case 'macy':
                // Macy logic
                $macyDataView = MacyDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($macyDataView->value) ? $macyDataView->value : (json_decode($macyDataView->value, true) ?: []);
                $spft = $sprice > 0 ? round(((($sprice * 0.76) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.76) - $lp - $ship) / $lp) * 100 : 0;


                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $macyDataView->value = $existing;
                $macyDataView->save();
                break;


            case 'walmart':
                // Walmart logic
                $walmartDataView = WalmartDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($walmartDataView->value) ? $walmartDataView->value : (json_decode($walmartDataView->value, true) ?: []);

                $spft = $sprice > 0 ? ((($sprice * 0.80) - $lp - $ship) / $sprice) * 100 : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.80) - $lp - $ship) / $lp) * 100 : 0;
                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');


                $walmartDataView->value = $existing;
                $walmartDataView->save();
                break;

            case 'top':
                // Save to all marketplaces
                Log::info('Saving top for SKU: ' . $sku . ' Price: ' . $sprice);
                $marketplaces = [
                    'shein' => 0.89,
                    'amz' => 0.70,
                    'ebay' => 0.72,
                    'shopifyb2c' => 0.75,
                    'ebay2' => 0.80,
                    'ebay3' => 0.71,
                    'doba' => 0.95,
                    'temu' => 0.87,
                    'reverb' => 0.84,
                    'macy' => 0.76,
                    'walmart' => 0.80,
                    'tiktok' => 0.64,
                    'aliexpress' => 0.89
                ];

                foreach ($marketplaces as $mp => $percent) {
                    $shipping = ($mp === 'temu') ? $temuship : $ship;
                    $spft = $sprice > 0 ? round(((($sprice * $percent) - $lp - $shipping) / $sprice) * 100, 2) : 0;
                    $sroi = $lp > 0 ? ((($sprice * $percent) - $lp - $shipping) / $lp) * 100 : 0;

                    switch ($mp) {
                        case 'shein':
                            $dataView = SheinDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value ?? null) ? $dataView->value : (isset($dataView->value) ? (json_decode($dataView->value, true) ?: []) : []);
                            break;
                        case 'amz':
                            $dataView = AmazonDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'ebay':
                            $dataView = EbayDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'shopifyb2c':
                            $dataView = Shopifyb2cDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'ebay2':
                            $dataView = EbayTwoDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'ebay3':
                            $dataView = EbayThreeDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'doba':
                            $dataView = DobaDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            $existing['FINAL_PRICE'] = number_format($sprice * 0.75, 2, '.', '');
                            break;
                        case 'temu':
                            $dataView = TemuDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'reverb':
                            $dataView = ReverbViewData::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->values) ? $dataView->values : (json_decode($dataView->values, true) ?: []);
                            break;
                        case 'macy':
                            $dataView = MacyDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'walmart':
                            $dataView = WalmartDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                        case 'tiktok':
                            // For TikTok, update TiktokSheet directly
                            $tiktokSheet = TiktokSheet::firstOrNew(['sku' => $sku]);
                            $tiktokSheet->price = $sprice;
                            $tiktokSheet->save();
                            continue 2; // Skip the rest of the processing for TikTok
                        case 'aliexpress':
                            $dataView = AliexpressDataView::firstOrNew(['sku' => $sku]);
                            $existing = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
                            break;
                    }

                    $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                    $existing['SPFT'] = number_format($spft, 2, '.', '');
                    $existing['SROI'] = number_format($sroi, 2, '.', '');

                    if ($mp === 'reverb') {
                        $dataView->values = $existing;
                    } elseif (in_array($mp, ['shein', 'shopifyb2c'])) {
                        $dataView->value = json_encode($existing);
                    } else {
                        $dataView->value = $existing;
                    }
                    $dataView->save();
                }

                // Update ProductMaster for doba final price
                $product = ProductMaster::where('sku', $sku)->first();
                if ($product) {
                    $values = is_string($product->Values) ? json_decode($product->Values, true) : $product->Values;
                    if (!is_array($values)) {
                        $values = [];
                    }
                    $values['doba_final_price'] = number_format($sprice * 0.75, 2, '.', '');
                    $product->Values = json_encode($values);
                    $product->save();
                }
                break;

            case 'aliexpress':
                // AliExpress logic
                $aliexpressDataView = AliexpressDataView::firstOrNew(['sku' => $sku]);
                $existing = is_array($aliexpressDataView->value) ? $aliexpressDataView->value : (json_decode($aliexpressDataView->value, true) ?: []);

                $spft = $sprice > 0 ? round(((($sprice * 0.89) - $lp - $ship) / $sprice) * 100, 2) : 0;
                $sroi = $lp > 0 ? ((($sprice * 0.89) - $lp - $ship) / $lp) * 100 : 0;

                $existing['SPRICE'] = number_format($sprice, 2, '.', '');
                $existing['SPFT'] = number_format($spft, 2, '.', '');
                $existing['SROI'] = number_format($sroi, 2, '.', '');

                $aliexpressDataView->value = $existing;
                $aliexpressDataView->save();
                break;

            default:
                return response()->json([
                    'message' => 'Unknown marketplace type',
                    'status' => 400
                ]);
        }

        return response()->json([
            'message' => "$type S Price, SPFT & SROI saved successfully",
            'data' => [
                'SPRICE' => $sprice,
                'SPFT' => $spft,
                'SROI' => $sroi
            ],
            'status' => 200
        ]);
    }


    public function pushShopifyPriceBySku(Request $request)
    {

        $sku = $request->input('sku');
        $price = $request->input('price');

        $variantId = ShopifySku::where('sku', $sku)->value('variant_id');

        if (!$variantId) {
            return response()->json([
                'status' => 'error',
                'message' => "Variant ID not found for SKU: {$sku}"
            ], 404);
        }

        $result = UpdatePriceApiController::updateShopifyVariantPrice($variantId, $price);

        if ($result['status'] === 'success') {
            return response()->json([
                'status' => 'success',
                'message' => "Price updated successfully for SKU {$sku}",
                'data' => $result['data']
            ], 200);
        } else {
            return response()->json([
                'status' => 'error',
                'message' => $result['message'] ?? 'Unknown error occurred',
                'code' => $result['code'] ?? 500,
            ], 500);
        }
    }



    public function pushEbayPriceBySku(Request $request)
    {
        $sku = $request->input('sku');
        $price = $request->input('price');

        if (!$sku || !$price) {
            return response()->json([
                'error' => 'SKU and price are required'
            ], 400);
        }

        $itemId = EbayMetric::where('sku', $sku)->value('item_id');

        if (!$itemId) {
            return response()->json([
                'error' => "eBay Item ID not found for SKU: {$sku}"
            ], 404);
        }

        try {
            // Use direct eBay API call for instant update
            $result = $this->ebay->reviseFixedPriceItem($itemId, $price);

            if ($result['success']) {
                return response()->json([
                    'success' => true,
                    'message' => "eBay price updated successfully for SKU: {$sku}",
                    'data' => $result
                ]);
            } else {
                return response()->json([
                    'success' => false,
                    'error' => $result['errors'] ?? 'Unknown error',
                    'data' => $result
                ], 400);
            }
        } catch (Exception $e) {
            return response()->json([
                'success' => false,
                'error' => 'eBay API Error: ' . $e->getMessage()
            ], 500);
        }
    }

    public function pushEbayTwoPriceBySku(Request $request)
    {
        $sku = $request->input('sku');
        $price = $request->input('price');

        if (!$sku || !$price) {
            return response()->json([
                'error' => 'SKU and price are required'
            ], 400);
        }

        $itemId = Ebay2Metric::where('sku', $sku)->value('item_id');

        if (!$itemId) {
            return response()->json([
                'error' => "eBay2 Item ID not found for SKU: {$sku}"
            ], 404);
        }

        try {
            // Use direct eBay API call for instant update
            $result = $this->ebay->reviseFixedPriceItem($itemId, $price);

            if ($result['success']) {
                return response()->json([
                    'success' => true,
                    'message' => "eBay2 price updated successfully for SKU: {$sku}",
                    'data' => $result
                ]);
            } else {
                return response()->json([
                    'success' => false,
                    'error' => $result['errors'] ?? 'Unknown error',
                    'data' => $result
                ], 400);
            }
        } catch (Exception $e) {
            return response()->json([
                'success' => false,
                'error' => 'eBay2 API Error: ' . $e->getMessage()
            ], 500);
        }
    }

    public function pushEbayThreePriceBySku(Request $request)
    {
        $sku = $request->input('sku');
        $price = $request->input('price');

        if (!$sku || !$price) {
            return response()->json([
                'error' => 'SKU and price are required'
            ], 400);
        }

        $itemId = Ebay3Metric::where('sku', $sku)->value('item_id');

        if (!$itemId) {
            return response()->json([
                'error' => "eBay3 Item ID not found for SKU: {$sku}"
            ], 404);
        }

        try {
            // Use direct eBay API call for instant update
            $result = $this->ebay->reviseFixedPriceItem($itemId, $price);

            if ($result['success']) {
                return response()->json([
                    'success' => true,
                    'message' => "eBay3 price updated successfully for SKU: {$sku}",
                    'data' => $result
                ]);
            } else {
                return response()->json([
                    'success' => false,
                    'error' => $result['errors'] ?? 'Unknown error',
                    'data' => $result
                ], 400);
            }
        } catch (Exception $e) {
            return response()->json([
                'success' => false,
                'error' => 'eBay3 API Error: ' . $e->getMessage()
            ], 500);
        }
    }


    public function pushPricewalmart(Request $request)
    {
        $sku = $request->input('sku');
        $price = $request->input('price');

        $itemId = DB::connection('apicentral')->table('walmart_api_data')->where('sku', $sku)->value('sku');
        $result = $this->walmart->updatePrice($itemId, $price);

        if (isset($result['errors'])) {
            return response()->json(['error' => $result['errors']], 400);
        }

        return response()->json(['success' => true, 'message' => 'Price update submitted']);
    }



    // Doba prices 

    public function pushdobaPriceBySku(Request $request)
    {

        $sku   = $request->input('sku');

        // Validate inputs
        if (!$sku) {

            return response()->json([
                'error' => 'SKU is required'
            ], 400);
        }

        // Get FINAL_PRICE from DobaDataView instead of request
        $dobaDataView = DobaDataView::where('sku', $sku)->first();
        if (!$dobaDataView) {
            return response()->json([
                'error' => "No Doba data found for SKU: {$sku}"
            ], 404);
        }

        $existing = is_array($dobaDataView->value) ? $dobaDataView->value : (json_decode($dobaDataView->value, true) ?: []);
        $price = $existing['FINAL_PRICE'] ?? null;

        if (!$price) {
            return response()->json([
                'error' => "FINAL_PRICE not found for SKU: {$sku}"
            ], 404);
        }

        // Find Doba Item ID from your DB
        $itemId = DobaMetric::where('sku', $sku)->value('item_id');
        if (!$itemId) {


            return response()->json([
                'error' => "Item not found for SKU: {$sku}"
            ], 404);
        }

        // Update price directly on Doba
        $result = $this->doba->updateItemPrice($itemId, $price);



        if (isset($result['errors'])) {


            return response()->json([
                'error' => $result['errors'],
                'debug' => $result['debug'] ?? null
            ], 400);
        }

        return response()->json([
            'success' => true,
            'message' => "Price update submitted for SKU: {$sku}, Item ID: {$itemId}, New Price: {$price}",
            'debug' => [
                'doba_response' => $result,
                'sku' => $sku,
                'item_id' => $itemId,
                'price' => $price,
                'timestamp' => now()
            ]
        ]);
    }

    public function debugDobaSignature(Request $request)
    {
        $timestamp = $request->input('timestamp');
        return response()->json($this->doba->debugSignature($timestamp));
    }

    /**
     * Advanced debug Doba API request
     */
    public function advancedDobaDebug(Request $request)
    {
        try {
            $sku = $request->input('sku', 'SP 12120 4OHMS');
            $price = $request->input('price', 32.00);

            $dobaService = new DobaApiService();
            $result = $dobaService->advancedDebugRequest($sku, $price);

            return response()->json([
                'success' => true,
                'sku' => $sku,
                'price' => $price,
                'debug_results' => $result
            ]);
        } catch (Exception $e) {
            return response()->json([
                'success' => false,
                'error' => $e->getMessage(),
                'trace' => $e->getTraceAsString()
            ], 500);
        }
    }







    public function saveRemark(Request $request)
    {
        $data = $request->validate([
            'sku' => 'required|string',
            'remark' => 'nullable|string',
        ]);

        $product = ProductMaster::where('sku', $data['sku'])->first();

        if (!$product) {
            return response()->json([
                'success' => false,
                'message' => 'Product not found',
            ], 404);
        }

        $product->remark = $data['remark'];
        $product->save();

        return response()->json([
            'success' => true,
            'message' => 'Remark saved successfully',
            'data' => [
                'sku' => $product->sku,
                'remark' => $product->remark,
            ]
        ]);
    }

    public function pricingMasterCopy(Request $request)
    {

        return view('pricing-master.pricing_master_copy', []);
    }
}
