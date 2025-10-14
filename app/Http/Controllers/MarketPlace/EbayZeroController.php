<?php

namespace App\Http\Controllers\MarketPlace;

use App\Http\Controllers\ApiController;
use App\Http\Controllers\Controller;
use App\Models\AmazonDataView;
use App\Models\EbayDataView;
use App\Models\EbayListingStatus;
use App\Models\EbayMetric;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\MarketplacePercentage;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Cache;

class EbayZeroController extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }

    
    public function ebayZero(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        // Get percentage from cache or database
        $percentage = Cache::remember('ebay_marketplace_percentage', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'eBay')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100; // Default to 100 if not set
        });

        return view('market-places.ebayZeroView', [
            'mode' => $mode,
            'demo' => $demo,
            'amazonPercentage' => $percentage
        ]);
    }

    public function getVieweBayZeroData(Request $request)
    {
        // 1. Fetch all ProductMaster rows (base)
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        // Fetch data from the Google Sheet using the ApiController method
        // Prepare SKU list for related models
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        // Fetch related data
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayMetrics = EbayMetric::whereIn('sku', $skus)->get()->keyBy('sku');
        // Fetch all EbayDataView rows for these SKUs
        $ebayDataViews = EbayDataView::whereIn('sku', $skus)->get()->keyBy('sku');

        // Build the result using ProductMaster as the base
        $processedData = [];
        foreach ($productMasters as $pm) {
            $sku = $pm->sku;
            $parentSku = $pm->parent;
            $imagePath = null;

            // Try to get image from Shopify first
            $shopify = $shopifyData[$sku] ?? null;
            if ($shopify && !empty($shopify->image_src)) {
                $imagePath = $shopify->image_src;
            } else {
                $values = is_array($pm->Values) ? $pm->Values : (is_string($pm->Values) ? json_decode($pm->Values, true) : []);
                if (isset($values['image_path'])) {
                    $imagePath = $values['image_path'];
                } elseif (isset($pm->image_path)) {
                    $imagePath = $pm->image_path;
                }
            }
            $buyerLink = $item->{'B Link'} ?? null;
            $sellerLink = $item->{'S Link'} ?? null;
            // Create base item object
            $item = (object) [
                'Parent' => $parentSku,
                '(Child) sku' => $sku,
                'B Link' => null,
                'S Link' => null,
                'INV' => $shopify ? $shopify->inv : 0,
                'L30' => $shopify ? $shopify->quantity : 0,
                'eBay L30' => 0,
                'eBay L60' => 0,
                'eBay Price' => 0,
                'OV CLICKS L30' => 0,
                'views' => $ebayMetric->views ?? 0,
                'image' => $imagePath,
                'A_Z_Reason' => '',
                'A_Z_ActionRequired' => '',
                'A_Z_ActionTaken' => '',
                'NR' => 'REQ',
                'B Link' => (filter_var($buyerLink, FILTER_VALIDATE_URL)) ? $buyerLink : null,
                'S Link' => (filter_var($sellerLink, FILTER_VALIDATE_URL)) ? $sellerLink : null,
            ];

            // eBay metrics
            if ($ebayMetrics->has($sku)) {
                $ebayMetric = $ebayMetrics[$sku];
                $item->{'eBay L30'} = $ebayMetric->ebay_l30;
                $item->{'eBay L60'} = $ebayMetric->ebay_l60;
                $item->{'eBay Price'} = $ebayMetric->ebay_price;
                $item->{'views'} = $ebayMetric->views ?? 0;
                $inv = $shopify->inv ?? 0;
                $eBayL30 = $item->{'eBay L30'} ?? 0;
                $item->{'E Dil%'} = ($inv > 0) ? round($eBayL30 / $inv, 2) : 0;
            }

            // EbayDataView
            $dataView = $ebayDataViews->get($sku);
            $value = $dataView ? $dataView->value : [];
            $item->{'A_Z_Reason'} = $value['A_Z_Reason'] ?? '';
            $item->{'A_Z_ActionRequired'} = $value['A_Z_ActionRequired'] ?? '';
            $item->{'A_Z_ActionTaken'} = $value['A_Z_ActionTaken'] ?? '';
            $item->{'NR'} = $value['NR'] ?? 'REQ';

            $processedData[] = $item;
        }

        // Filter: Only show rows with 0 views
        $filteredResults = array_filter($processedData, function ($item) {
            $childSku = $item->{'(Child) sku'} ?? '';
            $inv = $item->INV ?? 0;
            $views = $item->views ?? 1;
            return stripos($childSku, 'PARENT') === false && $inv > 0 && intval($views) === 0;
        });

        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => array_values($filteredResults),
            'status' => 200
        ]);

    }

    public function updateReasonAction(Request $request)
    {
        $sku = $request->input('sku');
        $reason = $request->input('reason');
        $actionRequired = $request->input('action_required');
        $actionTaken = $request->input('action_taken');

        if (!$sku) {
            return response()->json([
                'status' => 400,
                'message' => 'SKU is required.'
            ], 400);
        }

        $row = EbayDataView::firstOrCreate(['sku' => $sku]);
        $value = $row->value ?? [];
        $value['A_Z_Reason'] = $reason;
        $value['A_Z_ActionRequired'] = $actionRequired;
        $value['A_Z_ActionTaken'] = $actionTaken;
        $row->value = $value;
        $row->save();

        return response()->json([
            'status' => 200,
            'message' => 'Reason and actions updated successfully.'
        ]);
    }

    public function getZeroViewCount()
    {
        // Use the same logic as getVieweBayZeroData: INV > 0, views == 0, not parent SKU
        $productMasters = ProductMaster::orderBy('parent', 'asc')
            ->orderByRaw("CASE WHEN sku LIKE 'PARENT %' THEN 1 ELSE 0 END")
            ->orderBy('sku', 'asc')
            ->get();

        $skus = $productMasters->pluck('sku')->unique()->toArray();
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayMetrics = EbayMetric::whereIn('sku', $skus)->get()->keyBy('sku');

        $zeroCount = 0;
        foreach ($productMasters as $pm) {
            $sku = $pm->sku;
            $isParent = stripos($sku, 'PARENT') !== false;
            $inv = $shopifyData[$sku]->inv ?? 0;
            $views = $ebayMetrics[$sku]->views ?? 0;
            if (!$isParent && $inv > 0 && intval($views) === 0) {
                $zeroCount++;
            }
        }
        return $zeroCount;
    }

    public function getNrReqCount()
    {
        $productMasters = ProductMaster::whereNull('deleted_at')->get();
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $statusData = EbayDataView::whereIn('sku', $skus)->get()->keyBy('sku');

        $reqCount = 0;
        $nrCount = 0;
        $listedCount = 0;
        $pendingCount = 0;

        foreach ($productMasters as $item) {
            $sku = trim($item->sku);
            $inv = $shopifyData[$sku]->inv ?? 0;
            $isParent = stripos($sku, 'PARENT') !== false;

            if ($isParent || floatval($inv) <= 0) continue;

            $status = $statusData[$sku]->value ?? null;
            if (is_string($status)) {
                $status = json_decode($status, true);
            }

            // NR/REQ logic
            $nrReq = $status['NR'] ?? (floatval($inv) > 0 ? 'REQ' : 'NR');
            if ($nrReq === 'REQ') {
                $reqCount++;
            } elseif ($nrReq === 'NR') {
                $nrCount++;
            }

            // Listed/Pending logic
            $listed = $status['listed'] ?? (floatval($inv) > 0 ? 'Pending' : 'Listed');
            if ($listed === 'Listed') {
                $listedCount++;
            } elseif ($listed === 'Pending') {
                $pendingCount++;
            }
        }

        return [
            'NR'  => $nrCount,
            'REQ' => $reqCount,
            'Listed' => $listedCount,
            'Pending' => $pendingCount,
        ];
    }

    public function getLivePendingAndZeroViewCounts()
    {
        $productMasters = ProductMaster::whereNull('deleted_at')->get();
        $skus = $productMasters->pluck('sku')->unique()->toArray();

        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayDataViews = EbayListingStatus::whereIn('sku', $skus)->get()->keyBy('sku');
        $ebayMetrics = EbayMetric::whereIn('sku', $skus)->get()->keyBy('sku'); 

        $listedCount = 0;
        $zeroInvOfListed = 0;
        $liveCount = 0;
        $zeroViewCount = 0;

        foreach ($productMasters as $item) {
            $sku = trim($item->sku);
            $inv = $shopifyData[$sku]->inv ?? 0;
            $isParent = stripos($sku, 'PARENT') !== false;
            if ($isParent) continue;

            $status = $ebayDataViews[$sku]->value ?? null;
            if (is_string($status)) {
                $status = json_decode($status, true);
            }
            $listed = $status['listed'] ?? (floatval($inv) > 0 ? 'Pending' : 'Listed');
            $live = $status['live'] ?? null;

            // Listed count 
            if ($listed === 'Listed') {
                $listedCount++;
                if (floatval($inv) <= 0) {
                    $zeroInvOfListed++;
                }
            }

            // Live count
            if ($live === 'Live') {
                $liveCount++;
            }

            // Zero view: INV > 0, views == 0 (from ebay_metric table), not parent SKU (NR ignored)
            $views = $ebayMetrics[$sku]->views ?? null;
            // if (floatval($inv) > 0 && $views !== null && intval($views) === 0) {
            //     $zeroViewCount++;
            // }
            if ($inv > 0) {
                if ($views === null) {
                } elseif (intval($views) === 0) {
                    $zeroViewCount++;
                }
            }
        }

        // live pending = listed - 0-inv of listed - live
        $livePending = $listedCount - $zeroInvOfListed - $liveCount;

        return [
            'live_pending' => $livePending,
            'zero_view' => $zeroViewCount,
        ];
    }
}