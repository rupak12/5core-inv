<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\FbaTable;
use App\Models\FbaPrice;
use App\Models\FbaReportsMaster;
use App\Models\FbaMonthlySale;
use App\Models\FbaManualData;

class FbaDataController extends Controller
{
   private function getFbaData()
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

      $shopifyData = ShopifySku::whereIn('sku', $skus)
         ->get()
         ->keyBy(function ($item) {
            return trim(strtoupper($item->sku));
         });

      $skus = array_map(function ($sku) {
         return strtoupper(trim($sku));
      }, $skus);

      $fbaData = FbaTable::whereRaw("seller_sku LIKE '%FBA%' OR seller_sku LIKE '%fba%'")
         ->get()
         ->keyBy(function ($item) {
            $sku = $item->seller_sku;
            $base = preg_replace('/\s*FBA\s*/i', '', $sku);
            return strtoupper(trim($base));
         });

      $fbaPriceData = FbaPrice::whereRaw("seller_sku LIKE '%FBA%' OR seller_sku LIKE '%fba%'")
         ->get()
         ->keyBy(function ($item) {
            $sku = $item->seller_sku;
            $base = preg_replace('/\s*FBA\s*/i', '', $sku);
            return strtoupper(trim($base));
         });

      $fbaReportsData = FbaReportsMaster::whereRaw("seller_sku LIKE '%FBA%' OR seller_sku LIKE '%fba%'")
         ->get()
         ->keyBy(function ($item) {
            $sku = $item->seller_sku;
            $base = preg_replace('/\s*FBA\s*/i', '', $sku);
            return strtoupper(trim($base));
         });


         $fbaMonthlySales = FbaMonthlySale::whereRaw("seller_sku LIKE '%FBA%' OR seller_sku LIKE '%fba%'")
         ->get()
         ->keyBy(function ($item) {
            $sku = $item->seller_sku;
            $base = preg_replace('/\s*FBA\s*/i', '', $sku);
            return strtoupper(trim($base));
         });

      $fbaManualData = FbaManualData::all()->keyBy(function($item) {
         return strtoupper(trim($item->sku));
      });

      $matchedSkus = $fbaData->keys()->toArray();
      $unmatchedSkus = array_diff($skus, $matchedSkus);

      return compact('productData', 'shopifyData', 'fbaData', 'fbaPriceData', 'fbaReportsData', 'matchedSkus', 'unmatchedSkus', 'fbaMonthlySales', 'fbaManualData');
   }

   public function fbaPageView()
   {

      $data = $this->getFbaData();

      return view('fba.fba_views_data', $data);
   }

   public function fbaDataJson()
   {
      $data = $this->getFbaData();

      $fbaData = $data['fbaData'];
      $fbaPriceData = $data['fbaPriceData'];
      $fbaReportsData = $data['fbaReportsData'];
      $shopifyData = $data['shopifyData'];
      $fbaMonthlySales = $data['fbaMonthlySales'];
      $fbaManualData = $data['fbaManualData'];
      $productData = $data['productData']->keyBy(function ($p) {
         return strtoupper(trim($p->sku));
      });

      // Prepare table data with repeated parent name for all child SKUs
      $tableData = $fbaData->map(function ($fba, $sku) use ($fbaPriceData, $fbaReportsData, $shopifyData, $productData, $fbaMonthlySales, $fbaManualData) {
         $fbaPriceInfo = $fbaPriceData->get($sku);
         $fbaReportsInfo = $fbaReportsData->get($sku);
         $shopifyInfo = $shopifyData->get($sku);
         $product = $productData->get($sku);
         $monthlySales = $fbaMonthlySales->get($sku);
         $manual = $fbaManualData->get(strtoupper(trim($fba->seller_sku)));

         return [
            'Parent' => $product ? ($product->parent ?? '') : '',
            'SKU' => $sku,
            'FBA_SKU' => $fba->seller_sku,
            'FBA_Price' => $fbaPriceInfo ? round(($fbaPriceInfo->price ?? 0), 2) : 0,
            'l30_units' => $monthlySales ? ($monthlySales->l30_units ?? 0) : 0,
            'l60_units' => $monthlySales ? ($monthlySales->l60_units ?? 0) : 0,
            'FBA_Quantity' => $fba->quantity_available,
            'Current_Month_Views' => $fbaReportsInfo ? ($fbaReportsInfo->current_month_views ?? 0) : 0,
            'Fulfillment_Fee' => $fbaReportsInfo ? round(($fbaReportsInfo->fulfillment_fee ?? 0), 2) : 0,
            'ASIN' => $fba->asin,
            'Shopify_INV' => $shopifyInfo ? ($shopifyInfo->quantity ?? 0) : 0,
            'Barcode' => $manual ? ($manual->data['barcode'] ?? '') : '',
            'Dispatch_Date' => $manual ? ($manual->data['dispatch_date'] ?? '') : '',
            'Weight' => $manual ? ($manual->data['weight'] ?? 0) : 0,
            'Quantity_in_each_box' => $manual ? ($manual->data['quantity_in_each_box'] ?? 0) : 0,
            'Send_Cost' => $manual ? ($manual->data['send_cost'] ?? 0) : 0,
            'IN_Charges' => $manual ? ($manual->data['in_charges'] ?? 0) : 0,
            'Total_quantity_sent' => $manual ? ($manual->data['total_quantity_sent'] ?? 0) : 0,
            'Done' => $manual ? ($manual->data['done'] ?? false) : false,
            'FBA_Send' => $manual ? ($manual->data['fba_send'] ?? false) : false,
            'Warehouse_INV_Reduction' => $manual ? ($manual->data['warehouse_inv_reduction'] ?? false) : false,
            'Shipping_Amount' => $manual ? ($manual->data['shipping_amount'] ?? 0) : 0,
            'Inbound_Quantity' => $manual ? ($manual->data['inbound_quantity'] ?? 0) : 0,
            'FBA_Send' => $manual ? ($manual->data['fba_send'] ?? false) : false,
            'Dimensions' => $manual ? ($manual->data['Dimensions'] ?? 0) : 0,
            'Jan' => $monthlySales ? ($monthlySales->jan ?? 0) : 0,
            'Feb' => $monthlySales ? ($monthlySales->feb ?? 0) : 0,
            'Mar' => $monthlySales ? ($monthlySales->mar ?? 0) : 0,
            'Apr' => $monthlySales ? ($monthlySales->apr ?? 0) : 0,
            'May' => $monthlySales ? ($monthlySales->may ?? 0) : 0,
            'Jun' => $monthlySales ? ($monthlySales->jun ?? 0) : 0,
            'Jul' => $monthlySales ? ($monthlySales->jul ?? 0) : 0,
            'Aug' => $monthlySales ? ($monthlySales->aug ?? 0) : 0,
            'Sep' => $monthlySales ? ($monthlySales->sep ?? 0) : 0,
            'Oct' => $monthlySales ? ($monthlySales->oct ?? 0) : 0,
            'Nov' => $monthlySales ? ($monthlySales->nov ?? 0) : 0,
            'Dec' => $monthlySales ? ($monthlySales->dec ?? 0) : 0,
         ];
      })->values();

      // Group by Parent and process
      $grouped = collect($tableData)->groupBy('Parent');

      $finalData = $grouped->flatMap(function ($rows, $parentKey) {
         $children = $rows->filter(fn($item) => !isset($item['is_parent']) || !$item['is_parent']);

         if ($children->isEmpty()) {
            return $rows;
         }

         // Create parent row
         $parentRow = [
            'Parent' => $parentKey,
            'SKU' => $parentKey,
            'FBA_SKU' => '',
            'FBA_Price' => '',
            'l30_units' => $children->sum('l30_units'),
          
            'l60_units' => $children->sum('l60_units'),
      
            'FBA_Quantity' => $children->sum('FBA_Quantity'),
            'Current_Month_Views' => $children->sum('Current_Month_Views'),
            'Fulfillment_Fee' => round($children->sum('Fulfillment_Fee'), 2),
            'ASIN' => '',
            'Shopify_INV' => $children->sum('Shopify_INV'),
            'Barcode' => '',
            'Dispatch_Date' => '',
            'Weight' => $children->sum('Weight'),
            'Quantity_in_each_box' => $children->sum('Quantity_in_each_box'),
            'Total_quantity_sent' => $children->sum('Total_quantity_sent'),
            'Send_Cost' => $children->sum('Send_Cost'),
            'IN_Charges' => $children->sum('IN_Charges'),
            'Done' => false,
            'Warehouse_INV_Reduction' => false,
            'FBA_Send' => false,
            'Shipping_Amount' => $children->sum('Shipping_Amount'),
            'Inbound_Quantity' => $children->sum('Inbound_Quantity'),
            'Dimensions' => $children->sum('Dimensions'),
            'Jan' => $children->sum('Jan'),
            'Feb' => $children->sum('Feb'),
            'Mar' => $children->sum('Mar'),
            'Apr' => $children->sum('Apr'),
            'May' => $children->sum('May'),
            'Jun' => $children->sum('Jun'),
            'Jul' => $children->sum('Jul'),
            'Aug' => $children->sum('Aug'),
            'Sep' => $children->sum('Sep'),
            'Oct' => $children->sum('Oct'),
            'Nov' => $children->sum('Nov'),
            'Dec' => $children->sum('Dec'),
            'is_parent' => true
         ];

         // Return children first, then parent
         return $children->push($parentRow);
      })->values();

      return response()->json($finalData);
   }

   public function getFbaMonthlySales($sku)
   {
      $baseSku = strtoupper(trim($sku));

      $sales = FbaMonthlySale::whereRaw("seller_sku LIKE '%FBA%' OR seller_sku LIKE '%fba%'")
         ->get()
         ->filter(function ($item) use ($baseSku) {
            $sku = $item->seller_sku;
            $base = preg_replace('/\s*FBA\s*/i', '', $sku);
            return strtoupper(trim($base)) === $baseSku;
         })
         ->first();

      if (!$sales) {
         return response()->json(['error' => 'No data found'], 404);
      }

      $monthlyData = [
         'Jan' => $sales->jan ?? 0,
         'Feb' => $sales->feb ?? 0,
         'Mar' => $sales->mar ?? 0,
         'Apr' => $sales->apr ?? 0,
         'May' => $sales->may ?? 0,
         'Jun' => $sales->jun ?? 0,
         'Jul' => $sales->jul ?? 0,
         'Aug' => $sales->aug ?? 0,
         'Sep' => $sales->sep ?? 0,
         'Oct' => $sales->oct ?? 0,
         'Nov' => $sales->nov ?? 0,
         'Dec' => $sales->dec ?? 0,
      ];

      return response()->json([
         'sku' => $sku,
         'monthly_sales' => $monthlyData,
         'total_units' => $sales->total_units ?? 0,
         'avg_price' => $sales->avg_price ?? 0,
      ]);
   }

   public function updateFbaManualData(Request $request)
   {
      $sku = strtoupper(trim($request->input('sku')));
      $field = $request->input('field');
      $value = $request->input('value');

      $manual = FbaManualData::where('sku', $sku)->first();

      if (!$manual) {
         $manual = new FbaManualData();
         $manual->sku = $sku;
         $manual->data = [];
      }

      $data = $manual->data ?? [];
      $data[$field] = $value;
      $manual->data = $data;
      $manual->save();

      return response()->json(['success' => true]);
   }
}
