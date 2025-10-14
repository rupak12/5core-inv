<?php

namespace App\Http\Controllers\Marketplace;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use App\Http\Controllers\ApiController;
use App\Models\ChannelMaster;
use App\Models\MarketplacePercentage;
use App\Models\WalmartDataView;
use Illuminate\Support\Facades\Cache;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use App\Models\WalmartMetrics;
use App\Models\WalmartProductSheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class WalmartControllerMarket extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }

    public function overallWalmart(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        // Get percentage from cache or database
        // $percentage = Cache::remember('amazon_marketplace_percentage', now()->addDays(30), function () {
        //     $marketplaceData = MarketplacePercentage::where('marketplace', 'Amazon')->first();
        //     return $marketplaceData ? $marketplaceData->percentage : 100;
        // });

        $marketplaceData = ChannelMaster::where('channel', 'Walmart')->first();

        $percentage = $marketplaceData ? $marketplaceData->channel_percentage : 100;
        $adUpdates = $marketplaceData ? $marketplaceData->ad_updates : 0;

        return view('market-places.walmartAnalysis', [
            'mode' => $mode,
            'demo' => $demo,
            'percentage' => $percentage
        ]);
    }

    public function walmartPricingCVR(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        $percentage = Cache::remember('Walmart', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'Walmart')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100;
        });

        return view('market-places.walmartPricingCvr', [
            'mode' => $mode,
            'demo' => $demo,
            'percentage' => $percentage
        ]);
    }

    public function getViewWalmartData(Request $request)
    {
        // Get percentage from cache or database
        $percentage = Cache::remember('Walmart', now()->addDays(30), function () {
            $marketplaceData = MarketplacePercentage::where('marketplace', 'Walmart')->first();
            return $marketplaceData ? $marketplaceData->percentage : 100;
        });
        $percentageValue = $percentage / 100;

        // Fetch all product master records
        $productMasterRows = ProductMaster::all()->keyBy('sku');

        // Get all unique SKUs from product master
        $skus = $productMasterRows->pluck('sku')->toArray();

        // Fetch shopify data for these SKUs
        $shopifyData = ShopifySku::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch NR values for these SKUs from walmartDataView
        $walmartDataViews = WalmartDataView::whereIn('sku', $skus)->get()->keyBy('sku');

        // Fetch Walmart product sheet data
        $walmartMetrics = WalmartMetrics::whereIn('sku', $skus)->get()->keyBy('sku');

        $nrValues = [];
        $listedValues = [];
        $liveValues = [];

        foreach ($walmartDataViews as $sku => $dataView) {
            $value = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
            $nrValues[$sku] = $value['NR'] ?? false;
            $listedValues[$sku] = isset($value['Listed']) ? (int) $value['Listed'] : false;
            $liveValues[$sku] = isset($value['Live']) ? (int) $value['Live'] : false;
        }

        // Process data from product master and shopify tables
        $processedData = [];
        $slNo = 1;

        foreach ($productMasterRows as $productMaster) {
            $sku = $productMaster->sku;
            $isParent = stripos($sku, 'PARENT') !== false;

            // Initialize the data structure
            $processedItem = [
                'SL No.' => $slNo++,
                'Parent' => $productMaster->parent ?? null,
                'Sku' => $sku,
                'R&A' => false, // Default value, can be updated as needed
                'is_parent' => $isParent,
                'raw_data' => [
                    'parent' => $productMaster->parent,
                    'sku' => $sku,
                    'Values' => $productMaster->Values
                ]
            ];

            // Add values from product_master
            $values = $productMaster->Values ?: [];
            $processedItem['LP'] = $values['lp'] ?? 0;
            $processedItem['Ship'] = $values['ship'] ?? 0;
            $processedItem['COGS'] = $values['cogs'] ?? 0;

            // Add data from shopify_skus if available
            if (isset($shopifyData[$sku])) {
                $shopifyItem = $shopifyData[$sku];
                $processedItem['INV'] = $shopifyItem->inv ?? 0;
                $processedItem['L30'] = $shopifyItem->quantity ?? 0;
                // FIXED: Add A L30 and A Dil% from Shopify if available
                $processedItem['A L30'] = $shopifyItem->a_l30 ?? 0;
                $processedItem['A Dil%'] = $shopifyItem->a_dil ?? 0;
                $processedItem['Sess30'] = $shopifyItem->sess30 ?? 0;
                $processedItem['Tacos30'] = $shopifyItem->tacos30 ?? 0;
                $processedItem['SCVR'] = $shopifyItem->scvr ?? 0;
            } else {
                $processedItem['INV'] = 0;
                $processedItem['L30'] = 0;
                $processedItem['A L30'] = 0;
                $processedItem['A Dil%'] = 0;
                $processedItem['Sess30'] = 0;
                $processedItem['Tacos30'] = 0;
                $processedItem['SCVR'] = 0;
            }

            // Add data from walmart_product_sheets if available
            if (isset($walmartMetrics[$sku])) {
                $walmartMetric = $walmartMetrics[$sku];
                $processedItem['sheet_price'] = $walmartMetric->price ?? 0;
                $processedItem['sheet_pft'] = $walmartMetric->pft ?? 0;
                $processedItem['sheet_roi'] = $walmartMetric->roi ?? 0;
                $processedItem['sheet_l30'] = $walmartMetric->l30 ?? 0; // Walmart L30
                $processedItem['sheet_dil'] = $walmartMetric->dil ?? 0; // Walmart Dilution
                $processedItem['buy_link'] = $walmartMetric->buy_link ?? '';
            } else {
                $processedItem['sheet_price'] = 0;
                $processedItem['sheet_pft'] = 0;
                $processedItem['sheet_roi'] = 0;
                $processedItem['sheet_l30'] = 0;
                $processedItem['sheet_dil'] = 0;
                $processedItem['buy_link'] = '';
            }

            // Fetch NR value if available
            $processedItem['NR'] = $nrValues[$sku] ?? false;
            $processedItem['Listed'] = $listedValues[$sku] ?? false;
            $processedItem['Live'] = $liveValues[$sku] ?? false;

            // Default values for other fields
            $processedItem['price'] = $processedItem['sheet_price'] ?? 0;
            $processedItem['TOTAL PFT'] = 0;
            $processedItem['T Sales L30'] = $processedItem['sheet_l30'] ?? 0;
            $processedItem['percentage'] = $percentageValue;

            // Calculate profit and ROI percentages
            $price = floatval($processedItem['price']);
            $percentage = floatval($processedItem['percentage']);
            $lp = floatval($processedItem['LP']);
            $ship = floatval($processedItem['Ship']);

            if ($price > 0) {
                $pft_percentage = (($price * $percentage - $lp - $ship) / $price) * 100;
                $processedItem['PFT %'] = round($pft_percentage, 2);
            } else {
                $processedItem['PFT %'] = 0;
            }

            if ($lp > 0) {
                $roi_percentage = (($price * $percentage - $lp - $ship) / $lp) * 100;
                $processedItem['Roi'] = round($roi_percentage, 2);
            } else {
                $processedItem['Roi'] = 0;
            }

            $processedData[] = $processedItem;
        }

        return response()->json([
            'message' => 'Data fetched successfully',
            'data' => $processedData,
            'status' => 200
        ]);
    }

    public function updateAllWalmartSkus(Request $request)
    {
        try {
            $percent = $request->input('percent');

            if (!is_numeric($percent) || $percent < 0 || $percent > 100) {
                return response()->json([
                    'status' => 400,
                    'message' => 'Invalid percentage value. Must be between 0 and 100.'
                ], 400);
            }

            // Update database
            MarketplacePercentage::updateOrCreate(
                ['marketplace' => 'Walmart'],
                ['percentage' => $percent]
            );

            // Store in cache
            Cache::put('Walmart', $percent, now()->addDays(30));

            return response()->json([
                'status' => 200,
                'message' => 'Percentage updated successfully',
                'data' => [
                    'marketplace' => 'Wayfair',
                    'percentage' => $percent
                ]
            ]);
        } catch (\Exception $e) {
            return response()->json([
                'status' => 500,
                'message' => 'Error updating percentage',
                'error' => $e->getMessage()
            ], 500);
        }
    }

    // Save NR value for a SKU
    public function saveNrToDatabase(Request $request)
    {
        $sku = $request->input('sku');
        $nr = $request->input('nr');

        if (!$sku || $nr === null) {
            return response()->json(['error' => 'SKU and nr are required.'], 400);
        }

        $dataView = WalmartDataView::firstOrNew(['sku' => $sku]);
        $value = is_array($dataView->value) ? $dataView->value : (json_decode($dataView->value, true) ?: []);
        $value['NR'] = $nr;
        $dataView->value = $value;
        $dataView->save();

        return response()->json(['success' => true, 'data' => $dataView]);
    }

    public function updateListedLive(Request $request)
    {
        $request->validate([
            'sku'   => 'required|string',
            'field' => 'required|in:Listed,Live',
            'value' => 'required|boolean' // validate as boolean
        ]);

        // Find or create the product without overwriting existing value
        $product = WalmartDataView::firstOrCreate(
            ['sku' => $request->sku],
            ['value' => []]
        );

        // Decode current value (ensure it's an array)
        $currentValue = is_array($product->value)
            ? $product->value
            : (json_decode($product->value, true) ?? []);

        // Store as actual boolean
        $currentValue[$request->field] = filter_var($request->value, FILTER_VALIDATE_BOOLEAN);

        // Save back to DB
        $product->value = $currentValue;
        $product->save();

        return response()->json(['success' => true]);
    }

    public function importWalmartAnalytics(Request $request)
    {
        $request->validate([
            'excel_file' => 'required|file|mimes:xlsx,xls,csv'
        ]);

        try {
            $file = $request->file('excel_file');
            $spreadsheet = IOFactory::load($file->getPathName());
            $sheet = $spreadsheet->getActiveSheet();
            $rows = $sheet->toArray();

            // Clean headers
            $headers = array_map(function ($header) {
                return strtolower(trim(preg_replace('/[^a-zA-Z0-9_]/', '_', $header)));
            }, $rows[0]);

            unset($rows[0]);

            $allSkus = [];
            foreach ($rows as $row) {
                if (!empty($row[0])) {
                    $allSkus[] = $row[0];
                }
            }

            $existingSkus = ProductMaster::whereIn('sku', $allSkus)
                ->pluck('sku')
                ->toArray();

            $existingSkus = array_flip($existingSkus);

            $importCount = 0;
            foreach ($rows as $index => $row) {
                if (empty($row[0])) { // Check if SKU is empty
                    continue;
                }

                // Ensure row has same number of elements as headers
                $rowData = array_pad(array_slice($row, 0, count($headers)), count($headers), null);
                $data = array_combine($headers, $rowData);

                if (!isset($data['sku']) || empty($data['sku'])) {
                    continue;
                }

                // Only import SKUs that exist in product_masters (in-memory check)
                if (!isset($existingSkus[$data['sku']])) {
                    continue;
                }

                // Prepare values array
                $values = [];

                // Handle boolean fields
                if (isset($data['listed'])) {
                    $values['Listed'] = filter_var($data['listed'], FILTER_VALIDATE_BOOLEAN);
                }

                if (isset($data['live'])) {
                    $values['Live'] = filter_var($data['live'], FILTER_VALIDATE_BOOLEAN);
                }

                // Update or create record
                WalmartDataView::updateOrCreate(
                    ['sku' => $data['sku']],
                    ['value' => $values]
                );

                $importCount++;
            }

            return back()->with('success', "Successfully imported $importCount records!");
        } catch (\Exception $e) {
            return back()->with('error', 'Error importing file: ' . $e->getMessage());
        }
    }

    public function exportWalmartAnalytics()
    {
        $walmartData = WalmartDataView::all();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Header Row
        $headers = ['SKU', 'Listed', 'Live'];
        $sheet->fromArray($headers, NULL, 'A1');

        // Data Rows
        $rowIndex = 2;
        foreach ($walmartData as $data) {
            $values = is_array($data->value)
                ? $data->value
                : (json_decode($data->value, true) ?? []);

            $sheet->fromArray([
                $data->sku,
                isset($values['Listed']) ? ($values['Listed'] ? 'TRUE' : 'FALSE') : 'FALSE',
                isset($values['Live']) ? ($values['Live'] ? 'TRUE' : 'FALSE') : 'FALSE',
            ], NULL, 'A' . $rowIndex);

            $rowIndex++;
        }

        // Set column widths
        $sheet->getColumnDimension('A')->setWidth(20);
        $sheet->getColumnDimension('B')->setWidth(10);
        $sheet->getColumnDimension('C')->setWidth(10);

        // Output Download
        $fileName = 'Walmart_Analytics_Export_' . date('Y-m-d') . '.xlsx';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }

    public function downloadSample()
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Header Row
        $headers = ['SKU', 'Listed', 'Live'];
        $sheet->fromArray($headers, NULL, 'A1');

        // Sample Data
        $sampleData = [
            ['SKU001', 'TRUE', 'FALSE'],
            ['SKU002', 'FALSE', 'TRUE'],
            ['SKU003', 'TRUE', 'TRUE'],
        ];

        $sheet->fromArray($sampleData, NULL, 'A2');

        // Set column widths
        $sheet->getColumnDimension('A')->setWidth(20);
        $sheet->getColumnDimension('B')->setWidth(10);
        $sheet->getColumnDimension('C')->setWidth(10);

        // Output Download
        $fileName = 'Walmart_Analytics_Sample.xlsx';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');

        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');
        exit;
    }
}
