<?php

namespace App\Http\Controllers\Channels;

use App\Http\Controllers\Controller;
use App\Models\AccountHealthMaster;
use App\Models\AmazonDatasheet;
use App\Models\AtoZClaimsRate;
use App\Models\BestbuyUsaProduct;
use App\Models\ChannelMaster;
use App\Models\DobaMetric;
use App\Models\Ebay2Metric;
use App\Models\Ebay3Metric;
use App\Models\EbayMetric;
use App\Models\FullfillmentRate;
use App\Models\LateShipmentRate;
use App\Models\MacyProduct;
use App\Models\MarketplacePercentage;
use App\Models\NegativeSellerRate;
use App\Models\OdrRate;
use App\Models\OnTimeDeliveryRate;
use App\Models\PLSProduct;
use App\Models\ProductMaster;
use App\Models\RefundRate;
use App\Models\ReverbProduct;
use App\Models\TemuProductSheet;
use App\Models\TiendamiaProduct;
use App\Models\ValidTrackingRate;
use App\Models\VoilanceRate;
use App\Models\WaifairProductSheet;
use App\Models\WalmartMetrics;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Validator;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class AccountHealthMasterDashboardController extends Controller
{
    public function dashboard(Request $request)
    {
        $mode = $request->query('mode');
        $demo = $request->query('demo');

        return view('channels.account_health_master.dashboard', [
            'mode' => $mode,
            'demo' => $demo,
        ]);
    }

    public function getMasterChannelDataHealthDashboard(Request $request)
    {
        try {
            // Get channel filter if provided
            $channelFilter = $request->get('channel');

            $channelsQuery = ChannelMaster::where('status', 'Active')->orderBy('id', 'asc');

            // Apply channel filter if provided
            if ($channelFilter) {
                $channelsQuery->where('channel', 'LIKE', '%' . $channelFilter . '%');
            }

            $channels = $channelsQuery->get(['id', 'channel', 'sheet_link', 'type', 'w_ads', 'nr', 'update']);

            if ($channels->isEmpty()) {
                return response()->json([
                    'status' => 404,
                    'message' => 'No active channels found'
                ], 404);
            }

            $finalData = [];

            $controllerMap = [
                'amazon' => 'getAmazonChannelData',
                'ebay' => 'getEbayChannelData',
                'ebaytwo' => 'getEbaytwoChannelData',
                'ebaythree' => 'getEbaythreeChannelData',
                'macys' => 'getMacysChannelData',
                'tiendamia' => 'getTiendamiaChannelData',
                'bestbuyusa' => 'getBestbuyUsaChannelData',
                'reverb' => 'getReverbChannelData',
                'doba' => 'getDobaChannelData',
                'temu' => 'getTemuChannelData',
                'walmart' => 'getWalmartChannelData',
                'pls' => 'getPlsChannelData',
                'wayfair' => 'getWayfairChannelData',
            ];

            // Get channel IDs for batch queries
            $channelIds = $channels->pluck('id')->toArray();

            // Define rate models with corrected names
            $rateModels = [
                'ODR Rate' => OdrRate::class,
                'Fulfillment Rate' => FullfillmentRate::class,
                'Valid Tracking Rate' => ValidTrackingRate::class,
                'Late Shipment Rate' => LateShipmentRate::class,
                'On Time Delivery Rate' => OnTimeDeliveryRate::class,
                'Negative Seller Rate' => NegativeSellerRate::class,
                'AtoZ Claims Rate' => AtoZClaimsRate::class,
                'Violation Rate' => VoilanceRate::class,
                'Refund Rate' => RefundRate::class,
            ];

            // Batch fetch all rate data for filtered channels
            $ratesData = [];
            foreach ($rateModels as $rateKey => $model) {
                $rates = $model::whereIn('channel_id', $channelIds)
                    ->orderBy('report_date', 'desc')
                    ->get(['channel_id', 'current', 'allowed'])
                    ->keyBy('channel_id');
                $ratesData[$rateKey] = $rates;
            }

            foreach ($channels as $channelRow) {
                $channel = $channelRow->channel;
                $channelId = $channelRow->id;
                $nowDate = now()->toDateString();

                // Initialize row with default values
                $row = [
                    'channel' => ucfirst($channel),
                    'sheet_link' => $channelRow->sheet_link,
                    'l30_orders' => 0,
                    'l60_orders' => 0,
                    'l30_sales' => 0,
                    'l60_sales' => 0,
                    'growth' => 0,
                    'gprofit%' => 'N/A',
                    'gprofit_l60' => 'N/A',
                    'g_roi%' => 'N/A',
                    'g_roi_l60' => 'N/A',
                    'red_margin' => 0,
                    'nr' => $channelRow->nr ?? 0,
                    'type' => $channelRow->type ?? '',
                    'listed_count' => 0,
                    'w_ads' => $channelRow->w_ads ?? 0,
                    'update' => $channelRow->update ?? 0,
                    'account_health' => null,
                ];

                // Get channel-specific data
                $key = strtolower(str_replace([' ', '-', '&', '/'], '', trim($channel)));
                if (isset($controllerMap[$key]) && method_exists($this, $controllerMap[$key])) {
                    try {
                        $method = $controllerMap[$key];
                        $data = $this->$method($request)->getData(true);
                        if (!empty($data['data'])) {
                            $channelData = $data['data'][0];
                            $row['l60_sales'] = $channelData['l60_sales'] ?? 0;
                            $row['l30_sales'] = $channelData['l30_sales'] ?? 0;
                            $row['growth'] = $channelData['growth'] ?? 0;
                            $row['l60_orders'] = $channelData['l60_orders'] ?? 0;
                            $row['l30_orders'] = $channelData['l30_orders'] ?? 0;
                            $row['gprofit%'] = $channelData['gprofit%'] ?? 'N/A';
                            $row['gprofit_l60'] = $channelData['gprofit_l60'] ?? 'N/A';
                            $row['g_roi%'] = $channelData['g_roi%'] ?? 'N/A';
                            $row['g_roi_l60'] = $channelData['g_roi_l60'] ?? 'N/A';
                            $row['red_margin'] = $channelData['red_margin'] ?? 0;
                            $row['nr'] = $channelData['nr'] ?? 0;
                            $row['type'] = $channelData['type'] ?? '';
                            $row['listed_count'] = $channelData['listed_count'] ?? 0;
                            $row['w_ads'] = $channelData['w_ads'] ?? 0;
                            $row['update'] = $channelData['update'] ?? 0;
                            $row['account_health'] = $channelData['account_health'] ?? null;
                        }
                    } catch (\Exception $e) {
                        Log::error("Error fetching data for channel {$channel}: " . $e->getMessage());
                    }
                }

                // Fetch rate data
                foreach ($rateModels as $rateKey => $model) {
                    $rate = $ratesData[$rateKey][$channelId] ?? null;
                    $fieldKey = strtolower(str_replace(' ', '_', $rateKey));
                    if ($rate) {
                        $row[$fieldKey] = $rate->current ?? 'N/A';
                        $row["{$fieldKey}_allowed"] = $rate->allowed ?? '';
                    } else {
                        try {
                            $existingRate = $model::where('channel_id', $channelId)->first();
                            if (!$existingRate) {
                                $model::create([
                                    'channel_id' => $channelId,
                                    'report_date' => $nowDate,
                                    'current' => null, // Changed from 'N/A' to null
                                    'allowed' => '',
                                    'what' => '',
                                    'why' => '',
                                    'action' => '',
                                    'c_action' => '',
                                    'account_health_links' => '',
                                ]);
                            }
                            $row[$fieldKey] = 'N/A';
                            $row["{$fieldKey}_allowed"] = '';
                        } catch (\Exception $e) {
                            Log::error("Error creating default rate for channel {$channel}, rate {$rateKey}: " . $e->getMessage());
                            $row[$fieldKey] = 'N/A';
                            $row["{$fieldKey}_allowed"] = '';
                        }
                    }
                }

                $finalData[] = $row;
            }

            return response()->json([
                'status' => 200,
                'message' => 'Channel data fetched successfully',
                'data' => $finalData,
            ]);
        } catch (\Exception $e) {
            Log::error("Error in getMasterChannelDataHealthDashboard: " . $e->getMessage());
            return response()->json([
                'status' => 500,
                'message' => 'Server error while fetching channel data',
            ], 500);
        }
    }

    private function fetchChannelData($model, $channelName, $l30Field, $l60Field, $priceField, $marketplace, $skuField = 'sku')
    {
        try {
            $query = $model::where($skuField, 'not like', '%Parent%');

            $l30Orders = $query->sum($l30Field);
            $l60Orders = $query->sum($l60Field);

            $l30Sales = (clone $query)->selectRaw("SUM({$l30Field} * {$priceField}) as total")->value('total') ?? 0;
            $l60Sales = (clone $query)->selectRaw("SUM({$l60Field} * {$priceField}) as total")->value('total') ?? 0;

            $growth = $l30Sales > 0 ? (($l30Sales - $l60Sales) / $l30Sales) * 100 : 0;

            // Get marketplace percentage
            $percentage = MarketplacePercentage::where('marketplace', $marketplace)->value('percentage') ?? 100;
            $percentage = $percentage / 100;

            // Fetch only relevant product masters
            $skus = $query->pluck($skuField)->map(fn($s) => strtoupper($s))->toArray();
            $productMasters = ProductMaster::whereIn($skuField, $skus)->get()->keyBy(function ($item) {
                return strtoupper($item->sku);
            });

            // Calculate total profit
            $rows = $query->get([$skuField, $priceField, $l30Field, $l60Field]);
            $totalProfit = 0;
            $totalProfitL60 = 0;
            $totalLpValue = 0;

            foreach ($rows as $row) {
                $sku = strtoupper($row->$skuField);
                $price = (float) $row->$priceField;
                $unitsL30 = (int) $row->$l30Field;
                $unitsL60 = (int) $row->$l60Field;

                $soldAmount = $unitsL30 * $price;
                if ($soldAmount <= 0) {
                    continue;
                }

                $lp = 0;
                $ship = 0;

                if (isset($productMasters[$sku])) {
                    $pm = $productMasters[$sku];
                    $values = is_array($pm->Values) ? $pm->Values : (is_string($pm->Values) ? json_decode($pm->Values, true) : []);
                    $lp = isset($values['lp']) ? (float) $values['lp'] : ($pm->lp ?? 0);
                    $ship = isset($values['ship']) ? (float) $values['ship'] : ($pm->ship ?? 0);
                    $totalLpValue += $lp;
                }

                $profitPerUnit = ($price * $percentage) - $lp - $ship;
                $totalProfit += $profitPerUnit * $unitsL30;
                $totalProfitL60 += $profitPerUnit * $unitsL60;
            }

            $gProfitPct = $l30Sales > 0 ? ($totalProfit / $l30Sales) * 100 : 0;
            $gProfitL60 = $l60Sales > 0 ? ($totalProfitL60 / $l60Sales) * 100 : 0;
            $gRoi = $totalLpValue > 0 ? ($totalProfit / $totalLpValue) : 0;
            $gRoiL60 = $totalLpValue > 0 ? ($totalProfitL60 / $totalLpValue) : 0;

            // Channel data
            $channelData = ChannelMaster::where('channel', $channelName)->first();

            return [
                'channel' => $channelName,
                'l60_sales' => intval($l60Sales),
                'l30_sales' => intval($l30Sales),
                'growth' => round($growth, 2) . '%',
                'l60_orders' => $l60Orders,
                'l30_orders' => $l30Orders,
                'gprofit%' => round($gProfitPct, 2) . '%',
                'gprofit_l60' => round($gProfitL60, 2) . '%',
                'g_roi%' => round($gRoi, 2),
                'g_roi_l60' => round($gRoiL60, 2),
                'type' => $channelData->type ?? '',
                'w_ads' => $channelData->w_ads ?? 0,
                'nr' => $channelData->nr ?? 0,
                'update' => $channelData->update ?? 0,
                'account_health' => null,
            ];
        } catch (\Exception $e) {
            Log::error("Error fetching data for {$marketplace} channel: " . $e->getMessage());
            return [
                'channel' => $channelName,
                'l60_sales' => 0,
                'l30_sales' => 0,
                'growth' => 0,
                'l60_orders' => 0,
                'l30_orders' => 0,
                'gprofit%' => 'N/A',
                'gprofit_l60' => 'N/A',
                'g_roi%' => 'N/A',
                'g_roi_l60' => 'N/A',
                'type' => '',
                'w_ads' => 0,
                'nr' => 0,
                'update' => 0,
                'account_health' => null,
            ];
        }
    }

    public function getAmazonChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            AmazonDatasheet::class,
            'Amazon',
            'units_ordered_l30',
            'units_ordered_l60',
            'price',
            'Amazon'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Amazon channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getEbayChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            EbayMetric::class,
            'eBay',
            'ebay_l30',
            'ebay_l60',
            'ebay_price',
            'Ebay'
        );
        return response()->json([
            'status' => 200,
            'message' => 'eBay channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getEbaytwoChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            Ebay2Metric::class,
            'EbayTwo',
            'ebay_l30',
            'ebay_l60',
            'ebay_price',
            'EbayTwo'
        );
        return response()->json([
            'status' => 200,
            'message' => 'eBay2 channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getEbaythreeChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            Ebay3Metric::class,
            'EbayThree',
            'ebay_l30',
            'ebay_l60',
            'ebay_price',
            'EbayThree'
        );
        return response()->json([
            'status' => 200,
            'message' => 'eBay three channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getMacysChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            MacyProduct::class,
            'Macys',
            'm_l30',
            'm_l60',
            'price',
            'Macys'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Macys channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getTiendamiaChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            TiendamiaProduct::class,
            'Tiendamia',
            'm_l30',
            'm_l60',
            'price',
            'Tiendamia'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Tiendamia channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getBestbuyUsaChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            BestbuyUsaProduct::class,
            'BestBuy USA',
            'm_l30',
            'm_l60',
            'price',
            'BestbuyUSA'
        );
        return response()->json([
            'status' => 200,
            'message' => 'BestBuy USA channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getReverbChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            ReverbProduct::class,
            'Reverb',
            'r_l30',
            'r_l60',
            'price',
            'Reverb'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Reverb channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getDobaChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            DobaMetric::class,
            'Doba',
            'quantity_l30',
            'quantity_l60',
            'anticipated_income',
            'Doba'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Doba channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getTemuChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            TemuProductSheet::class,
            'Temu',
            'l30',
            'l60',
            'price',
            'Temu'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Temu channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getWalmartChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            WalmartMetrics::class,
            'Walmart',
            'l30',
            'l60',
            'price',
            'Walmart'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Walmart channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getPlsChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            PLSProduct::class,
            'PLS',
            'p_l30',
            'p_l60',
            'price',
            'Pls'
        );
        return response()->json([
            'status' => 200,
            'message' => 'PLS channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    public function getWayfairChannelData(Request $request)
    {
        $data = $this->fetchChannelData(
            WaifairProductSheet::class,
            'Wayfair',
            'l30',
            'l60',
            'price',
            'Wayfair'
        );
        return response()->json([
            'status' => 200,
            'message' => 'Wayfair channel data fetched successfully',
            'data' => [$data],
        ]);
    }

    /**
     * Export account health data to Excel
     */
    public function export(Request $request)
    {
        try {
            // Get all channel data
            $response = $this->getMasterChannelDataHealthDashboard($request);
            $data = $response->getData(true);

            if ($data['status'] !== 200) {
                return redirect()->back()->with('error', 'Failed to fetch data for export');
            }

            $channelData = $data['data'];

            // Create spreadsheet
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setTitle('Account Health Data');

            // Headers
            $headers = [
                'A1' => 'Channel',
                'B1' => 'L30 Sales',
                'C1' => 'L30 Orders',
                'D1' => 'L60 Orders',
                'E1' => 'Growth %',
                'F1' => 'Gross Profit %',
                'G1' => 'G ROI %',
                'H1' => 'NR',
                'I1' => 'Type',
                'J1' => 'Listed Count',
                'K1' => 'ODR Rate',
                'L1' => 'Fulfillment Rate',
                'M1' => 'Valid Tracking Rate',
                'N1' => 'On Time Delivery Rate',
                'O1' => 'AtoZ Claims Rate',
                'P1' => 'Violation Rate',
                'Q1' => 'Refund Rate',
                'R1' => 'Sheet Link',
                'S1' => 'Account Health Links'
            ];

            // Set headers
            foreach ($headers as $cell => $value) {
                $sheet->setCellValue($cell, $value);
            }

            // Style headers
            $sheet->getStyle('A1:S1')->getFont()->setBold(true);
            $sheet->getStyle('A1:S1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
            $sheet->getStyle('A1:S1')->getFill()->getStartColor()->setARGB('CCCCCC');

            // Fill data
            $row = 2;
            foreach ($channelData as $channel) {
                $sheet->setCellValue('A' . $row, $channel['Channel '] ?? '');
                $sheet->setCellValue('B' . $row, $channel['L30 Sales'] ?? 0);
                $sheet->setCellValue('C' . $row, $channel['L30 Orders'] ?? 0);
                $sheet->setCellValue('D' . $row, $channel['L60 Orders'] ?? 0);
                $sheet->setCellValue('E' . $row, $channel['Growth'] ?? 0);
                $sheet->setCellValue('F' . $row, $channel['Gprofit%'] ?? 'N/A');
                $sheet->setCellValue('G' . $row, $channel['G ROI%'] ?? 'N/A');
                $sheet->setCellValue('H' . $row, $channel['NR'] ?? 0);
                $sheet->setCellValue('I' . $row, $channel['type'] ?? '');
                $sheet->setCellValue('J' . $row, $channel['listed_count'] ?? 0);
                $sheet->setCellValue('K' . $row, $channel['ODR'] ?? 'N/A');
                $sheet->setCellValue('L' . $row, $channel['Fulfillment Rate'] ?? 'N/A');
                $sheet->setCellValue('M' . $row, $channel['Valid Tracking Rate'] ?? 'N/A');
                $sheet->setCellValue('N' . $row, $channel['On Time Delivery Rate'] ?? 'N/A');
                $sheet->setCellValue('O' . $row, $channel['AtoZ Claims Rate'] ?? 'N/A');
                $sheet->setCellValue('P' . $row, $channel['Voilation Rate'] ?? 'N/A');
                $sheet->setCellValue('Q' . $row, $channel['Refund Rate'] ?? 'N/A');
                $sheet->setCellValue('R' . $row, $channel['sheet_link'] ?? '');
                $sheet->setCellValue('S' . $row, $channel['Account health'] ?? '');
                $row++;
            }

            // Auto-size columns
            foreach (range('A', 'S') as $column) {
                $sheet->getColumnDimension($column)->setAutoSize(true);
            }

            // Generate filename
            $filename = 'account_health_master_' . date('Y-m-d_H-i-s') . '.xlsx';

            // Save and download
            $writer = new Xlsx($spreadsheet);
            $tempFile = tempnam(sys_get_temp_dir(), 'account_health');
            $writer->save($tempFile);

            return response()->download($tempFile, $filename)->deleteFileAfterSend(true);
        } catch (\Exception $e) {
            Log::error('Export error: ' . $e->getMessage());
            return redirect()->back()->with('error', 'Export failed: ' . $e->getMessage());
        }
    }

    /**
     * Import account health data from Excel
     */
    public function import(Request $request)
    {
        $validator = Validator::make($request->all(), [
            'excel_file' => 'required|file|mimes:xlsx,xls,csv',
            'import_type' => 'required|in:channel_data,health_rates,both',
            'update_mode' => 'required|in:update,create,replace'
        ]);

        if ($validator->fails()) {
            return response()->json([
                'success' => false,
                'message' => $validator->errors()->first()
            ], 422);
        }

        try {
            $file = $request->file('excel_file');
            $importType = $request->input('import_type');
            $updateMode = $request->input('update_mode');

            // Read Excel file
            $spreadsheet = IOFactory::load($file->getPathname());
            $worksheet = $spreadsheet->getActiveSheet();
            $data = $worksheet->toArray();

            // Remove header row
            $headers = array_shift($data);

            // Process data based on import type
            $results = [];

            if ($importType === 'channel_data' || $importType === 'both') {
                $results = array_merge($results, $this->importChannelData($data, $headers, $updateMode));
            }

            if ($importType === 'health_rates' || $importType === 'both') {
                $results = array_merge($results, $this->importHealthRates($data, $headers, $updateMode));
            }

            return response()->json([
                'success' => true,
                'message' => 'Import completed successfully',
                'results' => $results
            ]);
        } catch (\Exception $e) {
            Log::error('Import error: ' . $e->getMessage());
            return response()->json([
                'success' => false,
                'message' => 'Import failed: ' . $e->getMessage()
            ], 500);
        }
    }

    /**
     * Import channel performance data
     */
    private function importChannelData($data, $headers, $updateMode)
    {
        $created = 0;
        $updated = 0;
        $skipped = 0;

        foreach ($data as $row) {
            if (empty($row[0])) continue;

            $channelName = trim($row[0]);

            // Find channel
            $channel = ChannelMaster::where('channel', $channelName)->first();
            if (!$channel) {
                $skipped++;
                continue;
            }

            // Map data from Excel columns
            $healthData = [
                'channel' => $channelName,
                'l30_sales' => $this->getColumnValue($row, $headers, 'L30 Sales', 0),
                'l30_orders' => $this->getColumnValue($row, $headers, 'L30 Orders', 0),
                'report_date' => now()->toDateString(),
                'created_by' => auth()->id(),
            ];

            // Check if record exists
            $existing = AccountHealthMaster::where('channel', $channelName)
                ->whereDate('report_date', now()->toDateString())
                ->first();

            if ($existing) {
                if ($updateMode === 'update' || $updateMode === 'replace') {
                    $existing->update($healthData);
                    $updated++;
                } else {
                    $skipped++;
                }
            } else {
                AccountHealthMaster::create($healthData);
                $created++;
            }
        }

        return [
            'Channel Data Created' => $created,
            'Channel Data Updated' => $updated,
            'Channel Data Skipped' => $skipped
        ];
    }

    /**
     * Import health rates data
     */
    private function importHealthRates($data, $headers, $updateMode)
    {
        $rateModels = [
            'ODR Rate' => OdrRate::class,
            'Fulfillment Rate' => FullfillmentRate::class,
            'Valid Tracking Rate' => ValidTrackingRate::class,
            'On Time Delivery Rate' => OnTimeDeliveryRate::class,
            'AtoZ Claims Rate' => AtoZClaimsRate::class,
            'Violation Rate' => VoilanceRate::class,
            'Refund Rate' => RefundRate::class,
        ];

        $results = [];

        foreach ($data as $row) {
            if (empty($row[0])) continue;

            $channelName = trim($row[0]);
            $channel = ChannelMaster::where('channel', $channelName)->first();

            if (!$channel) continue;

            foreach ($rateModels as $rateKey => $model) {
                $rateValue = $this->getColumnValue($row, $headers, $rateKey, 'N/A');

                if ($rateValue === null || $rateValue === '') continue;

                $existing = $model::where('channel_id', $channel->id)->first();

                $rateData = [
                    'channel_id' => $channel->id,
                    'current' => $rateValue,
                    'report_date' => now()->toDateString(),
                ];

                if ($existing) {
                    if ($updateMode === 'update' || $updateMode === 'replace') {
                        // Shift previous data
                        $existing->prev_2 = $existing->prev_1;
                        $existing->prev_2_date = $existing->prev_1_date;
                        $existing->prev_1 = $existing->current;
                        $existing->prev_1_date = $existing->report_date;

                        // Update with new data
                        $existing->update($rateData);
                        $results[$rateKey . ' Updated'] = ($results[$rateKey . ' Updated'] ?? 0) + 1;
                    } else {
                        $results[$rateKey . ' Skipped'] = ($results[$rateKey . ' Skipped'] ?? 0) + 1;
                    }
                } else {
                    $model::create($rateData);
                    $results[$rateKey . ' Created'] = ($results[$rateKey . ' Created'] ?? 0) + 1;
                }
            }
        }

        return $results;
    }

    /**
     * Get column value from row by header name
     */
    private function getColumnValue($row, $headers, $columnName, $default = null)
    {
        $index = array_search($columnName, $headers);
        return $index !== false ? ($row[$index] ?? $default) : $default;
    }

    /**
     * Download sample files
     */
    public function downloadSample($type = 'combined')
    {
        try {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();

            switch ($type) {
                case 'channel':
                    $sheet->setTitle('Channel Data Sample');
                    $headers = ['Channel', 'L30 Sales', 'L30 Orders', 'L60 Orders', 'Growth %', 'Gross Profit %', 'G ROI %', 'NR', 'Type', 'Listed Count'];
                    $sampleData = [
                        ['Amazon', 15000, 150, 280, 25.5, 35.2, 12.8, 0, 'FBA', 1250],
                        ['eBay', 8500, 85, 140, 18.3, 28.7, 8.9, 1, 'Auction', 890]
                    ];
                    break;

                case 'rates':
                    $sheet->setTitle('Health Rates Sample');
                    $headers = ['Channel', 'ODR Rate', 'Fulfillment Rate', 'Valid Tracking Rate', 'On Time Delivery Rate', 'AtoZ Claims Rate', 'Violation Rate', 'Refund Rate'];
                    $sampleData = [
                        ['Amazon', '0.5%', '98.5%', '99.2%', '95.8%', '0.2%', '0.1%', '2.3%'],
                        ['eBay', '1.2%', '97.8%', '98.9%', '94.5%', '0.5%', '0.3%', '3.1%']
                    ];
                    break;

                default: // combined
                    $sheet->setTitle('Combined Sample');
                    $headers = ['Channel', 'L30 Sales', 'L30 Orders', 'ODR Rate', 'Fulfillment Rate', 'Valid Tracking Rate', 'On Time Delivery Rate', 'AtoZ Claims Rate', 'Violation Rate', 'Sheet Link'];
                    $sampleData = [
                        ['Amazon', 15000, 150, '0.5%', '98.5%', '99.2%', '95.8%', '0.2%', '0.1%', 'https://docs.google.com/spreadsheets/d/example1'],
                        ['eBay', 8500, 85, '1.2%', '97.8%', '98.9%', '94.5%', '0.5%', '0.3%', 'https://docs.google.com/spreadsheets/d/example2']
                    ];
                    break;
            }

            // Set headers
            $col = 'A';
            foreach ($headers as $header) {
                $sheet->setCellValue($col . '1', $header);
                $sheet->getStyle($col . '1')->getFont()->setBold(true);
                $col++;
            }

            // Add sample data
            $row = 2;
            foreach ($sampleData as $rowData) {
                $col = 'A';
                foreach ($rowData as $value) {
                    $sheet->setCellValue($col . $row, $value);
                    $col++;
                }
                $row++;
            }

            // Auto-size columns
            foreach ($headers as $index => $header) {
                $column = chr(65 + $index); // A, B, C, etc.
                $sheet->getColumnDimension($column)->setAutoSize(true);
            }

            $filename = "account_health_sample_{$type}.xlsx";

            $writer = new Xlsx($spreadsheet);
            $tempFile = tempnam(sys_get_temp_dir(), 'sample');
            $writer->save($tempFile);

            return response()->download($tempFile, $filename)->deleteFileAfterSend(true);
        } catch (\Exception $e) {
            Log::error('Sample download error: ' . $e->getMessage());
            return redirect()->back()->with('error', 'Failed to generate sample file');
        }
    }
}
