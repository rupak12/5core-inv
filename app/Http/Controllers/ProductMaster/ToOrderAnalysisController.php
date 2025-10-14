<?php

namespace App\Http\Controllers\ProductMaster;

use App\Http\Controllers\ApiController;
use App\Http\Controllers\Controller;
use App\Models\MfrgProgress;
use App\Models\ShopifySku;
use App\Models\Supplier;
use App\Models\ToOrderAnalysis;
use App\Models\ToOrderReview;
use Illuminate\Http\Request;
use Illuminate\Pagination\LengthAwarePaginator;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Str;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class ToOrderAnalysisController extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }

    //only for testing purpose
    public function test()
    {
                
        try {
            // Step 1: Get Product Master Sheet
            $response = $this->apiController->fetchDataFromProductMasterGoogleSheet();
            if ($response->getStatusCode() !== 200) {
                return response()->json(['message' => 'Failed to fetch Product Master data'], $response->getStatusCode());
            }

            $data = $response->getData();

            // Step 2: Get Product Master DB
            $productListData = DB::table('product_master')->orderBy('id')->get();

            // Step 3: Get Supplier Info
            $skus = collect($data->data ?? [])
                ->filter(fn($item) => !empty($item->{'SKU'}) && stripos($item->{'SKU'}, 'PARENT') === false)
                ->pluck('SKU')->unique()->toArray();

            $supplierData = \App\Models\Supplier::whereIn('sku', $skus)->get()->keyBy('sku');

            // Step 4: Filter and build result
            $processedData = [];

            foreach ($data->data ?? [] as $item) {
                $sheetSku = strtoupper(trim($item->{'SKU'} ?? ''));
                $sheetParent = strtoupper(trim($item->Parent ?? ''));

                if (empty($sheetSku) || stripos($sheetSku, 'PARENT') !== false) {
                    continue;
                }

                $prodData = $productListData->firstWhere('sku', $sheetSku);
                if ($prodData) {
                    $item->Parent = $prodData->parent ?? $item->Parent;
                    $item->SKU = $prodData->sku ?? $item->SKU;
                }

                $item->Supplier = $supplierData[$sheetSku]->supplier_name ?? '';

                $processedData[] = (object)[
                    'Parent' => $item->Parent ?? '',
                    'SKU' => $item->SKU ?? '',
                    'Approved QTY' => $item->{'Approved QTY'} ?? '',
                    'Supplier' => $item->Supplier ?? '',
                ];
            }

            return response()->json([
                'message' => 'Filtered data fetched successfully',
                'data' => $processedData,
                'status' => 200,
            ]);

        } catch (\Throwable $e) {
            return response()->json([
                'message' => 'Something went wrong!',
                'error' => $e->getMessage()
            ], 500);
        }
       
    }

    public function index(Request $request)
    {
        try {
            $search = strtolower($request->get('search', ''));
            $stageFilter = strtolower($request->get('stage', ''));

            // Fetch base data
            $toOrderRecords = DB::table('to_order_analysis')->get()->keyBy('sku');
            $productData = DB::table('product_master')->get()->keyBy(fn($item) => strtoupper(trim($item->sku)));
            $forecastData = DB::table('forecast_analysis')->get()->keyBy(fn($row) => strtoupper(trim($row->sku)));

            // ✅ Shopify image support
            $shopifySkus = ShopifySku::all()->keyBy(fn($item) => strtoupper(trim($item->sku)));

            $processedData = [];

            foreach ($toOrderRecords as $sku => $toOrder) {
                $sheetSku = strtoupper(trim($sku));
                if (empty($sheetSku)) continue;

                $product = $productData->get($sheetSku);
                $forecast = $forecastData->get($sheetSku);
                $parent = $toOrder->parent ?? $product->parent ?? '';
                $supplierName = '';

                $parentList = explode(',', $parent);
                foreach ($parentList as $singleParent) {
                    $singleParent = trim($singleParent);
                    $supplierRecord = DB::table('suppliers')
                        ->whereRaw("FIND_IN_SET(?, REPLACE(REPLACE(parent, ' ', ''), '\n', ''))", [str_replace(' ', '', $singleParent)])
                        ->first();

                    if ($supplierRecord) {
                        $supplierName = $supplierRecord->name;
                        break;
                    }
                }

                $cbm = 0;
                $imagePath = null;

                if (!empty($product?->Values)) {
                    $valuesArray = json_decode($product->Values, true);
                    if (is_array($valuesArray)) {
                        $cbm = (float)($valuesArray['cbm'] ?? 0);
                        $imagePath = $valuesArray['image_path'] ?? null;
                    }
                }

                // ✅ Image resolution
                $shopifyImage = $shopifySkus[$sheetSku]->image_src ?? null;
                $finalImage = $shopifyImage ?: $imagePath;

                $approvedQty = (int)($toOrder->approved_qty ?? 0);

                $processedData[] = (object)[
                    'Parent'          => $parent,
                    'SKU'             => $sheetSku,
                    'Approved QTY'    => $approvedQty,
                    'Date of Appr'    => $toOrder->date_apprvl ?? '',
                    'Clink'           => $forecast->clink ?? '',
                    'Supplier'        => $toOrder->supplier_name ?? $supplierName ?? '',
                    'RFQ Form Link'   => $toOrder->rfq_form_link ?? '',
                    'sheet_link'      => $toOrder->sheet_link ?? '',
                    'Rfq Report Link' => $toOrder->rfq_report_link ?? '',
                    'Stage'           => $toOrder->stage ?? '',
                    'nrl'             => $toOrder->nrl ?? '',
                    'Adv date'        => $toOrder->advance_date ?? '',
                    'order_qty'       => $toOrder->order_qty ?? '',
                    'is_parent'       => stripos($sheetSku, 'PARENT') !== false,
                    'cbm'             => $cbm,
                    'total_cbm'       => $cbm * $approvedQty,
                    'Image'           => $finalImage,
                ];
            }

            // ✅ Apply stage + search filter BEFORE pagination
            $filteredChildren = collect($processedData)->filter(function ($item) use ($search, $stageFilter) {
                $matchSearch = $search === '' || str_contains(strtolower($item->SKU . $item->Supplier . $item->Parent), $search);
                $matchStage = $stageFilter === '' || strtolower($item->Stage) === $stageFilter;

                return !$item->is_parent &&
                    ((int)$item->{'Approved QTY'} > 0) &&
                    strtolower($item->Stage ?? '') !== 'mfrg progress' &&
                    $matchSearch &&
                    $matchStage;
            })->values();

            // ✅ Pagination logic
            $page = $request->get('page', 1);
            $perPage = 25;
            $offset = ($page - 1) * $perPage;
            $paginatedChildren = $filteredChildren->slice($offset, $perPage)->values();

            $finalRows = collect();

            foreach ($paginatedChildren as $child) {
                $parent = $child->Parent;

                if (!$finalRows->contains(fn($item) => $item->is_parent && $item->Parent === $parent)) {
                    $parentRow = collect($processedData)->first(function ($item) use ($parent) {
                        return $item->is_parent && $item->Parent === $parent;
                    });

                    if ($parentRow) {
                        $totalApprovedQty = $filteredChildren->filter(fn($childItem) => $childItem->Parent === $parent)
                            ->sum(fn($item) => (int)$item->{'Approved QTY'});
                        $parentRow->{'Approved QTY'} = $totalApprovedQty;
                        $finalRows->push($parentRow);
                    }
                }

                $finalRows->push($child);
            }

            $paginator = new \Illuminate\Pagination\LengthAwarePaginator(
                $finalRows->values(),
                $filteredChildren->count(),
                $perPage,
                $page,
                ['path' => $request->url(), 'query' => $request->query()]
            );

            $groupedData = collect($processedData)->filter(fn($row) => !$row->is_parent)->groupBy('Parent');
            $processedArray = collect($processedData)->filter()->map(function ($item) {
                return (array) $item;
            });

            $filtered = $processedArray->filter(function ($item) {
                return isset($item['Approved QTY']) && floatval($item['Approved QTY']) > 0;
            });

            $groupedDataBySupplier = $filtered->groupBy('Supplier')->map(function ($group) {
                return $group->keyBy('SKU');
            });

            $groupedSupplierJson = $groupedDataBySupplier->toJson();

            $totalCBM = $filteredChildren->sum('total_cbm');
            $uniqueSuppliers = $filteredChildren->pluck('Supplier')->filter()->unique()->values();


            $viewData = [
                'data' => $paginator,
                'groupedDataJson' => $groupedData,
                'totalCBM' => $totalCBM,
                'allProcessedData' => $processedData,
                'groupedSupplierJson' => $groupedSupplierJson,
                'uniqueSuppliers' => $uniqueSuppliers,
            ];

            return $request->ajax()
                ? view('product-master.partials.to_order_table', $viewData)
                : view('product-master.to_order_analysis', $viewData);

        } catch (\Throwable $e) {
            return response()->json(['message' => 'Something went wrong!', 'error' => $e->getMessage()], 500);
        }
    }

    public function toOrderAnalysisNew(){
        return view('purchase-master.to-order-analysis');
    }

    public function getToOrderAnalysis()
    {
        try {
            // Fetch base data
            $toOrderRecords = DB::table('to_order_analysis')->get()->keyBy('sku');
            $productData = DB::table('product_master')->get()->keyBy(fn($item) => strtoupper(trim($item->sku)));
            $forecastData = DB::table('forecast_analysis')->get()->keyBy(fn($row) => strtoupper(trim($row->sku)));

            // ✅ Shopify image support
            $shopifySkus = ShopifySku::all()->keyBy(fn($item) => strtoupper(trim($item->sku)));

            $processedData = [];

            foreach ($toOrderRecords as $sku => $toOrder) {
                $sheetSku = strtoupper(trim($sku));
                if (empty($sheetSku)) continue;

                $product = $productData->get($sheetSku);
                $forecast = $forecastData->get($sheetSku);
                $parent = $toOrder->parent ?? $product->parent ?? '';
                $supplierName = '';

                // ✅ Find supplier
                $parentList = explode(',', $parent);
                foreach ($parentList as $singleParent) {
                    $singleParent = trim($singleParent);
                    $supplierRecord = DB::table('suppliers')
                        ->whereRaw("FIND_IN_SET(?, REPLACE(REPLACE(parent, ' ', ''), '\n', ''))", [str_replace(' ', '', $singleParent)])
                        ->first();

                    if ($supplierRecord) {
                        $supplierName = $supplierRecord->name;
                        break;
                    }
                }

                // ✅ Product Values
                $cbm = 0;
                $imagePath = null;

                if (!empty($product?->Values)) {
                    $valuesArray = json_decode($product->Values, true);
                    if (is_array($valuesArray)) {
                        $cbm = (float)($valuesArray['cbm'] ?? 0);
                        $imagePath = $valuesArray['image_path'] ?? null;
                    }
                }

                // ✅ Image preference
                $shopifyImage = $shopifySkus[$sheetSku]->image_src ?? null;
                $finalImage = $shopifyImage ?: $imagePath;

                $approvedQty = (int)($toOrder->approved_qty ?? 0);

                $processedData[] = [
                    'Parent'          => $parent,
                    'SKU'             => $sheetSku,
                    'approved_qty'    => $approvedQty,
                    'Date of Appr'    => $toOrder->date_apprvl ?? '',
                    'Clink'           => $forecast->clink ?? '',
                    'Supplier'        => $toOrder->supplier_name ?? $supplierName ?? '',
                    'RFQ Form Link'   => $toOrder->rfq_form_link ?? '',
                    'sheet_link'      => $toOrder->sheet_link ?? '',
                    'Rfq Report Link' => $toOrder->rfq_report_link ?? '',
                    'Stage'           => $toOrder->stage ?? '',
                    'nrl'             => $toOrder->nrl ?? '',
                    'Adv date'        => $toOrder->advance_date ?? '',
                    'order_qty'       => $toOrder->order_qty ?? '',
                    'is_parent'       => stripos($sheetSku, 'PARENT') !== false,
                    'cbm'             => $cbm,
                    'total_cbm'       => $cbm * $approvedQty,
                    'Image'           => $finalImage,
                ];
            }

            // ✅ Return JSON for Tabulator
            return response()->json([
                "data" => $processedData
            ]);

        } catch (\Throwable $e) {
            return response()->json([
                'message' => 'Something went wrong!',
                'error'   => $e->getMessage()
            ], 500);
        }
    }

    public function updateLink(Request $request)
    {
        $sku = trim(strtoupper($request->input('sku')));
        $column = $request->input('column');
        $value = $request->input('value');

        if (!in_array($column, ['approved_qty','Date of Appr', 'RFQ Form Link', 'Rfq Report Link', 'sheet_link', 'Stage', 'nrl', 'Supplier', 'order_qty', 'Adv date'])) {
            return response()->json(['success' => false, 'message' => 'Invalid column']);
        }

        $updateColumn = match ($column) {
            'Date of Appr'    => 'date_apprvl',
            'RFQ Form Link'   => 'rfq_form_link',
            'Rfq Report Link' => 'rfq_report_link',
            'Stage'           => 'stage',
            'nrl'             => 'nrl',
            'sheet_link'      => 'sheet_link',
            'Supplier'        => 'supplier_name',
            'Adv date'        => 'advance_date',
            'order_qty'       => 'order_qty',
            'approved_qty'    => 'approved_qty',
        };

        $existing = ToOrderAnalysis::whereRaw('TRIM(UPPER(sku)) = ?', [$sku])->first();

        if ($existing) {
            $existing->{$updateColumn} = $value;
            $existing->save();
        } else {
            ToOrderAnalysis::create([
                'sku' => $sku,
                $updateColumn => $value
            ]);
        }

        return response()->json(['success' => true]);
    }


    public function storeMFRG(Request $request)
    {
        try {
            $record = MfrgProgress::where('parent', $request->parent)
                ->where('sku', $request->sku)
                ->first();

            if ($record) {
                $record->qty = $request->order_qty;
                $record->supplier = $request->supplier;
                $record->adv_date = $request->adv_date;
                $record->save();
            } else {
                // Create new record
                MfrgProgress::create([
                    'parent' => $request->parent,
                    'sku' => $request->sku,
                    'qty' => $request->order_qty,
                    'supplier' => $request->supplier,
                    'adv_date' => $request->adv_date,
                ]);
            }

            return response()->json(['success' => true]);
        } catch (\Exception $e) {
            return response()->json(['success' => false, 'message' => $e->getMessage()], 500);
        }
    }


    public function storeToOrderReview(Request $request)
    {
        $data = [
            'supplier' => $request->input('supplier'),
            'positive_review' => $request->input('positive_review'),
            'negative_review' => $request->input('negative_review'),
            'improvement' => $request->input('improvement'),
            'date_updated' => $request->input('date_updated'),
        ];

        ToOrderReview::updateOrCreate(
            [
                'parent' => $request->input('parent'),
                'sku' => $request->input('sku'),
            ],
            $data
        );

        return response()->json(['success' => true]);
    }


    


}
