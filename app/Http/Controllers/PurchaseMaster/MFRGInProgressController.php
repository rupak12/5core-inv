<?php

namespace App\Http\Controllers\PurchaseMaster;

use App\Http\Controllers\Controller;
use App\Models\MfrgProgress;
use App\Models\ReadyToShip;
use App\Models\Supplier;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class MFRGInProgressController extends Controller
{
    public function index()
    {
        $mfrgData = MfrgProgress::all();

        $shopifyImages = DB::table('shopify_skus')
            ->select('sku', 'image_src')
            ->get()
            ->keyBy(fn($item) => strtoupper(trim($item->sku)));

        $productMaster = DB::table('product_master')->get()
            ->keyBy(fn($item) => strtoupper(trim($item->sku)));

        // Supplier Table Parent Mapping
        $supplierRows = Supplier::where('type', 'Supplier')->get();
        $supplierMapByParent = [];

        foreach ($supplierRows as $row) {
            $parents = array_map('trim', explode(',', strtoupper($row->parent ?? '')));
            foreach ($parents as $parent) {
                if (!empty($parent)) {
                    $supplierMapByParent[$parent][] = $row->name;
                }
            }
        }

        foreach ($mfrgData as $row) {
            $sku = strtoupper(trim($row->sku));
            $image = null;
            $cbm = null;
            $parent = null;
            $supplierNames = [];

            // Shopify Image
            if (isset($shopifyImages[$sku]) && !empty($shopifyImages[$sku]->image_src)) {
                $image = $shopifyImages[$sku]->image_src;
            }

            // Product Master Data
            if (isset($productMaster[$sku])) {
                $productRow = $productMaster[$sku];
                $values = json_decode($productRow->Values ?? '{}', true);

                if (is_array($values)) {
                    if (!empty($values['image_path'])) {
                        $image = 'storage/' . ltrim($values['image_path'], '/');
                    }
                    if (isset($values['cbm'])) {
                        $cbm = $values['cbm'];
                    }
                }

                $parent = strtoupper(trim($productRow->parent));
            }

            // Supplier from Parent Mapping
            if (!empty($parent) && isset($supplierMapByParent[$parent])) {
                $supplierNames = $supplierMapByParent[$parent];
            }

            if (!empty($row->supplier)) {
                $row->supplier = $row->supplier; // keep manual value
            } else {
                $row->supplier = implode(', ', $supplierNames); // mapping value
            }

            $row->Image = $image;
            $row->CBM = $cbm;
        }


        $suppliers = Supplier::pluck('name');
        
        return view('purchase-master.mfrg-progress.index', [
            'data' => $mfrgData,
            'suppliers' => $suppliers,
        ]);
    }

    public function newMfrgView(){
        return view('purchase-master.mfrg-progress.mfrg-new');
    }

    public function getMfrgProgressData()
    {
        $normalizeSku = fn($sku) => strtoupper(
            preg_replace('/\s+/', ' ',
                trim(
                    str_replace(["\xC2\xA0","\xE2\x80\x8B","\r","\n","\t"], ' ', $sku)
                )
            )
        );

        $mfrgData = MfrgProgress::all();

        $shopifyImages = DB::table('shopify_skus')
            ->select('sku', 'image_src')
            ->get()
            ->keyBy(fn($item) => $normalizeSku($item->sku));

        $productMaster = DB::table('product_master')
            ->get()
            ->keyBy(fn($item) => $normalizeSku($item->sku));

        $supplierRows = Supplier::where('type', 'Supplier')->get();
        $supplierMapByParent = [];
        foreach ($supplierRows as $row) {
            $parents = array_map('trim', explode(',', $normalizeSku($row->parent ?? '')));
            foreach ($parents as $parent) {
                if (!empty($parent)) {
                    $supplierMapByParent[$parent][] = $row->name;
                }
            }
        }

        $processedData = [];

        foreach ($mfrgData as $row) {
            $sku = $normalizeSku($row->sku);
            $image = null;
            $cbm = null;
            $parent = null;
            $supplierNames = [];

            if (isset($shopifyImages[$sku]) && !empty($shopifyImages[$sku]->image_src)) {
                $image = $shopifyImages[$sku]->image_src;
            }

            if (isset($productMaster[$sku])) {
                $productRow = $productMaster[$sku];
                $values = json_decode($productRow->Values ?? '{}', true);

                if (is_array($values)) {
                    if (!empty($values['image_path'])) {
                        $image = 'storage/' . ltrim($values['image_path'], '/');
                    }
                    if (isset($values['cbm'])) {
                        $cbm = $values['cbm'];
                    }
                }

                $parent = $normalizeSku($productRow->parent ?? '');
            }

            if (!empty($parent) && isset($supplierMapByParent[$parent])) {
                $supplierNames = $supplierMapByParent[$parent];
            }

            $row->supplier = !empty($row->supplier) ? $row->supplier : implode(', ', $supplierNames);
            $row->Image = $image;
            $row->CBM = $cbm;

            $processedData[] = $row;
        }

        return response()->json([
            "data" => $processedData
        ]);
    }


    public function convert(Request $request)
    {
        $amount = $request->query('amount', 1);
        $from = $request->query('from', 'USD');
        $to = $request->query('to', 'CNY');

        try {
            $apiUrl = "https://api.frankfurter.app/latest?amount=$amount&from=$from&to=$to";
            $response = Http::get($apiUrl);

            if ($response->successful()) {
                return response()->json($response->json());
            } else {
                return response()->json(['error' => 'Frankfurter API error'], 500);
            }
        } catch (\Exception $e) {
            return response()->json(['error' => $e->getMessage()], 500);
        }
    }

    public function inlineUpdateBySku(Request $request)
    {
        $sku = $request->input('sku');
        $column = $request->input('column');

        $validColumns = [
            'advance_amt', 'pay_conf_date', 'o_links', 'adv_date', 'del_date', 'total_cbm',
            'barcode_sku', 'artwork_manual_book', 'notes', 'ready_to_ship', 'rate', 'rate_currency',
            'photo_packing', 'photo_int_sale','supplier','created_at'
        ];

        if (!in_array($column, $validColumns)) {
            return response()->json(['success' => false, 'message' => 'Invalid column.']);
        }

        $progress = MfrgProgress::where('sku', $sku)->first();
        if (!$progress) {
            return response()->json(['success' => false, 'message' => 'SKU not found.']);
        }

        if ($request->hasFile('value') && in_array($column, ['photo_packing', 'photo_int_sale', 'barcode_sku'])) {
            $file = $request->file('value');
            $filename = uniqid() . '_' . time() . '.' . $file->getClientOriginalExtension();
            $destinationPath = public_path('uploads/mfrg_images');
            
            if (!file_exists($destinationPath)) {
                mkdir($destinationPath, 0777, true);
            }

            $file->move($destinationPath, $filename);
            $url = url("uploads/mfrg_images/{$filename}");

            $progress->{$column} = $url;
            $progress->save();

            return response()->json(['success' => true, 'url' => $url]);
        }

        if ($column === 'advance_amt') {
            if (!$progress->supplier) {
                return response()->json(['success' => false, 'message' => 'Supplier not found.']);
            }

            MfrgProgress::where('supplier', $progress->supplier)->update([
                'advance_amt' => $request->input('value')
            ]);

            return response()->json(['success' => true, 'message' => 'Advance updated.']);
        }

        $progress->{$column} = $request->input('value');
        $progress->save();

        return response()->json(['success' => true]);
    }


    public function storeDataReadyToShip(Request $request)
    {
        try {
            $data = [
                'supplier' => $request->supplier,
                'cbm' => $request->totalCbm,
            ];

            $readyToShip = ReadyToShip::where('parent', $request->parent)
                ->where('sku', $request->sku)
                ->first();

            if ($readyToShip) {
                $readyToShip->update($data);
            } else {
                ReadyToShip::create(array_merge([
                    'parent' => $request->parent,
                    'sku' => $request->sku,
                ], $data));
            }

            return response()->json(['success' => true]);
        } catch (\Exception $e) {
            return response()->json(['success' => false, 'message' => $e->getMessage()], 500);
        }
    }


}
