<?php

namespace App\Http\Controllers\PurchaseMaster;

use App\Http\Controllers\Controller;
use App\Models\TransitContainerDetail;
use App\Models\Supplier;
use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;

class TransitContainerDetailsController extends Controller
{
    public function index()
    {
        $allRecords = TransitContainerDetail::where(function($q){
            $q->whereNull('status')->orWhereRaw("TRIM(status) = ''");
        })->get();

        $tabs = TransitContainerDetail::where(function($q){
            $q->whereNull('status')->orWhereRaw("TRIM(status) = ''");
        })->distinct()->pluck('tab_name')->toArray();

        if (empty($tabs)) {
            $tabs = ['Container 1'];
        }

        $skuParentMap = ProductMaster::pluck('parent', 'sku')
            ->mapWithKeys(function ($parent, $sku) {
                $normSku = strtoupper(trim(preg_replace('/\s+/', ' ', $sku)));
                return [$normSku => strtoupper(trim($parent))];
            })->toArray();

        $supplierData = Supplier::select('name', 'parent')->get();
        $parentSupplierMap = [];
        foreach ($supplierData as $supplier) {
            $parentList = array_map('trim', explode(',', $supplier->parent));
            foreach ($parentList as $parent) {
                $key = strtoupper(trim(preg_replace('/\s+/', ' ', $parent)));
                $parentSupplierMap[$key][] = $supplier->name;
            }
        }

        $shopifyImages = ShopifySku::pluck('image_src', 'sku')->mapWithKeys(function ($value, $key) {
            $normSku = strtoupper(trim(preg_replace('/\s+/', ' ', $key)));
            return [$normSku => $value];
        })->toArray();

        $productValuesMap = ProductMaster::pluck('Values', 'sku')->mapWithKeys(function ($value, $key) {
            $normSku = strtoupper(trim(preg_replace('/\s+/', ' ', $key)));
            return [$normSku => $value];
        })->toArray();

        // 🔥 Transform TransitContainerDetail Records
        $allRecords->transform(function ($record) use ($skuParentMap, $parentSupplierMap, $shopifyImages, $productValuesMap) {
            $sku = strtoupper(trim(preg_replace('/\s+/', ' ', $record->our_sku ?? '')));

            $parent = $skuParentMap[$sku] ?? null;

            if (empty($record->parent) && $parent) {
                $record->parent = $parent;
            }

            $parentKey = strtoupper(trim(preg_replace('/\s+/', ' ', $record->parent ?? '')));
            $record->supplier_names = $parentSupplierMap[$parentKey] ?? [];

            $record->image_src = $shopifyImages[$sku] ?? null;
            $record->Values = $productValuesMap[$sku] ?? null;

            return $record;
        });

        $groupedData = $allRecords->groupBy('tab_name');
        foreach ($tabs as $tab) {
            if (!isset($groupedData[$tab])) {
                $groupedData[$tab] = collect([]);
            }
        }

        $suppliers = Supplier::select('id', 'name')->get();

        return view('purchase-master.transit_container.index', [
            'tabs' => $tabs,
            'groupedData' => $groupedData,
            'suppliers' => $suppliers
        ]);
    }


    public function addTab(Request $request)
    {
        $tabName = trim($request->tab_name);

        if (!$tabName) {
            return response()->json(['success' => false, 'message' => 'Tab name is required.'], 400);
        }

        $exists = TransitContainerDetail::where('tab_name', $tabName)->exists();
        if ($exists) {
            return response()->json(['success' => false, 'message' => 'Tab name already exists.'], 400);
        }

        TransitContainerDetail::create([
            'tab_name' => $tabName,
        ]);

        return response()->json(['success' => true]);
    }

    public function saveRow(Request $request)
    {
        $data = $request->all();

        if (empty($data['tab_name'])) {
            return response()->json(['success' => false, 'message' => 'Tab name is missing.'], 422);
        }

        if (!empty($data['id'])) {
            $row = TransitContainerDetail::find($data['id']);
            if ($row) {
                $row->update($data);
            } else {
                return response()->json(['success' => false, 'message' => 'Row not found.']);
            }
        } else {
            $row = TransitContainerDetail::create($data);
        }

        return response()->json(['success' => true, 'id' => $row->id]);
    }

    public function uploadImage(Request $request)
    {
        if ($request->hasFile('image')) {
            $file = $request->file('image');
            $filename = time() . '_' . uniqid() . '.' . $file->getClientOriginalExtension();
            $path = public_path('uploads/transit/');
            $file->move($path, $filename);

            return response()->json([
                'success' => true,
                'url' => url('uploads/transit/' . $filename),
            ]);
        }

        return response()->json([
            'success' => false,
            'message' => 'No file uploaded',
        ]);
    }

    //save transit conatiner items
    public function transitContainerStoreItems(Request $request){
        $request->validate([
            'tab_name'       => 'required|string|max:255',
            'our_sku.*'       => 'required|string|max:255',
            'supplier_name.*' => 'required|string|max:255',
            'no_of_units.*'   => 'nullable|numeric',
            'total_ctn.*'     => 'nullable|numeric',
            'pcs_qty.*'       => 'nullable|numeric',
            'rate.*'          => 'nullable|numeric',
            'unit.*'          => 'nullable|string',
            'cbm.*'          => 'nullable|numeric',
            'changes.*'       => 'nullable|string',
            'specification.*' => 'nullable|string',
        ]);

        foreach ($request->our_sku as $index => $sku) {
            $data = [
                'tab_name'      => $request->tab_name,
                'our_sku'       => $sku,
                'supplier_name' => $request->supplier_name[$index] ?? null,
                'no_of_units'   => $request->no_of_units[$index] ?? null,
                'total_ctn'     => $request->total_ctn[$index] ?? null,
                'pcs_qty'       => $request->pcs_qty[$index] ?? null,
                'rate'          => $request->rate[$index] ?? null,
                'unit'          => $request->unit[$index] ?? null,
                'cbm'          => $request->cbm[$index] ?? null,
                'changes'       => $request->changes[$index] ?? null,
                'specification' => $request->specification[$index] ?? null,
            ];

            TransitContainerDetail::updateOrCreate(
                [
                    'tab_name' => $request->tab_name,
                    'our_sku'  => null,
                ],
                $data
            );
        }
        return back()->with('success', 'Items saved successfully!');
    }

    public function deleteTransitItem(Request $request)
    {
        $ids = $request->ids;

        TransitContainerDetail::whereIn('id', $ids)->update([
            'status' => 'inactive'
        ]);

        return response()->json([
            'success' => true,
            'message' => 'Rows marked as inactive successfully.'
        ]);
    }

    //transit container changes
    public function transitContainerChanges(){

        $allRecords = TransitContainerDetail::all();

        $tabs = TransitContainerDetail::select('tab_name')->distinct()->pluck('tab_name')->toArray();
        if (empty($tabs)) {
            $tabs = ['Container 1'];
        }

        $skuParentMap = ProductMaster::pluck('parent', 'sku')->toArray();

        $supplierData = Supplier::select('name', 'parent')->get();
        $parentSupplierMap = [];
        foreach ($supplierData as $supplier) {
            $parentList = array_map('trim', explode(',', $supplier->parent));
            foreach ($parentList as $parent) {
                $key = strtolower($parent);
                if (!isset($parentSupplierMap[$key])) {
                    $parentSupplierMap[$key] = [];
                }
                $parentSupplierMap[$key][] = $supplier->name;
            }
        }

        $shopifyImages = ShopifySku::pluck('image_src', 'sku')->mapWithKeys(function($value, $key) {
            return [strtoupper(trim($key)) => $value];
        })->toArray();

        $productValuesMap = ProductMaster::pluck('Values', 'sku')->mapWithKeys(function($value, $key) {
            return [strtoupper(trim($key)) => $value];
        })->toArray();

        // First enrich all records
        $allRecords->transform(function ($record) use ($skuParentMap, $parentSupplierMap, $shopifyImages, $productValuesMap) {
            $sku = strtoupper(trim($record->our_sku ?? ''));

            if (empty($record->parent) && isset($skuParentMap[$sku])) {
                $record->parent = $skuParentMap[$sku];
            }

            $parentKey = strtolower(trim($record->parent ?? ''));
            $record->supplier_names = $parentSupplierMap[$parentKey] ?? [];

            $record->image_src = $shopifyImages[$sku] ?? null;
            $record->Values = $productValuesMap[$sku] ?? null;

            return $record;
        });

        // Then filter out 'Sourcing' parents (after parent field is enriched)
        $filteredRecords = $allRecords->filter(function ($record) {
            return strtolower(trim($record->parent)) !== 'sourcing';
        });

        $groupedData = $filteredRecords->groupBy('tab_name');

        foreach ($tabs as $tab) {
            if (!isset($groupedData[$tab])) {
                $groupedData[$tab] = collect([]);
            }
        }
        return view('purchase-master.transit_container.changes', compact('tabs', 'groupedData'));
    }

    //transit container new
    public function transitContainerNew()
    {
        $allRecords = TransitContainerDetail::all();

        $tabs = TransitContainerDetail::select('tab_name')->distinct()->pluck('tab_name')->toArray();
        if (empty($tabs)) {
            $tabs = ['Container 1'];
        }

        $skuParentMap = ProductMaster::pluck('parent', 'sku')->toArray();

        $supplierData = Supplier::select('name', 'parent')->get();
        $parentSupplierMap = [];
        foreach ($supplierData as $supplier) {
            $parentList = array_map('trim', explode(',', $supplier->parent));
            foreach ($parentList as $parent) {
                $key = strtolower($parent);
                if (!isset($parentSupplierMap[$key])) {
                    $parentSupplierMap[$key] = [];
                }
                $parentSupplierMap[$key][] = $supplier->name;
            }
        }

        $shopifyImages = ShopifySku::pluck('image_src', 'sku')->mapWithKeys(function ($value, $key) {
            return [strtoupper(trim($key)) => $value];
        })->toArray();

        $productValuesMap = ProductMaster::pluck('Values', 'sku')->mapWithKeys(function ($value, $key) {
            return [strtoupper(trim($key)) => $value];
        })->toArray();

        $allRecords->transform(function ($record) use ($skuParentMap, $parentSupplierMap, $shopifyImages, $productValuesMap) {
            $sku = strtoupper(trim($record->our_sku ?? ''));

            if (empty($record->parent) && isset($skuParentMap[$sku])) {
                $record->parent = $skuParentMap[$sku];
            }

            $parentKey = strtolower(trim($record->parent ?? ''));
            $record->supplier_names = $parentSupplierMap[$parentKey] ?? [];

            $record->image_src = $shopifyImages[$sku] ?? null;
            $record->Values = $productValuesMap[$sku] ?? null;

            return $record;
        });

        // Filter to include ONLY 'Sourcing' parent
        $filteredRecords = $allRecords->filter(function ($record) {
            return strtolower(trim($record->parent)) === 'sourcing';
        });

        $groupedData = $filteredRecords->groupBy('tab_name');
        foreach ($tabs as $tab) {
            if (!isset($groupedData[$tab])) {
                $groupedData[$tab] = collect([]);
            }
        }

        return view('purchase-master.transit_container.new_transit', [
            'tabs' => $tabs,
            'groupedData' => $groupedData
        ]);
    }
    
}
