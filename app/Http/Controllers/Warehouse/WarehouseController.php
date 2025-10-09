<?php

namespace App\Http\Controllers\Warehouse;

use App\Http\Controllers\Controller;
use App\Models\IncomingData;
use Illuminate\Http\Request;
use App\Models\Warehouse;
use App\Models\Inventory;
use Carbon\Carbon;
use Illuminate\Http\JsonResponse;

class WarehouseController extends Controller
{
    /**
     * Display a listing of the resource.
     */
    public function index()
    {
        $warehouses = Warehouse::latest()->get();
        return view('warehouses.all-warehouse', compact('warehouses'));
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
        $validated = $request->validate([
            'name' => 'required|string|max:255',
            'group' => 'required|string',
            'location' => 'required|string',
        ]);

        Warehouse::create($validated);

        return response()->json(['success' => true, 'message' => 'Warehouse added successfully.']);
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
        return response()->json(Warehouse::findOrFail($id));
    }

    /**
     * Update the specified resource in storage.
     */
    public function update(Request $request, string $id)
    {
        $validated = $request->validate([
            'name' => 'required|string|max:255',
            'group' => 'required|string',
            'location' => 'required|string',
        ]);

        $warehouse = Warehouse::findOrFail($id);
        $warehouse->update($validated);

        return response()->json(['success' => true, 'message' => 'Warehouse updated successfully.']);
    }

    /**
     * Remove the specified resource from storage.
     */
    public function destroy(string $id)
    {
       $warehouse = Warehouse::findOrFail($id);
        $warehouse->delete();

        return response()->json(['success' => true, 'message' => 'Warehouse deleted Successfully']);

    }

    public function list()
    {
        return response()->json(Warehouse::all());
    }

    public function viewWarehouseSkus($id): JsonResponse
    {
        // $warehouse = Warehouse::findOrFail($id);

        // $skus = Inventory::select('sku')
        //     ->select('verified_stock', 'sku','warehouse_id','approved_by','created_at')
        //     ->where('warehouse_id', $id)
        //     ->groupBy('sku')
        //     ->get();

        $skus = Inventory::with('warehouse')
            ->where('warehouse_id', $id)
            ->select('sku', 'verified_stock', 'warehouse_id', 'approved_by', 'created_at')
            ->orderBy('created_at', 'desc')
            ->get();

        $data = $skus->map(function ($item) {
            return [
                'sku' => $item->sku,
                'total_quantity' => $item->verified_stock,
                'warehouse_name' => $item->warehouse->name ?? '',
                'approved_by' => $item->approved_by,
                'date' => Carbon::parse($item->created_at)->format('d M Y, h:i A'),
            ];
        }); 

        return response()->json(['data' => $data]);

        // return view('warehouses.warehouse-skus-listing', compact('warehouse', 'skus'));
    }


    public function returnGodown(Request $request)
    {
        $warehouseId = 1;
        $data = Inventory::with('warehouse')->where('warehouse_id', $warehouseId)->get();

        if ($request->ajax()) {
            return response()->json(['data' => $data]);
        }
        
        return view('warehouses.return-godown', compact('data'));
    }

    public function openBoxGodown(Request $request)
    {
        $warehouseId = 2;
        $openboxData = Inventory::with('warehouse')->where('warehouse_id', $warehouseId)->get();
        $incomingData = IncomingData::with('warehouse')->where('warehouse_id', $warehouseId)->get();
        $data = $openboxData->merge($incomingData)->values();

        if ($request->ajax()) {
            return response()->json(['data' => $data]);
        }

        return view('warehouses.openbox-godown', compact('data'));
    }

    public function showroomGodown(Request $request)
    {
        $warehouseId = 3;
        $data = Inventory::with('warehouse')->where('warehouse_id', $warehouseId)->get();

        if ($request->ajax()) {
            return response()->json(['data' => $data]);
        }

        return view('warehouses.showroom-godown', compact('data'));
    }

    public function usedItemGodown(Request $request)
    {
        $warehouseId = 4;
        $data = Inventory::with('warehouse')->where('warehouse_id', $warehouseId)->get();

        if ($request->ajax()) {
            return response()->json(['data' => $data]);
        }

        return view('warehouses.useditem-godown', compact('data'));
    }

    public function trashGodown(Request $request)
    {
        $warehouseId = 5;
        $data = Inventory::with('warehouse')->where('warehouse_id', $warehouseId)->get();

        if ($request->ajax()) {
            return response()->json(['data' => $data]);
        }

        return view('warehouses.trash-godown', compact('data'));
    }


    public function mainGodown(Request $request)
    {
        $warehouseId = 6;
        $data = IncomingData::with('warehouse')->where('warehouse_id', $warehouseId)->get();

        if ($request->ajax()) {
            return response()->json(['data' => $data]);
        }

        return view('warehouses.main-godown', compact('data'));
    }


}


