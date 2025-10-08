@extends('layouts.vertical', ['title' => 'MFRG In Progress', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
@section('css')
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
<style>
    .custom-select-wrapper {
        position: relative;
        font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
    }
    .custom-select-box {
        transition: border-color 0.18s, box-shadow 0.18s;
        background: #fff;
    }
    .custom-select-box.active, .custom-select-box:focus-within {
        border-color: #3bc0c3;
        box-shadow: 0 0 0 2px #3bc0c340;
    }
    .custom-select-dropdown {
        animation: fadeIn 0.18s;
    }
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(-8px);}
        to { opacity: 1; transform: translateY(0);}
    }
    .custom-select-option {
        cursor: pointer;
        font-size: 1rem;
        transition: background 0.13s, color 0.13s;
        margin: 0 4px;
        color: #222;
        background: #fff;
        border-radius: 6px;
        user-select: none;
    }
    .custom-select-option.selected,
    .custom-select-option:hover,
    .custom-select-option.bg-primary {
        background: #3bc0c3 !important;
        color: #fff !important;
    }
    .custom-select-option:not(:last-child) {
        margin-bottom: 2px;
    }
    .custom-select-dropdown::-webkit-scrollbar {
        width: 7px;
        background: #f4f6fa;
        border-radius: 6px;
    }
    .custom-select-dropdown::-webkit-scrollbar-thumb {
        background: #e0e6ed;
        border-radius: 6px;
    }
    .preview-popup {
        position: fixed;
        display: none;
        z-index: 9999;
        pointer-events: none;
        width: 350px;
        height: 350px;
        object-fit: cover;
        border-radius: 8px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        transition: all 0.2s ease;
    }
    td {
        overflow: visible !important;
    }
    
</style>
@endsection
@section('content')
@include('layouts.shared.page-title', ['page_title' => 'MFRG In Progress', 'sub_title' => 'MFRG In Progress'])
<div class="row">
    <div class="col-12">
        <div class="card shadow-sm">
            <div class="card-body">
                <div class="d-flex justify-content-between align-items-center mb-4">
                    <h4 class="mb-0 font-weight-bold">
                        <i class="mdi mdi-factory mr-2" style="color:#3bc0c3;"></i>
                        MFRG In Progress
                    </h4>
                </div>

                <div class="column-controls card mb-4 p-3 shadow-sm" id="columnControls" style="background: #f8f9fa; border-radius: 10px;">
                    <div class="d-flex flex-wrap align-items-center justify-content-between gap-3">
                        <div class="d-flex align-items-center gap-3 flex-wrap w-100">
                            <!-- Title -->
                            {{-- ▶️ Navigation --}}
                            <div class="col-auto">
                                <label class="form-label fw-semibold mb-1 d-block">▶️ Navigation</label>
                                <div class="btn-group time-navigation-group" role="group">
                                    <button id="play-backward" class="btn btn-light rounded-circle shadow-sm me-2" title="Previous parent">
                                        <i class="fas fa-step-backward"></i>
                                    </button>
                                    <button id="play-pause" class="btn btn-light rounded-circle shadow-sm me-2" style="display: none;" title="Pause">
                                        <i class="fas fa-pause"></i>
                                    </button>
                                    <button id="play-auto" class="btn btn-primary rounded-circle shadow-sm me-2" title="Play">
                                        <i class="fas fa-play"></i>
                                    </button>
                                    <button id="play-forward" class="btn btn-light rounded-circle shadow-sm" title="Next parent">
                                        <i class="fas fa-step-forward"></i>
                                    </button>
                                </div>
                            </div>

                            <!-- Search Table -->
                            <input type="text" class="form-control" id="wholeSearchInput"
                                placeholder="🔍 Search entire table..."
                                style="width: 200px; font-size: 0.97rem; height: 36px; border-radius: 6px;">

                            <!-- Toggle Columns Dropdown -->
                            <div class="column-dropdown position-relative">
                                <button class="btn text-white column-dropdown-btn d-flex align-items-center gap-1" id="columnDropdownBtn" style="border-radius: 6px;">
                                    <i class="mdi mdi-format-columns"></i> Toggle Columns
                                </button>
                                <div class="column-dropdown-content" id="columnDropdownContent"
                                    style="position: absolute; left: 0; top: 110%; min-width: 220px; z-index: 20; background: #fff; box-shadow: 0 2px 12px rgba(60,192,195,0.10); border-radius: 8px; border: 1px solid #e3e3e3; padding: 12px; max-height: 350px; overflow-y: auto;">
                                    <!-- Dynamic Checkboxes -->
                                </div>
                            </div>

                            <!-- Show All Columns -->
                            <button class="btn text-white show-all-columns d-flex align-items-center gap-1" id="showAllColumns" style="border-radius: 6px;">
                                <i class="mdi mdi-eye-check-outline"></i> Show All
                            </button>

                            <!-- Supplier Dropdown -->
                            <div class="custom-select-wrapper" style="min-width: 150px; position: relative;">
                                <div class="custom-select-box d-flex align-items-center justify-content-between" id="customSelectBox"
                                    style="border: 1.5px solid #e0e6ed; border-radius: 7px; background: #fff; height: 38px; padding: 0 14px; cursor: pointer; box-shadow: 0 1px 4px rgba(60,192,195,0.07); transition: border-color 0.2s;">
                                    <span id="customSelectSelectedText" class="flex-grow-1 text-truncate" style="font-size: 1rem; color: #222;">Select supplier</span>
                                    <i class="mdi mdi-menu-down" style="font-size: 1.3rem; color: #3bc0c3;"></i>
                                </div>
                                <div class="custom-select-dropdown shadow" id="customSelectDropdown"
                                    style="display: none; position: absolute; z-index: 30; background: #fff; min-width: 220px; max-width: 320px; border-radius: 10px; border: 1.5px solid #e0e6ed; margin-top: 4px;">
                                    <div class="p-2 border-bottom" style="background: #f8f9fa; border-radius: 10px 10px 0 0;">
                                        <input type="text" class="form-control border-0 shadow-none" id="customSelectSearchInput"
                                            placeholder="🔍 Search supplier..." style="font-size: 0.97rem; height: 32px; border-radius: 6px; background: #f4f6fa;">
                                    </div>
                                    <div id="customSelectOptions" style="max-height: 160px; overflow-y: auto; padding: 2px 0;">
                                        <div class="custom-select-option px-3 py-2 rounded" data-value="">Select supplier</div>
                                        @foreach ($suppliers as $item)
                                            <div class="custom-select-option px-3 py-2 rounded" data-value="">{{ $item }}</div>
                                        @endforeach
                                    </div>
                                </div>
                            </div>

                            <!-- 💰 Advance + Pending Summary -->
                            <div id="advance-total-wrapper" style="display: none;">
                                <div id="advance-total-display" class="py-1 px-2 rounded shadow-sm d-inline-flex align-items-center gap-2 flex-wrap"
                                    style="background: linear-gradient(90deg, #e6f4f1 60%, #f8f9fa 100%); color: #10635b; font-weight: 600; font-size: 16px; border: 1.5px solid #3bc0c3; box-shadow: 0 2px 8px rgba(60,192,195,0.08); transition: all 0.3s ease;">
                                    <div class="d-flex align-items-center gap-2" style="min-width: 170px;">
                                        <span class="rounded-circle d-flex align-items-center justify-content-center" style="background: #d1f2eb; width: 36px; height: 36px;">
                                            <i class="mdi mdi-cash-multiple" style="font-size: 22px; color: #10b39c;"></i>
                                        </span>
                                        <span>
                                            <span style="font-size: 13px; color: #23979b;">Total Advance</span><br>
                                            <span style="font-size: 18px; color: #10635b;">$ <span id="advance-amount">0</span></span>
                                        </span>
                                    </div>

                                    <div class="vr" style="height: 38px; width: 2px; background: #cde7e2; margin: 0 18px;"></div>

                                    <div class="d-flex align-items-center gap-2" style="min-width: 170px;">
                                        <span class="rounded-circle d-flex align-items-center justify-content-center" style="background: #ffeaea; width: 36px; height: 36px;">
                                            <i class="mdi mdi-alert-decagram" style="font-size: 22px; color: #ff6b6b;"></i>
                                        </span>
                                        <span>
                                            <span style="font-size: 13px; color: #ff6b6b;">Total Pending</span><br>
                                            <span style="font-size: 18px; color: #b23c3c;">$ <span id="pending-amount">0</span></span>
                                        </span>
                                    </div>
                                </div>
                            </div>

                            <div class="mb-3" style="min-width: 120px; position: relative;">
                                <label class="form-label fw-semibold mb-1 d-block">Pending Status</label>
                                <select id="row-data-pending-status" class="form-select border border-primary">
                                    <option value="">select color</option>
                                    <option value="green">Green <span id="greenCount"></span></option>
                                    <option value="yellow">yellow <span id="yellowCount"></span></option>
                                    <option value="red">red <span id="redCount"></span></option>
                                </select>
                            </div>

                            <!-- Other Stats -->
                            <div class="py-1 px-3 bg-dark rounded shadow-sm d-inline-flex align-items-center gap-2 text-white fw-bold fs-6 border border-light">
                                <span>Total Amount: <span id="total-amount">0</span></span>
                            </div>
                            <div class="py-1 px-3 bg-info rounded shadow-sm d-inline-flex align-items-center gap-2 text-white fw-bold fs-6 border border-light">
                                <span>Total Ord. Qty: <span id="total-order-qty">0</span></span>
                            </div>
                            <div class="py-1 px-3 rounded shadow-sm d-inline-flex align-items-center gap-2 text-white fw-bold fs-6 border border-light" 
                                style="background: #23979b;">
                                <span>Total CBM: <span id="total-cbm">0</span></span>
                            </div>
                        </div>
                    </div>
                </div>
                            
                <div class="wide-table-wrapper table-container">
                    <table class="wide-table">
                        <thead>
                            <tr>
                                <th data-column="1">Image<div class="resizer"></div></th>
                                <th data-column="2">
                                    Parent
                                    <div class="resizer"></div>
                                    <input type="text" class="form-control column-search" data-search-column="2" placeholder="Search Parent..." style="margin-top:4px; font-size:12px; height:28px;">
                                    <div class="search-results" data-results-column="2" style="position:relative; z-index:10;"></div>
                                </th>
                                <th data-column="3">
                                    SKU
                                    <div class="resizer"></div>
                                    <input type="text" class="form-control column-search" data-search-column="3" placeholder="Search SKU..." style="margin-top:4px; font-size:12px; height:28px; width: 150px;">
                                    <div class="search-results" data-results-column="3" style="position:relative; z-index:10;"></div>
                                </th>
                                <th data-column="4" class="text-center">Order<br/>QTY<div class="resizer"></div></th>
                                <th data-column="5">Rate<div class="resizer"></div></th>
                                <th data-column="6" class="text-center">Supplier<div class="resizer"></div></th>
                                <th data-column="7" hidden>Advance<br/>Amt<div class="resizer"></div></th>
                                <th data-column="8" hidden>Adv<br/>Date<div class="resizer"></div></th>
                                <th data-column="9" hidden>pay conf.<br/>date<div class="resizer"></div></th>
                                {{-- <th data-column="9">pay term<div class="resizer"></div></th> --}}
                                <th data-column="10" class="text-center">Order<br/>Date<div class="resizer"></div></th>
                                <th data-column="11">Del<br/>Date<div class="resizer"></div></th>
                                {{-- <th data-column="12">O Links<div class="resizer"></div></th> --}}
                                <th data-column="12" hidden>value<div class="resizer"></div></th>
                                <th data-column="13">Payment<br/>Pending<div class="resizer"></div></th>
                                {{-- <th data-column="15">photo<br/>packing<div class="resizer"></div></th> --}}
                                {{-- <th data-column="16">photo int.<br/>sale<div class="resizer"></div></th> --}}
                                <th data-column="14">CBM<div class="resizer"></div></th>
                                <th data-column="15" hidden>total<br/>cbm<div class="resizer"></div></th>
                                {{-- <th data-column="19" class="text-center">BARCODE<br/>&<br/>SKU<div class="resizer"></div></th> --}}
                                {{-- <th data-column="20">artwork<br/>&<br/>maual<br/>book<div class="resizer"></div></th> --}}
                                {{-- <th data-column="21">notes<div class="resizer"></div></th> --}}
                                <th data-column="16">Ready to<br/>ship<div class="resizer"></div></th>
                            </tr>
                        </thead>
                        <tbody>
                            @foreach ($data as $item)
                                @php
                                    $readyToShip = $item->ready_to_ship ?? '';
                                @endphp
                                @continue($readyToShip === 'Yes')
                                <tr>
                                    <td data-column="1">
                                        @if(!empty($item->Image))
                                            <img src="{{ $item->Image }}" class="hover-img" data-src="{{ $item->Image }}" alt="Image" style="width: 40px; height: 40px; object-fit: cover; border-radius: 6px;">
                                        @else
                                            <span class="text-muted">No</span>
                                        @endif
                                    </td>
                                    <td data-column="2" class="text-center">
                                        {{ $item->parent ?? '' }}
                                    </td>
                                    <td data-column="3" class="text-center">
                                        <span class="sku-short" style="cursor:pointer;">
                                            {{ \Illuminate\Support\Str::limit($item->sku ?? '', 10) }}
                                        </span>
                                        <span class="sku-full d-none" id="sku-full">{{ $item->sku ?? '' }}</span>
                                    </td>
                                    <td data-column="4" data-qty="{{ $item->qty ?? 0 }}" style="text-align: end;">{{ $item->qty ?? '' }}</td>
                                    <td data-column="5">
                                        <div class="input-group input-group-sm" style="width:105px;">
                                            <span class="input-group-text" style="padding: 0 6px;">
                                                <select data-sku="{{ $item->sku }}" data-column="rate_currency" class="form-select form-select-sm currency-select auto-save" style="border: none; background: transparent; font-size: 13px; padding: 0 2px;">
                                                    <option value="USD" {{ ($item->rate_currency ?? '') == 'USD' ? 'selected' : '' }}>$</option>
                                                    <option value="CNY" {{ ($item->rate_currency ?? '') == 'CNY' ? 'selected' : '' }}>¥</option>
                                                </select>
                                            </span>
                                            <input data-sku="{{ $item->sku }}" data-column="rate" type="number" value="{{ $item->rate ?? '' }}" 
                                                class="form-control form-control-sm amount-input auto-save" style="background: #f9f9f9; font-size: 13px;" />
                                        </div>
                                    </td>

                                    <td data-column="6" class="text-center">
                                        <input type="text" class="form-control form-control-sm auto-save" data-sku="{{ $item->sku }}" data-column="supplier" value="{{ $item->supplier ?? '' }}" placeholder="supplier">
                                    </td>
                                    <td data-column="7" hidden>
                                        @php
                                            $supplier = $item->supplier ?? '';
                                            $grouped = collect($data)->where('supplier', $supplier);
                                            $firstAdvance = $grouped->first()->advance_amt ?? 0;
                                        @endphp

                                        @if ($loop->first || $item->sku === $grouped->first()->sku)
                                            <input type="number"
                                                class="form-control form-control-sm auto-save"
                                                data-sku="{{ $item->sku }}"
                                                data-supplier="{{ $supplier }}"
                                                data-column="advance_amt"
                                                value="{{ $firstAdvance }}"
                                                min="0"
                                                max="10000"
                                                step="0.01"
                                                style="min-width: 90px; max-width: 120px; font-size: 13px;"
                                                placeholder="Advance Amt"
                                                onchange="if(this.value > 10000) {
                                                    alert('Amount cannot exceed 10000');
                                                    this.value = '';
                                                    return false;
                                                }">
                                        @else
                                            <input type="number" class="form-control form-control-sm"
                                                value="{{ $firstAdvance }}"
                                                disabled
                                                style="min-width: 90px; max-width: 120px; font-size: 13px; background: #e9ecef;" />
                                        @endif
                                    </td>
                                    <td data-column="8" hidden>
                                        <input type="date" value="{{ !empty($item->adv_date) ? \Carbon\Carbon::parse($item->adv_date)->format('Y-m-d') : '' }}" data-sku="{{ $item->sku }}" 
                                        data-column="adv_date" class="form-control form-control-sm auto-save" style="width: 80px; font-size: 13px;">
                                    </td>
                                    <td data-column="9" hidden>
                                        <input type="date" value="{{ !empty($item->pay_conf_date) ? \Carbon\Carbon::parse($item->pay_conf_date)->format('Y-m-d') : '' }}" data-sku="{{ $item->sku }}"
                                         data-column="pay_conf_date" class="form-control form-control-sm auto-save" style="width: 80px; font-size: 13px;">
                                    </td>
                                    @php
                                        $bgColor = '';
                                        $daysDiff = null;

                                        if (!empty($item->created_at)) {
                                            $daysDiff = \Carbon\Carbon::parse($item->created_at)->diffInDays(\Carbon\Carbon::today());

                                            if ($daysDiff > 45) {
                                                $bgColor = 'background-color: red; color: white;';
                                            } elseif ($daysDiff > 30) {
                                                $bgColor = 'background-color: yellow; color: black;';
                                            }else{
                                                $bgColor = 'background-color: green; color: white;';
                                            }
                                        }
                                    @endphp
                                    <td data-column="10">
                                        <input type="date" data-sku="{{ $item->sku }}" data-column="created_at" value="{{ !empty($item->created_at) ? \Carbon\Carbon::parse($item->created_at)->format('Y-m-d') : '' }}" 
                                        class="form-control form-control-sm auto-save" style="width: 80px; font-size: 13px; {{ $bgColor }}">
                                        @if ($daysDiff !== null)
                                            <small style="font-size: 12px; color: rgb(72, 69, 69);">
                                                {{ $daysDiff }} days ago
                                            </small>
                                        @endif
                                    </td>
                                    <td data-column="11">
                                        <input type="date" data-sku="{{ $item->sku }}" data-column="del_date" 
                                            value="{{ !empty($item->del_date) ? \Carbon\Carbon::parse($item->del_date)->format('Y-m-d') : '' }}" 
                                        class="form-control form-control-sm auto-save" style="width: 80px; font-size: 13px;">
                                    </td>
                                    {{-- <td data-column="12">
                                        <div class="input-group input-group-sm align-items-center" style="gap: 4px;">
                                            <span class="input-group-text open-link-icon border-0 p-0 bg-transparent" style="{{ empty($item->o_links) ? 'display:none;' : '' }}; background: none !important;">
                                                <a href="{{ $item->o_links ?? '#' }}" target="_blank" title="Open Link" style="color: #3bc0c3; font-size: 20px; display: flex; align-items: center; background: none;">
                                                    <i class="mdi mdi-open-in-new" style="transition: color 0.2s; cursor: pointer;"></i>
                                                </a>
                                            </span>
                                            <span class="input-group-text edit-link-icon border-0 p-0 bg-transparent" style="cursor:pointer; background: none !important;">
                                                <a href="javascript:void(0);" class="edit-o-links" title="Edit" style="color: #6c757d; font-size: 20px; display: flex; align-items: center; background: none;">
                                                    <i class="mdi mdi-pencil-outline" style="transition: color 0.2s; cursor: pointer;"></i>
                                                </a>
                                            </span>
                                                </a>
                                            </span>
                                            <input type="text" class="form-control form-control-sm o-links-input d-none auto-save" value="{{ $item->o_links ?? '' }}" data-sku="{{ $item->sku }}" data-column="o_links" placeholder="Paste or type link here..." style="font-size: 13px; min-width: 180px; border-radius: 20px; box-shadow: 0 1px 4px rgba(60,192,195,0.08); border: 1px solid #e3e3e3; padding-left: 14px; background: #f8fafd;">
                                        </div>
                                    </td> --}}

                                    <td class="total-value d-none" data-column="12">
                                        {{ is_numeric($item->qty ?? null) && is_numeric($item->rate ?? null) ? ($item->qty * $item->rate) : '' }}
                                    </td>
                                    <td data-column="13">
                                        @php
                                            $supplier = $item->supplier ?? '';
                                            $grouped = collect($data)->where('supplier', $supplier);

                                            $supplierAdvance = $grouped->first()->advance_amt ?? 0;

                                            $totalValue = $grouped->sum(function ($row) {
                                                return (is_numeric($row->qty) && is_numeric($row->rate)) ? $row->qty * $row->rate : 0;
                                            });

                                            $thisRowValue = (is_numeric($item->qty ?? null) && is_numeric($item->rate ?? null)) ? $item->qty * $item->rate : 0;

                                            $rowAdvance = $totalValue > 0 ? ($thisRowValue / $totalValue) * $supplierAdvance : 0;
                                            $pending = $thisRowValue - $rowAdvance;
                                        @endphp
                                        {{ number_format($pending, 0) }}
                                    </td>

                                    {{-- <td data-column="15">
                                        <div class="image-upload-field d-flex align-items-center gap-2">
                                            @if(!empty($item->photo_packing))
                                                <a href="{{ $item->photo_packing }}" target="_blank" class="me-1" title="View Photo" style="width:50px;">
                                                    <img src="{{ $item->photo_packing }}" class="img-thumbnail border" style="height: 50px; width: 50px; object-fit: cover; background: #f8f9fa;">
                                                </a>
                                            @endif

                                            <label class="btn btn-sm btn-outline-primary d-flex align-items-center mb-0" style="padding: 4px 10px; border-radius: 6px; font-size: 15px;" title="Upload Photo">
                                                <i class="mdi mdi-upload" style="font-size: 18px; margin-right: 4px;"></i>
                                                <span style="font-size: 13px;"></span>
                                                <input type="file" class="d-none auto-upload" data-column="photo_packing" data-sku="{{ $item->sku }}">
                                            </label>
                                        </div>
                                    </td> --}}

                                    {{-- <td data-column="16">
                                        <div class="image-upload-field d-flex align-items-center gap-2">
                                            @if(!empty($item->photo_int_sale))
                                                <a href="{{ $item->photo_int_sale }}" target="_blank" class="me-1" title="View Photo" style="width:50px;">
                                                    <img src="{{ $item->photo_int_sale }}" class="img-thumbnail border" style="height: 50px; width: 50px; object-fit: cover; background: #f8f9fa;">
                                                </a>
                                            @endif

                                            <label class="btn btn-sm btn-outline-primary d-flex align-items-center mb-0" style="padding: 4px 10px; border-radius: 6px; font-size: 15px;" title="Upload Photo">
                                                <i class="mdi mdi-upload" style="font-size: 18px; margin-right: 4px;"></i>
                                                <span style="font-size: 13px;"></span>
                                                <input type="file" class="d-none auto-upload" data-column="photo_int_sale" data-sku="{{ $item->sku }}">
                                            </label>
                                        </div>
                                    </td> --}}

                                    <td data-column="14">
                                        {{ isset($item->CBM) ? number_format($item->CBM, 4) : 'N/A' }}
                                    </td>

                                    <td data-column="15" hidden>
                                        <input type="number"
                                            data-sku="{{ $item->sku }}"
                                            data-column="total_cbm"
                                            step="0.000000001"
                                            value="{{ is_numeric($item->qty ?? null) && is_numeric($item->CBM ?? null) ? number_format($item->qty * $item->CBM, 2, '.', '') : '' }}"
                                            class="form-control form-control-sm auto-save"
                                            style="min-width: 90px; width: 100px; font-size: 13px;"
                                            placeholder="Total CBM"
                                            readonly>
                                    </td>

                                    {{-- <td data-column="19">
                                        <div class="image-upload-field d-flex align-items-center gap-2">
                                            @if(!empty($item->barcode_sku))
                                                <a href="{{ $item->barcode_sku }}" target="_blank" class="me-1" title="View Photo" style="width:50px;">
                                                    <img src="{{ $item->barcode_sku }}" class="img-thumbnail border" style="height: 50px; width: 50px; object-fit: cover; background: #f8f9fa;">
                                                </a>
                                            @endif

                                            <label class="btn btn-sm btn-outline-primary d-flex align-items-center mb-0" style="padding: 4px 10px; border-radius: 6px; font-size: 15px;" title="Upload Photo">
                                                <i class="mdi mdi-upload" style="font-size: 18px; margin-right: 4px;"></i>
                                                <span style="font-size: 13px;"></span>
                                                <input type="file" class="d-none auto-upload" data-column="barcode_sku" data-sku="{{ $item->sku }}">
                                            </label>
                                        </div>
                                    </td>

                                    <td data-column="20">
                                        <input type="text" class="form-control form-control-sm auto-save" data-sku="{{ $item->sku }}" data-column="artwork_manual_book" value="{{ $item->artwork_manual_book ?? '' }}" placeholder="Artwork Manual Book">
                                    </td>

                                    <td data-column="21">
                                        <input type="text" class="form-control form-control-sm auto-save" data-sku="{{ $item->sku }}" data-column="notes" value="{{ $item->notes ?? '' }}" style="font-size: 13px;" placeholder="Notes">
                                    </td> --}}

                                    <td data-column="16">
                                        <select class="form-select form-select-sm auto-save" data-sku="{{ $item->sku }}" data-column="ready_to_ship" style="width: 75px;">
                                            <option value="No" {{ $item->ready_to_ship == 'No' ? 'selected' : '' }}>No</option>
                                            <option value="Yes" {{ $item->ready_to_ship == 'Yes' ? 'selected' : '' }}>Yes</option>
                                        </select>
                                    </td>
                                </tr>
                            @endforeach
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
</div>

<script>
    document.body.style.zoom = '85%';

    const popup = document.createElement('img');
    popup.className = 'preview-popup';
    document.body.appendChild(popup);

    document.querySelectorAll('.hover-img').forEach(img => {
        img.addEventListener('mouseenter', e => {
            popup.src = img.dataset.src;
            popup.style.display = 'block';
        });
        img.addEventListener('mousemove', e => {
            popup.style.top = (e.clientY + 20) + 'px';
            popup.style.left = (e.clientX + 20) + 'px';
        });
        img.addEventListener('mouseleave', e => {
            popup.style.display = 'none';
        });
    });

    document.addEventListener('DOMContentLoaded', function () {
        document.documentElement.setAttribute("data-sidenav-size", "condensed");

        const table = document.querySelector('.wide-table');
        const rows = table.querySelectorAll('tbody tr');

        // Column Resizing
        initColumnResizing();
        // restoreColumnWidths();

        // Column Visibility
        setupColumnVisibility();

        // Full Table Search
        setupWholeTableSearch();

        // ✅ Column-Specific Filter (Professional Version)
        setupColumnSearch();

        // Inline Auto-Save
        setupAutoSave();

        // Currency Conversion
        setupCurrencyConversion();

        // Open/Edit O Links
        setupOlinkEditor();

        // Supplier Wise Advance + Pending
        setupSupplierAdvanceCalculation();

        // Total CBM calculation
        calculateTotalCBM();

        //total order qty
        calculateTotalOrderQty();

        //total amount
        calculateTotalAmount();

        // ========= FUNCTIONS ========= //

        function initColumnResizing() {
            const resizers = document.querySelectorAll('.resizer');
            resizers.forEach(resizer => {
                resizer.addEventListener('mousedown', initResize);
            });

            function initResize(e) {
                e.preventDefault();
                const th = e.target.parentElement;
                const startX = e.clientX;
                const startWidth = th.offsetWidth;
                e.target.classList.add('resizing');
                th.style.width = th.style.minWidth = th.style.maxWidth = startWidth + 'px';

                document.addEventListener('mousemove', resize);
                document.addEventListener('mouseup', stopResize);

                function resize(e) {
                    const newWidth = startWidth + e.clientX - startX;
                    if (newWidth > 80) {
                        th.style.width = th.style.minWidth = th.style.maxWidth = newWidth + 'px';

                        if (th.dataset.column === "3") { 
                            const cells = document.querySelectorAll('td[data-column="3"]');
                            cells.forEach(cell => {
                                const shortSpan = cell.querySelector('.sku-short');
                                const fullSpan = cell.querySelector('.sku-full');

                                if (newWidth > 200) { 
                                    shortSpan.classList.add("d-none");
                                    fullSpan.classList.remove("d-none");
                                } else { 
                                    fullSpan.classList.add("d-none");
                                    shortSpan.classList.remove("d-none");
                                }
                            });
                        }
                    }
                }

                function stopResize() {
                    document.removeEventListener('mousemove', resize);
                    document.removeEventListener('mouseup', stopResize);
                    document.querySelectorAll('.resizing').forEach(el => el.classList.remove('resizing'));
                    saveColumnWidths();
                }
            }

            function saveColumnWidths() {
                const widths = {};
                document.querySelectorAll('.wide-table thead th').forEach(th => {
                    const col = th.getAttribute('data-column');
                    widths[col] = th.offsetWidth;
                });
                localStorage.setItem('columnWidths', JSON.stringify(widths));
            }

            function restoreColumnWidths() {
                const widths = JSON.parse(localStorage.getItem('columnWidths') || '{}');
                Object.keys(widths).forEach(col => {
                    const th = document.querySelector(`.wide-table thead th[data-column="${col}"]`);
                    if (th) {
                        th.style.width = th.style.minWidth = th.style.maxWidth = widths[col] + 'px';
                    }
                });
            }
            restoreColumnWidths();
        }

        function setupColumnVisibility() {
            const showAllBtn = document.getElementById('showAllColumns');
            const dropdownBtn = document.getElementById('columnDropdownBtn');
            const dropdownContent = document.getElementById('columnDropdownContent');
            const ths = document.querySelectorAll('.wide-table thead th');

            function capitalizeWords(str) {
                return str.replace(/\w\S*/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
            }

            dropdownBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                dropdownContent.classList.toggle('show');
            });

            window.addEventListener('click', function (e) {
                if (!e.target.matches('.column-dropdown-btn') && !dropdownContent.contains(e.target)) {
                    dropdownContent.classList.remove('show');
                }
            });

            dropdownContent.innerHTML = '';
            ths.forEach((th, i) => {
                const colIndex = i + 1;
                const colName = capitalizeWords(th.childNodes[0].nodeValue.trim());
                const item = document.createElement('div');
                item.className = 'column-checkbox-item';
                const checkbox = document.createElement('input');
                checkbox.type = 'checkbox';
                checkbox.id = `column-${colIndex}`;
                checkbox.className = 'column-checkbox';
                checkbox.setAttribute('data-column', colIndex);
                const label = document.createElement('label');
                label.htmlFor = `column-${colIndex}`;
                label.innerHTML = `${colName} <i class="mdi mdi-eye text-primary"></i>`;
                item.appendChild(checkbox);
                item.appendChild(label);
                dropdownContent.appendChild(item);
            });

            function saveHiddenColumns(hidden) {
                localStorage.setItem('hiddenColumns', JSON.stringify(hidden));
            }

            function getHiddenColumns() {
                return JSON.parse(localStorage.getItem('hiddenColumns') || '[]');
            }

            const hiddenColumns = getHiddenColumns();
            document.querySelectorAll('.column-checkbox').forEach(checkbox => {
                const columnIndex = checkbox.getAttribute('data-column');
                const th = document.querySelector(`.wide-table thead th[data-column="${columnIndex}"]`);
                const label = document.querySelector(`label[for="column-${columnIndex}"]`);
                const colName = capitalizeWords(th.childNodes[0].nodeValue.trim());
                if (hiddenColumns.includes(columnIndex)) {
                    checkbox.checked = false;
                    document.querySelectorAll(`[data-column="${columnIndex}"]`).forEach(cell => cell.style.display = 'none');
                    label.innerHTML = `${colName} <i class="mdi mdi-eye-off text-muted"></i>`;
                } else {
                    checkbox.checked = true;
                    document.querySelectorAll(`[data-column="${columnIndex}"]`).forEach(cell => cell.style.display = '');
                    label.innerHTML = `${colName} <i class="mdi mdi-eye text-primary"></i>`;
                }
            });

            dropdownContent.querySelectorAll('.column-checkbox').forEach(checkbox => {
                checkbox.addEventListener('change', function () {
                    const columnIndex = this.getAttribute('data-column');
                    const th = document.querySelector(`.wide-table thead th[data-column="${columnIndex}"]`);
                    const label = document.querySelector(`label[for="column-${columnIndex}"]`);
                    const colName = capitalizeWords(th.childNodes[0].nodeValue.trim());
                    let hidden = getHiddenColumns();
                    if (this.checked) {
                        document.querySelectorAll(`[data-column="${columnIndex}"]`).forEach(cell => cell.style.display = '');
                        label.innerHTML = `${colName} <i class="mdi mdi-eye text-primary"></i>`;
                        hidden = hidden.filter(c => c !== columnIndex);
                    } else {
                        document.querySelectorAll(`[data-column="${columnIndex}"]`).forEach(cell => cell.style.display = 'none');
                        label.innerHTML = `${colName} <i class="mdi mdi-eye-off text-muted"></i>`;
                        if (!hidden.includes(columnIndex)) hidden.push(columnIndex);
                    }
                    saveHiddenColumns(hidden);
                });
            });

            showAllBtn.addEventListener('click', function () {
                document.querySelectorAll('.column-checkbox').forEach(checkbox => {
                    checkbox.checked = true;
                    const columnIndex = checkbox.getAttribute('data-column');
                    document.querySelectorAll(`[data-column="${columnIndex}"]`).forEach(cell => cell.style.display = '');
                    const th = document.querySelector(`.wide-table thead th[data-column="${columnIndex}"]`);
                    const label = document.querySelector(`label[for="column-${columnIndex}"]`);
                    const colName = capitalizeWords(th.childNodes[0].nodeValue.trim());
                    label.innerHTML = `${colName} <i class="mdi mdi-eye text-primary"></i>`;
                });
                saveHiddenColumns([]);
            });
        }

        function setupWholeTableSearch() {
            const searchInput = document.getElementById('wholeSearchInput');
            searchInput.addEventListener('input', function () {
                const search = this.value.trim().toLowerCase();
                rows.forEach(row => {
                    const match = Array.from(row.querySelectorAll('td')).some(td => td.textContent.toLowerCase().includes(search));
                    row.style.display = match ? '' : 'none';
                });
            });
        }

        // ✅ Fixed version of column search (supports multiple filters)
        function setupColumnSearch() {
            document.querySelectorAll('.column-search').forEach(input => {
                input.addEventListener('input', function () {
                    const filters = {};
                    document.querySelectorAll('.column-search').forEach(searchInput => {
                        const col = searchInput.getAttribute('data-search-column');
                        const val = searchInput.value.trim().toLowerCase();
                        if (val !== '') {
                            filters[col] = val;
                        }
                    });

                    rows.forEach(row => {
                        let show = true;
                        for (const col in filters) {
                            const cell = row.querySelector(`td[data-column="${col}"]`);
                            if (!cell || !cell.textContent.toLowerCase().includes(filters[col])) {
                                show = false;
                                break;
                            }
                        }
                        row.style.display = show ? '' : 'none';
                    });
                });
            });
        }

        function setupAutoSave() {
            document.querySelectorAll('.auto-save').forEach(function (input) {
                input.addEventListener('change', function () {
                    const sku = this.dataset.sku;
                    const column = this.dataset.column;
                    const value = this.value;
                    const row = this.closest('tr');

                    if (!sku || !column) return;
                    
                    // if (column === 'ready_to_ship' && value === 'Yes') {
                    //     const photoPacking = row.querySelector('td[data-column="15"] a')?.href?.trim() || '';
                    //     const photoIntSale = row.querySelector('td[data-column="16"] a')?.href?.trim() || '';
                    //     const barcodeSku = row.querySelector('td[data-column="19"] a')?.href?.trim() || '';
                    //     const artwork = row.querySelector('td[data-column="20"] input')?.value?.trim() || '';
                    //     const notes = row.querySelector('td[data-column="21"] input')?.value?.trim() || '';

                    //     if (!photoPacking || !photoIntSale || !barcodeSku || !artwork || !notes) {
                    //         alert("❌ Please fill all fields before marking 'Ready to Ship':\n- Photo Packing\n- Photo Internal Sale\n- Barcode SKU\n- Artwork Manual Book\n- Notes");
                    //         this.value = 'No';
                    //         return;
                    //     }
                    // }

                    // ✅ Save via AJAX
                    fetch('/mfrg-progresses/inline-update-by-sku', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'X-CSRF-TOKEN': '{{ csrf_token() }}'
                        },
                        body: JSON.stringify({ sku, column, value })
                    })
                    .then(res => res.json())
                    .then(res => {
                        if (res.success) {
                            this.style.border = '2px solid green';
                            setTimeout(() => this.style.border = '', 1000);

                            // ✅ Recalculate Total on rate change
                            if (column === 'rate') {
                                const qtyCell = row.querySelector('td[data-column="4"]');
                                const totalCell = row.querySelector('td[data-column="12"]');
                                const qty = parseFloat(qtyCell?.innerText?.trim() || '0');
                                const rate = parseFloat(value);
                                if (!isNaN(qty) && !isNaN(rate)) {
                                    totalCell.innerText = (qty * rate).toFixed(2);
                                }
                            }

                            // ✅ Recalculate Pending on advance_amt change
                            if (column === 'advance_amt') {
                                const totalCell = row.querySelector('td[data-column="12"]');
                                const pendingCell = row.querySelector('td[data-column="13"]');
                                const total = parseFloat(totalCell?.innerText?.trim() || '0');
                                const advance = parseFloat(value);
                                if (!isNaN(total) && !isNaN(advance)) {
                                    pendingCell.innerText = (total - advance).toFixed(2);
                                }
                            }

                            // ✅ Insert into Ready to Ship table
                            if (column === 'ready_to_ship' && value === 'Yes') {
                                const parent = row.querySelector('td:nth-child(2)')?.innerText?.trim() || '';
                                const skuVal = row.querySelector('#sku-full')?.innerText?.trim() || '';
                                const supplier = row.querySelector('td:nth-child(6)')?.innerText?.trim() || '';
                                const totalCbm = row.querySelector('td[data-column="15"] input')?.value?.trim() || '';

                                fetch('/ready-to-ship/insert', {
                                    method: 'POST',
                                    headers: {
                                        'Content-Type': 'application/json',
                                        'X-CSRF-TOKEN': '{{ csrf_token() }}'
                                    },
                                    body: JSON.stringify({
                                        parent: parent,
                                        sku: skuVal,
                                        supplier: supplier,
                                        totalCbm: totalCbm
                                    })
                                })
                                .then(r => r.json())
                                .then(r => {
                                    if (r.success) {
                                        row.remove();
                                    } else {
                                        alert('❌ Failed to insert into Ready to Ship: ' + r.message);
                                    }
                                })
                                .catch(() => {
                                    alert('❌ Error during Ready to Ship insert.');
                                });
                            }

                        } else {
                            this.style.border = '2px solid red';
                            console.log('❌ Error:', res.message);
                        }
                    })
                    .catch(() => {
                        this.style.border = '2px solid red';
                        alert('❌ AJAX error occurred.');
                    });
                });
            });
        }

        function setupAutoUpload() {
            document.querySelectorAll('.auto-upload').forEach(function(input) {
                input.addEventListener('change', function () {
                    const sku = this.dataset.sku;
                    const column = this.dataset.column;
                    const file = this.files[0];
                    const parentDiv = input.closest('.image-upload-field');

                    if (!sku || !column || !file) return;

                    const formData = new FormData();
                    formData.append('sku', sku);
                    formData.append('column', column);
                    formData.append('value', file); 

                    fetch('/mfrg-progresses/inline-update-by-sku', {
                        method: 'POST',
                        headers: {
                            'X-CSRF-TOKEN': '{{ csrf_token() }}'
                        },
                        body: formData
                    })
                    .then(res => res.json())
                    .then(res => {
                        if (res.success && res.url) {
                            // ✅ Update with image preview + upload field again
                            parentDiv.innerHTML = `
                                <a href="${res.url}" target="_blank" style="width:50px;">
                                    <img src="${res.url}" alt="Uploaded" style="height: 50px; width: 50px; object-fit: cover; background: #f8f9fa;">
                                </a>
                                <label class="btn btn-sm btn-outline-primary d-flex align-items-center mb-0" style="padding: 4px 10px; border-radius: 6px; font-size: 15px;" title="Upload Photo">
                                    <i class="mdi mdi-upload" style="font-size: 18px; margin-right: 4px;"></i>
                                    <span style="font-size: 13px;"></span>
                                    <input type="file" class="d-none auto-upload" data-column="${column}" data-sku="${sku}">
                                </label>
                            `;
                            setupAutoUpload(); // re-bind to new input
                        } else {
                            alert("❌ Upload failed: " + (res.message || 'Unknown error'));
                        }
                    })
                    .catch(() => {
                        alert("❌ AJAX error during upload");
                    });
                });
            });
        }

        setupAutoUpload();

        function setupCurrencyConversion() {
            document.querySelectorAll('.input-group').forEach(group => {
                const select = group.querySelector('.currency-select');
                const input = group.querySelector('.amount-input');

                if (!select || !input) return;

                let baseCurrency = select.value;
                let baseAmount = parseFloat(input.value) || 0;

                input.addEventListener('input', () => {
                    baseAmount = parseFloat(input.value) || 0;
                    baseCurrency = select.value;
                });

                select.addEventListener('change', function () {
                    const newCurrency = select.value;

                    if (baseCurrency === newCurrency || isNaN(baseAmount) || baseAmount === 0) return;

                    const url = `/convert-currency?amount=${baseAmount}&from=${baseCurrency}&to=${newCurrency}`;

                    fetch(url)
                        .then(res => res.json())
                        .then(data => {
                            console.log('Currency API response:', data);
                            if (data.rates && data.rates[newCurrency]) {
                                const converted = data.rates[newCurrency];
                                input.value = parseFloat(converted).toFixed(2);
                                baseAmount = parseFloat(converted);
                                baseCurrency = newCurrency;
                            } else {
                                alert('Invalid currency response (missing rates).');
                            }
                        })
                        .catch(err => {
                            alert('Currency conversion failed.');
                            console.error(err);
                        });
                });
            });
        }

        function setupOlinkEditor() {
            document.querySelectorAll('.edit-o-links').forEach(function (btn) {
                btn.addEventListener('click', function () {
                    const group = btn.closest('.input-group');
                    const input = group.querySelector('.o-links-input');
                    const openIcon = group.querySelector('.open-link-icon');
                    const editIcon = group.querySelector('.edit-link-icon');

                    input.classList.remove('d-none');
                    input.focus();
                    openIcon.style.display = 'none';
                    editIcon.style.display = 'none';

                    input.addEventListener('blur', function () {
                        input.classList.add('d-none');
                        openIcon.style.display = 'inline-flex';
                        editIcon.style.display = 'inline-flex';
                    });

                    input.addEventListener('keydown', function (e) {
                        if (e.key === 'Enter') {
                            input.blur();
                        }
                    });
                });
            });
        }

        function setupSupplierAdvanceCalculation() {
            const selectBox = document.getElementById('customSelectBox');
            const dropdown = document.getElementById('customSelectDropdown');
            const selectedText = document.getElementById('customSelectSelectedText');
            const searchInput = document.getElementById('customSelectSearchInput');
            const optionsContainer = document.getElementById('customSelectOptions');
            const wrapper = document.getElementById('advance-total-wrapper');

            let allOptions = Array.from(optionsContainer.querySelectorAll('.custom-select-option'));

            selectBox.addEventListener('click', function () {
                dropdown.style.display = dropdown.style.display === 'block' ? 'none' : 'block';
                selectBox.classList.toggle('active', dropdown.style.display === 'block');
                searchInput.value = '';
                allOptions.forEach(option => option.style.display = '');
                setTimeout(() => searchInput.focus(), 100);
            });

            optionsContainer.addEventListener('click', function (e) {
                if (!e.target.classList.contains('custom-select-option')) return;

                // UI update
                allOptions.forEach(opt => opt.classList.remove('selected', 'bg-primary', 'text-white'));
                e.target.classList.add('selected', 'bg-primary', 'text-white');
                selectedText.textContent = e.target.textContent;
                dropdown.style.display = 'none';
                selectBox.classList.remove('active');

                const selectedSupplier = e.target.textContent.trim();
                const allRows = document.querySelectorAll('tbody tr');

                if (!selectedSupplier || selectedSupplier === 'Select supplier') {
                    document.getElementById('advance-total-wrapper').style.display = 'none';
                    return;
                }

                // Reset visibility
                allRows.forEach(row => row.style.display = '');

                // Filter rows matching selected supplier
                let matchingRows = [];
                allRows.forEach(row => {
                    const supplierCell = row.querySelector('td[data-column="6"]');
                    if (supplierCell && supplierCell.textContent.trim().toLowerCase() === selectedSupplier.toLowerCase()) {
                        row.style.display = '';
                        matchingRows.push(row);
                    } else {
                        row.style.display = 'none';
                    }
                });

                if (!wrapper) {
                    console.warn("Wrapper not found! Skipping animation and totals.");
                    return;
                }

                if (matchingRows.length === 0) {
                    wrapper.style.display = 'none';
                    console.log("No matching rows, hiding wrapper.");
                    return;
                } else {
                    wrapper.style.display = 'block';
                }

                // Calculate total group value
                let totalGroupValue = 0;
                matchingRows.forEach(row => {
                    const qty = parseFloat(row.querySelector('td[data-column="4"]')?.innerText || '0') || 0;
                    const rate = parseFloat(row.querySelector('input[data-column="rate"]')?.value || '0') || 0;
                    totalGroupValue += qty * rate;
                });

                // Calculate total advance
                let totalAdvance = 0;
                matchingRows.forEach(row => {
                    const input = row.querySelector('input[data-supplier]');
                    if (input && !input.disabled) {
                        totalAdvance += parseFloat(input.value || '0') || 0;
                    }
                });

                let totalPending = 0;
                let supplierTotalCBM = 0;
                let supplierTotalOrderQty = 0;
                let supplierTotalAmount = 0;

                matchingRows.forEach(row => {
                    const qty = parseFloat(row.querySelector('td[data-column="4"]')?.innerText || '0') || 0;
                    const rate = parseFloat(row.querySelector('input[data-column="rate"]')?.value || '0') || 0;
                    const rowTotal = qty * rate;

                    const advanceInput = row.querySelector('input[data-column="advance"]');
                    const pendingInput = row.querySelector('input[data-column="pending"]');

                    const cbmInput = row.querySelector('input[data-column="total_cbm"]');
                    const orderQtyCell = row.querySelector('td[data-column="4"]');

                    const cbm = parseFloat(cbmInput?.value || '0');

                    if (!isNaN(cbm)) supplierTotalCBM += cbm;
                    if (!isNaN(qty)) supplierTotalOrderQty += qty;
                    supplierTotalAmount += rowTotal;

                    let rowAdvance = 0;
                    if (totalGroupValue > 0 && rowTotal > 0) {
                        rowAdvance = (rowTotal / totalGroupValue) * totalAdvance;
                    }

                    const rowPending = rowTotal - rowAdvance;
                    totalPending += rowPending;

                    if (advanceInput) advanceInput.value = rowAdvance.toFixed(2);
                    if (pendingInput) pendingInput.value = rowPending.toFixed(2);
                });

                // Update summary display
                document.getElementById('advance-amount').textContent = totalAdvance.toFixed(2);
                document.getElementById('pending-amount').textContent = totalPending.toFixed(2);

                // Update supplier cbm and order qty
                document.getElementById('total-cbm').textContent = supplierTotalCBM.toFixed(2);
                document.getElementById('total-order-qty').textContent = supplierTotalOrderQty;
                document.getElementById('total-amount').textContent = supplierTotalAmount;

                // Animate wrapper
                wrapper.classList.add('animate__animated', 'animate__fadeIn');
                setTimeout(() => wrapper.classList.remove('animate__animated', 'animate__fadeIn'), 800);
            });


            // 🔍 Search filter
            searchInput.addEventListener('input', function () {
                const search = this.value.trim().toLowerCase();
                allOptions.forEach(option => {
                    option.style.display = option.textContent.toLowerCase().includes(search) ? '' : 'none';
                });
            });

            // ⌨️ Keyboard navigation
            searchInput.addEventListener('keydown', function (e) {
                let visibleOptions = allOptions.filter(opt => opt.style.display !== 'none');
                let selectedIdx = visibleOptions.findIndex(opt => opt.classList.contains('selected'));
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    if (selectedIdx < visibleOptions.length - 1) {
                        if (selectedIdx >= 0) visibleOptions[selectedIdx].classList.remove('selected', 'bg-primary', 'text-white');
                        visibleOptions[selectedIdx + 1].classList.add('selected', 'bg-primary', 'text-white');
                        visibleOptions[selectedIdx + 1].scrollIntoView({ block: 'nearest' });
                    }
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    if (selectedIdx > 0) {
                        visibleOptions[selectedIdx].classList.remove('selected', 'bg-primary', 'text-white');
                        visibleOptions[selectedIdx - 1].classList.add('selected', 'bg-primary', 'text-white');
                        visibleOptions[selectedIdx - 1].scrollIntoView({ block: 'nearest' });
                    }
                } else if (e.key === 'Enter') {
                    if (selectedIdx >= 0) {
                        visibleOptions[selectedIdx].click();
                    }
                }
            });

            // 🖱️ Close on outside click
            document.addEventListener('mousedown', function (e) {
                if (!selectBox.contains(e.target) && !dropdown.contains(e.target)) {
                    dropdown.style.display = 'none';
                    selectBox.classList.remove('active');
                }
            });
        }

        document.querySelectorAll('.sku-short').forEach(el => {
            el.addEventListener('click', function () {
                const shortSpan = this;
                const fullSpan = this.nextElementSibling;

                shortSpan.classList.add("d-none");
                fullSpan.classList.remove("d-none");

                fullSpan.addEventListener('click', function () {
                    fullSpan.classList.add("d-none");
                    shortSpan.classList.remove("d-none");
                }, { once: true });
            });
        });

        // document.getElementById('play-auto').addEventListener('click', () => {
        //     isPlaying = true;
        //     currentSupplierIndex = 0;
        //     renderGroup(supplierKeys[currentSupplierIndex]);
        //     document.getElementById('play-pause').style.display = 'inline-block';
        //     document.getElementById('play-auto').style.display = 'none';
        // });

        // document.getElementById('play-forward').addEventListener('click', () => {
        //     if (!isPlaying) return;
        //     currentSupplierIndex = (currentSupplierIndex + 1) % supplierKeys.length;
        //     renderGroup(supplierKeys[currentSupplierIndex]);
        // });

        // document.getElementById('play-backward').addEventListener('click', () => {
        //     if (!isPlaying) return;
        //     currentSupplierIndex = (currentSupplierIndex - 1 + supplierKeys.length) % supplierKeys.length;
        //     renderGroup(supplierKeys[currentSupplierIndex]);
        // });

        // document.getElementById('play-pause').addEventListener('click', () => {
        //     isPlaying = false;
        //     tableBody.innerHTML = originalTableHtml;
        //     document.getElementById('play-pause').style.display = 'none';
        //     document.getElementById('play-auto').style.display = 'inline-block';
        //     attachEditableListeners();
        //     attachStageListeners();
        // });

    });
</script>
<script>
    document.addEventListener("DOMContentLoaded", function () {
        const rows = document.querySelectorAll("table.wide-table tbody tr");
        const filterSelect = document.getElementById("row-data-pending-status");
        const greenSpan = document.getElementById("greenCount");
        const yellowSpan = document.getElementById("yellowCount");
        const redSpan = document.getElementById("redCount");

        const suppliers = [];
        let supplierIndex = 0;
        let intervalId = null;

        // Collect unique suppliers
        rows.forEach(row => {
            const supplierCell = row.querySelector('td[data-column="6"]');
            if (supplierCell) {
                const supplierName = supplierCell.textContent.trim();
                if (supplierName && !suppliers.includes(supplierName)) {
                    suppliers.push(supplierName);
                }
            }
        });

        function showSupplierRows(supplier) {
            rows.forEach(row => {
                const cell = row.querySelector('td[data-column="6"]');
                if (cell && cell.textContent.trim() === supplier) {
                    row.style.display = "";
                } else {
                    row.style.display = "none";
                }
            });

            // Optionally show current supplier name
            const title = document.getElementById("current-supplier");
            if (title) title.textContent = "Supplier: " + supplier;

            calculateTotalCBM();
            calculateTotalAmount();
            calculateTotalOrderQty();
        }

        function playNextSupplier() {
            supplierIndex = (supplierIndex + 1) % suppliers.length;
            showSupplierRows(suppliers[supplierIndex]);
        }

        document.getElementById("play-auto").addEventListener("click", function () {
            this.style.display = "none";
            document.getElementById("play-pause").style.display = "inline-block";
            showSupplierRows(suppliers[supplierIndex]);
        });

        document.getElementById("play-pause").addEventListener("click", function () {
            this.style.display = "none";
            document.getElementById("play-auto").style.display = "inline-block";
            rows.forEach(row => row.style.display = "");
                const title = document.getElementById("current-supplier");
            if (title) title.textContent = "";
            calculateTotalCBM();
            calculateTotalAmount();
            calculateTotalOrderQty();
        });


        document.getElementById("play-forward").addEventListener("click", function () {
            
            playNextSupplier();
        });

        document.getElementById("play-backward").addEventListener("click", function () {
            
            supplierIndex = (supplierIndex - 1 + suppliers.length) % suppliers.length;
            showSupplierRows(suppliers[supplierIndex]);
        });

        function updateCounts() {
            let green = 0, yellow = 0, red = 0;

            rows.forEach(row => {
                const dateInput = row.querySelector('input[data-column="del_date"]');
                if (!dateInput) return;

                const bg = dateInput.style.backgroundColor.trim().toLowerCase();
                if (bg === "green") green++;
                else if (bg === "yellow") yellow++;
                else if (bg === "red") red++;
            });

            greenSpan.textContent = `(${green})`;
            yellowSpan.textContent = `(${yellow})`;
            redSpan.textContent = `(${red})`;
        }

        function filterDateRows(type) {
            rows.forEach(row => {
                const dateInput = row.querySelector('input[data-column="del_date"]');
                if (!dateInput) return;

                const bg = dateInput.style.backgroundColor.trim().toLowerCase();

                row.style.display = (!type || bg === type) ? "" : "none";
            });

            calculateTotalCBM();
            calculateTotalAmount();
            calculateTotalOrderQty();
        }

        updateCounts();

        filterSelect.addEventListener("change", function () {
            filterDateRows(this.value);
        });
    });

    function calculateTotalCBM() {
        let totalCBM = 0;

        document.querySelectorAll('table.wide-table tbody tr').forEach(row => {
            if (row.style.display !== "none") {
                const input = row.querySelector('input[data-column="total_cbm"]');
                if (input) {
                    const value = parseFloat(input.value);
                    if (!isNaN(value)) totalCBM += value;
                }
            }
        });

        document.getElementById('total-cbm').textContent = totalCBM.toFixed(0);
    } 

    function calculateTotalAmount() {
        let totalAmount = 0;

        document.querySelectorAll('table.wide-table tbody tr').forEach(row => {
            if (row.style.display !== "none") { 
                const td = row.querySelector('.total-value');
                if (td) {
                    const value = parseFloat(td.textContent.trim());
                    if (!isNaN(value)) {
                        totalAmount += value;
                    }
                }
            }
        });

        document.getElementById('total-amount').textContent = totalAmount.toFixed(0);
    }

    function calculateTotalOrderQty() {
        let totalOrderQty = 0;
        document.querySelectorAll('table.wide-table tbody tr').forEach(row => {
            if (row.style.display !== "none") { 
                const cell = row.querySelector('[data-column="4"]');
                if (cell) {
                    const value = parseFloat(cell.textContent.trim());
                    if (!isNaN(value)) {
                        totalOrderQty += value;
                    }
                }
            }
        });

        document.getElementById('total-order-qty').textContent = totalOrderQty;
    }

</script>

@endsection