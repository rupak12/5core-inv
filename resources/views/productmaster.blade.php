@extends('layouts.vertical', ['title' => 'CP Master', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])

@section('css')
    @vite(['node_modules/admin-resources/rwd-table/rwd-table.min.css'])
    <link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.3.6/css/buttons.dataTables.min.css">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css" rel="stylesheet">

    <style>
        /* Your existing CSS styles remain unchanged */
        .table-responsive {
            position: relative;
            border: 1px solid #e9ecef;
            border-radius: 10px;
            max-height: 600px;
            overflow-y: auto;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
            background-color: white;
        }

        .table-responsive thead th {
            position: sticky;
            top: 0;
            background: linear-gradient(135deg, #2c6ed5 0%, #1a56b7 100%) !important;
            color: white;
            z-index: 10;
            padding: 15px 18px;
            font-weight: 600;
            border-bottom: none;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
            font-size: 13px;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            transition: all 0.2s ease;
        }

        .table-responsive thead th:hover {
            background: linear-gradient(135deg, #1a56b7 0%, #0a3d8f 100%) !important;
        }

        .table-responsive thead input {
            background-color: rgba(255, 255, 255, 0.9);
            border: none;
            border-radius: 4px;
            color: #333;
            padding: 6px 10px;
            margin-top: 8px;
            font-size: 12px;
            width: 100%;
            transition: all 0.2s;
        }

        .table-responsive thead input:focus {
            background-color: white;
            box-shadow: 0 0 0 2px rgba(26, 86, 183, 0.3);
            outline: none;
        }

        .table-responsive thead input::placeholder {
            color: #8e9ab4;
            font-style: italic;
        }

        .table-responsive tbody td {
            padding: 12px 18px;
            vertical-align: middle;
            border-bottom: 1px solid #edf2f9;
            font-size: 13px;
            color: #495057;
            transition: all 0.2s ease;
        }

        .table-responsive tbody tr:nth-child(even) {
            background-color: #f8fafc;
        }

        .table-responsive tbody tr:hover {
            background-color: #e8f0fe;
            transform: translateY(-1px);
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);
        }

        .table-responsive tbody tr:hover td {
            color: #000;
        }

        .table-responsive .text-center {
            text-align: center;
        }

        .table {
            margin-bottom: 0;
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
        }

        .edit-btn {
            border-radius: 6px;
            padding: 6px 12px;
            transition: all 0.2s;
            background: #fff;
            border: 1px solid #1a56b7;
            color: #1a56b7;
        }

        .edit-btn:hover {
            background: #1a56b7;
            color: white;
            transform: translateY(-2px);
            box-shadow: 0 3px 8px rgba(26, 86, 183, 0.2);
        }

        /* Additional styles remain unchanged */
        .form-control {
            border: 2px solid #E2E8F0;
            border-radius: 6px;
            padding: 0.75rem;
        }

        #status {
            display: block !important;
            position: static !important;
            width: 100% !important;
            height: auto !important;
            margin: auto !important;
        }

        .dt-buttons .btn {
            margin-left: 10px;
        }

        .modal-header-gradient {
            background: linear-gradient(135deg, #6B73FF 0%, #000DFF 100%);
            border-bottom: 4px solid #4D55E6;
            padding: 1.5rem;
        }

        .modal-footer-gradient {
            background: linear-gradient(135deg, #F8FAFF 0%, #E6F0FF 100%);
            border-top: 4px solid #E2E8F0;
            padding: 1.5rem;
        }

        .rainbow-loader {
            display: none;
            text-align: center;
            padding: 20px;
        }

        .loading-text {
            margin-top: 10px;
            font-weight: bold;
        }

        /* Add these new styles for the plus button and multi-select functionality */
        .selection-controls {
            position: relative;
            z-index: 100;
            display: inline-flex;
            align-items: center;
            opacity: 0;
            transition: opacity 0.3s;
        }

        .card-body:hover .selection-controls {
            opacity: 1;
        }

        .select-toggle-btn {
            background: linear-gradient(135deg, #4ADE80 0%, #22C55E 100%);
            color: white;
            border: none;
            border-radius: 50%;
            width: 36px;
            height: 36px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 3px 10px rgba(34, 197, 94, 0.3);
            transition: all 0.2s;
        }

        .select-toggle-btn:hover {
            transform: scale(1.1);
            box-shadow: 0 5px 15px rgba(34, 197, 94, 0.4);
        }

        .select-toggle-text {
            margin-right: 10px;
            font-weight: 600;
            color: #333;
        }

        .checkbox-column {
            width: 40px;
            text-align: center;
            display: none;
        }

        .select-all-checkbox {
            width: 18px;
            height: 18px;
            cursor: pointer;
        }

        .row-checkbox {
            width: 16px;
            height: 16px;
            cursor: pointer;
        }

        .selection-actions {
            display: none;
            position: fixed;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 1050;
            background: linear-gradient(135deg, #2c6ed5 0%, #1a56b7 100%);
            padding: 10px 20px;
            border-radius: 50px;
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
        }

        .selection-actions .btn {
            margin: 0 5px;
            border-radius: 20px;
            font-weight: 600;
            padding: 5px 15px;
        }

        .selection-count {
            color: white;
            font-weight: bold;
            margin-right: 15px;
            display: inline-block;
        }

        /* Add these styles to your CSS section */
        .field-operation {
            padding: 10px;
            border-radius: 6px;
            background-color: #f8f9fa;
            transition: all 0.2s;
        }

        .field-operation:hover {
            background-color: #e9ecef;
        }

        #addFieldBtn {
            border-radius: 20px;
            padding: 6px 15px;
        }

        #applyChangesBtn {
            background: linear-gradient(135deg, #2c6ed5 0%, #1a56b7 100%);
            border: none;
        }

        .remove-field {
            width: 36px;
            height: 36px;
            padding: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 4px;
        }

        .custom-toast {
            z-index: 2000;
        }

        /* Add to your CSS section */
        .sku-tooltip {
            position: absolute;
            z-index: 9999;
            background: #fff;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 6px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
            display: none;
            max-width: 180px;
            max-height: 180px;
        }

        .sku-tooltip img {
            max-width: 160px;
            max-height: 160px;
            border-radius: 6px;
            display: block;
        }

        /* Add to your <style> section */
        .custom-dropdown {
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            z-index: 1000;
            background: #fff;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
            max-height: 180px;
            overflow-y: auto;
            display: none;
        }

        .custom-dropdown .dropdown-item {
            padding: 8px 14px;
            cursor: pointer;
            font-size: 14px;
            color: #333;
        }

        .custom-dropdown .dropdown-item:hover {
            background: #e8f0fe;
            color: #1a56b7;
        }
    </style>
@endsection

@section('content')
    <meta name="csrf-token" content="{{ csrf_token() }}">

    @include('layouts.shared/page-title', [
        'page_title' => 'CP Master',
        'sub_title' => 'Product master Analysis',
    ])

    <div class="row">
        <div class="col-12">
            <div class="card">
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-6">
                            <div class="d-flex align-items-center gap-3">
                                <!-- Keep selection-controls unchanged -->
                                <div class="selection-controls" 
                                    style="position: relative; opacity: 1; display: inline-flex; margin-right: 15px;">
                                    <span class="select-toggle-text">Multi Add</span>
                                    <button type="button" class="select-toggle-btn" id="toggleSelection">
                                        <i class="fas fa-plus"></i>
                                    </button>
                                </div>
                                <!-- Compact field selector -->
                                <div class="field-selector-wrapper" style="width: 200px;">
                                    <select class="form-select form-select-md" style="border-radius: 6px; border: 1px solid #92c1ff; font-size: 13px;">
                                        <option value="">Filter By...</option>
                                        <option value="lp">LP</option>
                                        <option value="cp">CP</option>
                                        <option value="frght">FRGHT</option>
                                        <option value="ship">SHIP</option>
                                        <option value="temu_ship">TEMU SHIP</option>
                                        <option value="ebay2_ship">EBAY2 SHIP</option>
                                        <option value="initial_quantity">INITIAL QTY</option>
                                        <option value="label_qty">Label QTY</option>
                                        <option value="wt_act">WT ACT</option>
                                        <option value="wt_decl">WT DECL</option>
                                        <option value="l">L</option>
                                        <option value="w">W</option>
                                        <option value="h">H</option>
                                        <option value="status">Status</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-6 text-end">
                            <div class="input-group mb-3">
                                <span class="input-group-text"><i class="fas fa-search"></i></span>
                                <input type="text" id="customSearch" class="form-control"
                                    placeholder="Search products...">
                                <button class="btn btn-outline-secondary" type="button" id="clearSearch">Clear</button>

                                <!-- @php
                                    $isAdmin =
                                        auth()->check() && in_array(auth()->user()->role, ['admin', 'superadmin']);
                                    @endphp -->

                                <!-- @if ($isAdmin) -->
                                    <button class="btn btn-outline-primary ms-2" type="button" id="culomnPermissionBtn">
                                        <i class="fas"></i> Permission
                                    </button>
                                <!-- @endif -->
                            </div>
                        </div>
                    </div>

                    <div class="col-md-6 text-end mb-3">
                            <button type="button" class="btn btn-primary" data-bs-toggle="modal"
                                data-bs-target="#addProductModal">
                                <i class="fas fa-plus me-1"></i> ADD PRODUCT
                            </button>

                        <button type="button" class="btn btn-success ms-2" id="downloadExcel">
                            <i class="fas fa-file-excel me-1"></i> Download Excel
                        </button>
                        <button type="button" class="btn btn-warning ms-2" id="importFromApiBtn" hidden>
                            <i class="fas fa-cloud-download-alt me-1"></i> Import from API Sheet
                        </button>
                    </div>

                    <!-- Permission Modal -->
                    <div class="modal fade" id="permissionModal" tabindex="-1" aria-labelledby="permissionModalLabel"
                        aria-hidden="true">
                        <div class="modal-dialog">
                            <div class="modal-content" style="background: #fff;">
                                <div class="modal-header modal-header-gradient">
                                    <h5 class="modal-title" id="permissionModalLabel" style="color: white;">Column
                                        Permission</h5>
                                    <button type="button" class="btn-close" data-bs-dismiss="modal"
                                        aria-label="Close"></button>
                                </div>
                                <div class="modal-body">
                                    <!-- Email Dropdown -->
                                    <div class="mb-3" style="position:relative;">
                                        <label for="emailInput" class="form-label">Select Email</label>
                                        <input type="text" id="emailInput" class="form-control" autocomplete="off"
                                            placeholder="Type to search email...">
                                        <div id="emailDropdownList" class="custom-dropdown"></div>
                                    </div>
                                    <!-- Column Dropdown & Add -->
                                    <div class="mb-3" style="position:relative;">
                                        <label for="columnInput" class="form-label">Add Column</label>
                                        <input type="text" id="columnInput" class="form-control" autocomplete="off"
                                            placeholder="Type to search column...">
                                        <div id="columnDropdownList" class="custom-dropdown"></div>
                                        <button type="button" class="btn btn-sm btn-success mt-2" id="addColumnBtn">Add
                                            Column</button>
                                    </div>
                                    <!-- Current Columns -->
                                    <div class="mb-3">
                                        <label class="form-label">Current Columns</label>
                                        <div id="currentColumns"></div>
                                    </div>
                                    <p class="mt-3">Set column permissions here.</p>
                                </div>
                                <div class="modal-footer modal-footer-gradient">
                                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                                    <button type="button" class="btn btn-primary" id="savePermissionBtn">Save
                                        changes</button>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Add Product Modal -->
                    <div class="modal fade" id="addProductModal" tabindex="-1" aria-labelledby="addProductModalLabel"
                        aria-hidden="true">
                        <div class="modal-dialog modal-lg">
                            <div class="modal-content"
                                style="border: none; border-radius: 0; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
                                <!-- Modal Header -->
                                <div class="modal-header"
                                    style="background: linear-gradient(135deg, #6B73FF 0%, #000DFF 100%); border-bottom: 4px solid #4D55E6; padding: 1.5rem; border-radius: 0;">
                                    <h5 class="modal-title" id="addProductModalLabel"
                                        style="color: white; font-weight: 800; font-size: 1.8rem; letter-spacing: 0.5px;">
                                        <i class="fas fa-plus-circle me-2"></i>ADD NEW PRODUCT LISTING
                                    </h5>
                                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"
                                        aria-label="Close"></button>
                                </div>

                                <!-- Modal Body -->
                                <div class="modal-body" style="background-color: #F8FAFF; padding: 2rem;">
                                    <div id="form-errors" class="mb-3"></div>
                                    <form id="addProductForm">
                                        <!-- Row 1 -->
                                        <div class="row mb-5">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="sku" class="form-label fw-bold"
                                                        style="color: #4A5568;">SKU*</label>
                                                    <input type="text" class="form-control" id="sku"
                                                        placeholder="Enter SKU"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="parent" class="form-label fw-bold"
                                                        style="color: #4A5568;">Parent</label>
                                                    <div class="input-group">
                                                        <input type="text" class="form-control" id="parent"
                                                            placeholder="Enter or select parent"
                                                            style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;"
                                                            list="parentOptions">
                                                        <datalist id="parentOptions"></datalist>
                                                        <button class="btn btn-outline-secondary" type="button"
                                                            id="refreshParents" style="border-radius: 0 6px 6px 0;">
                                                            <i class="fas fa-sync-alt"></i>
                                                        </button>
                                                    </div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="labelQty" class="form-label fw-bold"
                                                        style="color: #4A5568;">Label QTY*</label>
                                                    <input type="text" class="form-control" id="labelQty"
                                                        placeholder="Enter QTY"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="status" class="form-label fw-bold"
                                                        style="color: #4A5568;">Status</label>
                                                    <select class="form-control" id="status" name="status"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                        <option value="">Select Status</option>
                                                        <option value="active">Active</option>
                                                        <option value="inactive">Inactive</option>
                                                        <option value="DC">DC</option>
                                                        <option value="upcoming">Upcoming</option>
                                                        <option value="2BDC">2BDC</option>
                                                    </select>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="unit" class="form-label fw-bold"
                                                        style="color: #4A5568;">Unit</label>
                                                    <select class="form-control" id="unit" name="unit" required>
                                                        <option value="Pieces">Pieces</option>
                                                        <option value="Pair">Pair</option>
                                                    </select>
                                                </div>
                                            </div>
                                        </div>

                                        <!-- Row 2 -->
                                        <div class="row mb-4">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="cp" class="form-label fw-bold"
                                                        style="color: #4A5568;">CP</label>
                                                    <input type="text" class="form-control" id="cp"
                                                        placeholder="Enter cp"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="lp" class="form-label fw-bold"
                                                        style="color: #4A5568;">LP</label>
                                                    <input type="text" class="form-control" id="lp"
                                                        placeholder="Enter LP"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: #EDF2F7;"
                                                        readonly>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="freght" class="form-label fw-bold"
                                                        style="color: #4A5568;">FRGHT</label>
                                                    <input type="text" class="form-control" id="freght"
                                                        placeholder="Enter FRIGHT"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: #EDF2F7;"
                                                        readonly>
                                                </div>
                                            </div>
                                            <div class="col-md-3" hidden>
                                                <div class="form-group">
                                                    <label for="lps" class="form-label fw-bold"
                                                        style="color: #4A5568;">LPS</label>
                                                    <input type="text" class="form-control" id="lps"
                                                        placeholder="Enter LPS"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: #EDF2F7;"
                                                        readonly>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="wtAct" class="form-label fw-bold"
                                                        style="color: #4A5568;">WT ACT*</label>
                                                    <input type="text" class="form-control" id="wtAct"
                                                        placeholder="Enter WT ACT"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                        </div>

                                        <!-- Row 3 -->
                                        <div class="row mb-4">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="wtDecl" class="form-label fw-bold"
                                                        style="color: #4A5568;">WT DECL*</label>
                                                    <input type="text" class="form-control" id="wtDecl"
                                                        placeholder="Enter WT DECL"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="ship" class="form-label fw-bold"
                                                        style="color: #4A5568;">SHIP*</label>
                                                    <input type="text" class="form-control" id="ship"
                                                        placeholder="Enter SHIP"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="temu_ship" class="form-label fw-bold"
                                                        style="color: #4A5568;">TEMU SHIP</label>
                                                    <input type="text" class="form-control" id="temu_ship"
                                                        placeholder="Enter TEMU SHIP"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="ebay2_ship" class="form-label fw-bold"
                                                        style="color: #4A5568;">EBAY2 SHIP</label>
                                                    <input type="text" class="form-control" id="ebay2_ship"
                                                        placeholder="Enter EBAY2 SHIP"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="initial_quantity" class="form-label fw-bold"
                                                        style="color: #4A5568;">INITIAL QUANTITY</label>
                                                    <input type="text" class="form-control" id="initial_quantity"
                                                        placeholder="Enter INITIAL QUANTITY"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="w" class="form-label fw-bold"
                                                        style="color: #4A5568;">W*</label>
                                                    <input type="text" class="form-control" id="w"
                                                        placeholder="Enter W"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="l" class="form-label fw-bold"
                                                        style="color: #4A5568;">L*</label>
                                                    <input type="text" class="form-control" id="l"
                                                        placeholder="Enter L"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                        </div>

                                        <!-- Row 4 -->
                                        <div class="row mb-4">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="h" class="form-label fw-bold"
                                                        style="color: #4A5568;">H*</label>
                                                    <input type="text" class="form-control" id="h"
                                                        placeholder="Enter H"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="cbm" class="form-label fw-bold"
                                                        style="color: #4A5568;">CBM</label>
                                                    <input type="text" class="form-control" id="cbm"
                                                        placeholder="Enter CBM"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: #EDF2F7;"
                                                        readonly>
                                                    <div class="invalid-feedback"></div>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="pcbox" class="form-label fw-bold"
                                                        style="color: #4A5568;">Pcs/Box</label>
                                                    <input type="text" class="form-control" id="pcbox"
                                                        placeholder="Enter Pcs/Box"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="l1" class="form-label fw-bold"
                                                        style="color: #4A5568;">L1</label>
                                                    <input type="text" class="form-control" id="l1"
                                                        placeholder="Enter L1"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                        </div>

                                        <!-- Row 5 -->
                                        <div class="row mb-4">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="upc" class="form-label fw-bold"
                                                        style="color: #4A5568;">UPC</label>
                                                    <input type="text" class="form-control" id="upc"
                                                        placeholder="Enter UPC"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="b" class="form-label fw-bold"
                                                        style="color: #4A5568;">B</label>
                                                    <input type="text" class="form-control" id="b"
                                                        placeholder="Enter b"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="h1" class="form-label fw-bold"
                                                        style="color: #4A5568;">H1</label>
                                                    <input type="text" class="form-control" id="h1"
                                                        placeholder="Enter h1"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="l2Url" class="form-label fw-bold"
                                                        style="color: #4A5568;">L(2) URL</label>
                                                    <input type="text" class="form-control" id="l2Url"
                                                        placeholder="Enter URL"
                                                        style="border: 2px solid #E2E8F0; border-radius: 6px; padding: 0.75rem; background-color: white;">
                                                </div>
                                            </div>
                                        </div>

                                        <div class="row mb-4">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="dc" class="form-label fw-bold"
                                                        style="color: #4A5568;">DC</label>
                                                    <input type="text" class="form-control" id="dc"
                                                        placeholder="DC" disabled>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="weight" class="form-label fw-bold"
                                                        style="color: #4A5568;">Weight</label>
                                                    <input type="text" class="form-control" id="weight"
                                                        placeholder="Weight" disabled>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="msrp" class="form-label fw-bold"
                                                        style="color: #4A5568;">MSRP</label>
                                                    <input type="text" class="form-control" id="msrp"
                                                        placeholder="MSRP" disabled>
                                                </div>
                                            </div>
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="map" class="form-label fw-bold"
                                                        style="color: #4A5568;">MAP</label>
                                                    <input type="text" class="form-control" id="map"
                                                        placeholder="MAP" disabled>
                                                </div>
                                            </div>
                                        </div>

                                        <div class="row mb-4">
                                            <div class="col-md-3">
                                                <div class="form-group">
                                                    <label for="productImage" class="form-label fw-bold"
                                                        style="color: #4A5568;">Product Image</label>
                                                    <input type="file" class="form-control" id="productImage"
                                                        name="image" accept="image/*">
                                                    <div id="imagePreview" class="mt-2"></div>
                                                </div>
                                            </div>
                                        </div>
                                    </form>
                                </div>

                                <!-- Modal Footer -->
                                <div class="modal-footer"
                                    style="background: linear-gradient(135deg, #F8FAFF 0%, #E6F0FF 100%); border-top: 4px solid #E2E8F0; padding: 1.5rem; border-radius: 0;">
                                    <button type="button" class="btn btn-lg" data-bs-dismiss="modal"
                                        style="background: linear-gradient(135deg, #FF6B6B 0%, #FF0000 100%); color: white; border: none; border-radius: 6px; padding: 0.75rem 2rem; font-weight: 700; letter-spacing: 0.5px;">
                                        <i class="fas fa-times me-2"></i>Cancel
                                    </button>
                                    <button type="button" class="btn btn-lg" id="saveProductBtn"
                                        style="background: linear-gradient(135deg, #4ADE80 0%, #22C55E 100%); color: white; border: none; border-radius: 6px; padding: 0.75rem 2rem; font-weight: 700; letter-spacing: 0.5px;">
                                        <i class="fas fa-save me-2"></i>Save Product
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Progress Modal -->
                    <div id="progressModal" class="modal fade" tabindex="-1">
                        <div class="modal-dialog modal-lg">
                            <div class="modal-content">
                                <div class="modal-header bg-primary text-white">
                                    <h5 class="modal-title">Processing Data</h5>
                                </div>
                                <div class="modal-body">
                                    <div id="progress-container" class="mb-3"></div>
                                    <div id="error-container"></div>
                                    <div id="success-alert" class="alert alert-success" style="display:none">
                                        All sheets updated successfully!
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button id="cancelUploadBtn" class="btn btn-secondary">Cancel</button>
                                    <button id="doneBtn" class="btn btn-primary" style="display:none">Done</button>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Process Selected Modal -->
                    <div class="modal fade" id="processSelectedModal" tabindex="-1"
                        aria-labelledby="processSelectedModalLabel" aria-hidden="true">
                        <div class="modal-dialog">
                            <div class="modal-content">
                                <div class="modal-header"
                                    style="background: linear-gradient(135deg, #2c6ed5 0%, #1a56b7 100%); color: white;">
                                    <h5 class="modal-title" id="processSelectedModalLabel">Process Selected Items</h5>
                                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"
                                        aria-label="Close"></button>
                                </div>
                                <div class="modal-body">
                                    <p>Selected <span id="selectedItemCount" class="fw-bold">0</span> items. Choose fields
                                        to update:</p>

                                    <div id="fieldOperations">
                                        <div class="field-operation mb-3">
                                            <div class="row g-2 align-items-center">
                                                <div class="col-3">
                                                    <select class="form-select field-selector">
                                                        <option value="">Select Field</option>
                                                        <option value="lp">LP</option>
                                                        <option value="cp">CP</option>
                                                        <option value="frght">FRGHT</option>
                                                        <option value="ship">SHIP</option>
                                                        <option value="temu_ship">TEMU SHIP</option>
                                                        <option value="ebay2_ship">EBAY2 SHIP</option>
                                                        <option value="initial_quantity">INITIAL QUANTITY</option>
                                                        <option value="label_qty">Label QTY</option>
                                                        <option value="wt_act">WT ACT</option>
                                                        <option value="wt_decl">WT DECL</option>
                                                        <option value="l">L</option>
                                                        <option value="w">W</option>
                                                        <option value="h">H</option>
                                                        <option value="status">Status</option>
                                                    </select>
                                                </div>
                                                <div class="col-3">
                                                    <select class="form-select operation-selector">
                                                        <option value="set">=</option>
                                                        <option value="add">+</option>
                                                        <option value="subtract">-</option>
                                                        <option value="multiply">×</option>
                                                        <option value="divide">÷</option>
                                                    </select>
                                                </div>
                                                <div class="col-4">
                                                    <input type="text" class="form-control field-value"
                                                        placeholder="Enter value">
                                                </div>
                                                <div class="col-2">
                                                    <button type="button" class="btn btn-outline-danger remove-field">
                                                        <i class="fas fa-trash"></i>
                                                    </button>
                                                </div>
                                            </div>
                                        </div>
                                    </div>

                                    <button id="addFieldBtn" class="btn btn-outline-primary mt-2">
                                        <i class="fas fa-plus"></i> Add Field
                                    </button>

                                    <div class="alert alert-info mt-3">
                                        <i class="fas fa-info-circle me-2"></i>
                                        Changes will be applied to all selected items
                                    </div>

                                    <div id="batchUpdateResult" class="mt-3"></div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-secondary"
                                        data-bs-dismiss="modal">Cancel</button>
                                    <button type="button" class="btn btn-primary" id="applyChangesBtn">Apply
                                        Changes</button>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Selection actions bar -->
                    <div class="selection-actions" id="selectionActions">
                        <span class="selection-count">0 items selected</span>
                        <button class="btn btn-sm btn-light" id="cancelSelection">Cancel</button>
                        <button class="btn btn-sm btn-success" id="processSelected">Process Selected</button>
                    </div>

                    <div class="table-responsive">
                        <table id="row-callback-datatable" class="table dt-responsive nowrap w-100">
                            <thead>
                                <tr>
                                    <th class="checkbox-column">
                                        <input type="checkbox" class="select-all-checkbox" id="selectAll">
                                    </th>
                                    <th>
                                        <div style="display: flex; align-items: center; gap: 10px;">
                                            <span>Parent</span>
                                            <span id="parentCount">(0)</span>
                                        </div>
                                        <input type="text" id="parentSearch" class="form-control-sm"
                                            placeholder="Search Parent">
                                    </th>
                                    <th>
                                        <div style="display: flex; align-items: center; gap: 10px;">
                                            <span>SKU</span>
                                            <span id="skuCount">(0)</span>
                                        </div>
                                        <input type="text" id="skuSearch" class="form-control-sm"
                                            placeholder="Search SKU">
                                    </th>
                                    <th>UPC</th>
                                    <th>INV</th>
                                    <th>OV L30</th>
                                    <th>STATUS</th>
                                    <th>Unit</th>
                                    <th>LP</th>
                                    <th>CP$</th>
                                    <th>FRGHT</th>
                                    <th>SHIP</th>
                                    <th>TEMU SHIP</th>
                                    <th>EBAY2 SHIP</th>
                                    <th>INITIAL QUANTITY</th>
                                    <th>Label QTY</th>
                                    <th>WT ACT</th>
                                    <th>WT DECL</th>
                                    <th>L</th>
                                    <th>W</th>
                                    <th>H</th>
                                    <th>CBM</th>
                                    <th>L(2)</th>
                                    <th>Action</th>
                                </tr>
                            </thead>
                            <tbody id="table-body"></tbody>
                        </table>
                    </div>

                    <div id="rainbow-loader" class="rainbow-loader">
                        <div class="wave"></div>
                        <div class="wave"></div>
                        <div class="wave"></div>
                        <div class="wave"></div>
                        <div class="wave"></div>
                        <div class="loading-text">Loading Product Master Data...</div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div id="skuImageTooltip" class="sku-tooltip"></div>
@endsection

@section('script')
    <script>
        window.userPermissions = @json($permissions ?? []);
        const productPermissions = window.userPermissions['cp_masters'] || [];
        const emailColumnMap = @json($emailColumnMap ?? []);
        // Add this to the top of your DOMContentLoaded event handler
        const currentUserEmail = @json(auth()->user()->email ?? '');
        console.log('Current user email:', currentUserEmail);

        document.addEventListener('DOMContentLoaded', function() {
            // Emails and columns setup
            const emails = @json($emails ?? []);
            const columns = ["UPC","INV", "OV L30", "STATUS", "Unit", "LP", "CP$", "FRGHT", "SHIP",
                "TEMU SHIP", "EBAY2 SHIP", "INITIAL QUANTITY", "Label QTY", "WT ACT", "WT DECL", "L", "W", "H", "CBM", "L(2)", "Action"
            ];
            let selectedColumns = [];
            let selectedEmail = '';

            // Email dropdown logic
            const input = document.getElementById('emailInput');
            const dropdown = document.getElementById('emailDropdownList');
            input.addEventListener('focus', () => showDropdown(emails));
            input.addEventListener('input', function() {
                const value = this.value.toLowerCase();
                showDropdown(emails.filter(email => email.toLowerCase().includes(value)));
            });
            dropdown.addEventListener('mousedown', function(e) {
                if (e.target.classList.contains('dropdown-item')) {
                    input.value = e.target.textContent;
                    hideDropdown();
                    selectedEmail = input.value;
                    loadUserColumns(selectedEmail); // Now uses local data
                }
            });
            document.addEventListener('mousedown', function(e) {
                if (!input.contains(e.target) && !dropdown.contains(e.target)) hideDropdown();
            });

            function showDropdown(list) {
                dropdown.innerHTML = '';
                if (list.length === 0) return hideDropdown();
                list.forEach(email => {
                    const item = document.createElement('div');
                    item.className = 'dropdown-item';
                    item.textContent = email;
                    dropdown.appendChild(item);
                });
                dropdown.style.display = 'block';
            }

            function hideDropdown() {
                dropdown.style.display = 'none';
            }

            // Column dropdown logic
            const columnInput = document.getElementById('columnInput');
            const columnDropdown = document.getElementById('columnDropdownList');
            columnInput.addEventListener('focus', () => showColumnDropdown(columns));
            columnInput.addEventListener('input', function() {
                const value = this.value.toLowerCase();
                showColumnDropdown(columns.filter(col => col.toLowerCase().includes(value)));
            });
            columnDropdown.addEventListener('mousedown', function(e) {
                if (e.target.classList.contains('dropdown-item')) {
                    columnInput.value = e.target.textContent;
                    hideColumnDropdown();
                }
            });
            document.addEventListener('mousedown', function(e) {
                if (!columnInput.contains(e.target) && !columnDropdown.contains(e.target))
                    hideColumnDropdown();
            });

            function showColumnDropdown(list) {
                columnDropdown.innerHTML = '';
                if (list.length === 0) return hideColumnDropdown();
                list.forEach(col => {
                    const item = document.createElement('div');
                    item.className = 'dropdown-item';
                    item.textContent = col;
                    columnDropdown.appendChild(item);
                });
                columnDropdown.style.display = 'block';
            }

            function hideColumnDropdown() {
                columnDropdown.style.display = 'none';
            }

            // Add column to selectedColumns
            document.getElementById('addColumnBtn').addEventListener('click', function() {
                const col = columnInput.value.trim();
                if (col && !selectedColumns.includes(col)) {
                    selectedColumns.push(col);
                    renderCurrentColumns();
                    columnInput.value = '';
                }
            });

            // Render current columns with remove buttons
            function renderCurrentColumns() {
                const container = document.getElementById('currentColumns');
                container.innerHTML = '';
                selectedColumns.forEach(col => {
                    const tag = document.createElement('span');
                    tag.className = 'badge bg-primary me-2';
                    tag.textContent = col;
                    const removeBtn = document.createElement('button');
                    removeBtn.className = 'btn btn-sm btn-danger ms-1';
                    removeBtn.textContent = '×';
                    removeBtn.onclick = () => {
                        selectedColumns = selectedColumns.filter(c => c !== col);
                        renderCurrentColumns();
                    };
                    tag.appendChild(removeBtn);
                    container.appendChild(tag);
                });
            }

            // Load user's current columns from backend
            function loadUserColumns(email) {
                selectedColumns = emailColumnMap[email] || [];
                renderCurrentColumns();
            }

            // Open permission modal on button click
            // document.getElementById('culomnPermissionBtn').addEventListener('click', function() {
            //     const modal = document.getElementById('permissionModal');
            //     if (modal && typeof bootstrap !== 'undefined' && bootstrap.Modal) {
            //         const permissionModal = new bootstrap.Modal(modal);
            //         permissionModal.show();
            //     }
            // });

            // Save permission AJAX
            document.getElementById('savePermissionBtn').addEventListener('click', function() {
                if (!selectedEmail) {
                    showToast('danger', 'Please select an email');
                    return;
                }
                fetch('/auth/save-column-permission', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'X-CSRF-TOKEN': document.querySelector('meta[name="csrf-token"]').content
                        },
                        body: JSON.stringify({
                            user_email: selectedEmail,
                            columns: selectedColumns,
                            module: 'product_master'
                        })
                    })
                    .then(res => res.json())
                    .then(data => {
                        if (data.success) {
                            showToast('success', 'Permission saved!');
                            document.getElementById('permissionModal').classList.remove('show');
                            document.getElementById('permissionModal').setAttribute('aria-hidden',
                                'true');
                            document.body.classList.remove('modal-open');
                            document.querySelector('.modal-backdrop')?.remove();
                        } else {
                            showToast('danger', data.message || 'Error saving permission');
                        }
                    })
                    .catch(() => {
                        showToast('danger', 'Error saving permission');
                    });
            });

            // Toast notification function
            function showToast(type, message) {
                document.querySelectorAll('.custom-toast').forEach(t => t.remove());
                const toast = document.createElement('div');
                toast.className =
                    `custom-toast toast align-items-center text-bg-${type} border-0 show position-fixed top-0 end-0 m-4`;
                toast.style.zIndex = 2000;
                toast.setAttribute('role', 'alert');
                toast.setAttribute('aria-live', 'assertive');
                toast.setAttribute('aria-atomic', 'true');
                toast.innerHTML = `
                    <div class="d-flex">
                        <div class="toast-body">${message}</div>
                        <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
                    </div>
                `;
                document.body.appendChild(toast);

                setTimeout(() => {
                    toast.classList.remove('show');
                    setTimeout(() => toast.remove(), 500);
                }, 3000);

                toast.querySelector('[data-bs-dismiss="toast"]').onclick = () => {
                    toast.classList.remove('show');
                    setTimeout(() => toast.remove(), 500);
                };
            }
        });
    </script>
    <script src="https://code.jquery.com/jquery-3.6.4.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            // Initialize with 75% zoom
            document.body.style.zoom = "75%";

            // Store the loaded data globally
            let tableData = [];

            // Track selected items with both SKU and ID
            let selectedItems = {}; // Format: { sku: { id: 123, checked: true } }

            // Get CSRF token from meta tag
            const csrfToken = document.querySelector('meta[name="csrf-token"]').content;

            // Show loader immediately
            document.getElementById('rainbow-loader').style.display = 'block';

            // Initialize all components
            initializeTable();

            // Centralized AJAX request function with CSRF protection
            function makeRequest(url, method, data = {}) {
                const headers = {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json',
                    'X-CSRF-TOKEN': csrfToken
                };

                // Include CSRF token in request body for POST/PUT/PATCH/DELETE
                if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(method.toUpperCase())) {
                    data._token = csrfToken;
                }

                return fetch(url, {
                    method: method,
                    headers: headers,
                    body: method === 'GET' ? null : JSON.stringify(data)
                });
            }

            // Load product data from server
            function loadData() {
                makeRequest('/product-master-data-view', 'GET')
                    .then(response => {
                        if (!response.ok) {
                            throw new Error(`HTTP error! status: ${response.status}`);
                        }
                        return response.json();
                    })
                    .then(response => {
                        if (response && response.data && Array.isArray(response.data)) {
                            tableData = response.data;
                            renderTable(tableData);
                            updateParentOptions();

                            // Add this block to update counts
                            const parentSet = new Set();
                            let skuCount = 0;
                            tableData.forEach(item => {
                                if (item.Parent) parentSet.add(item.Parent);
                                // Only count SKUs that do NOT contain 'PARENT'
                                if (item.SKU && !String(item.SKU).toUpperCase().includes('PARENT'))
                                    skuCount++;
                            });
                            document.getElementById('parentCount').textContent = `(${parentSet.size})`;
                            document.getElementById('skuCount').textContent = `(${skuCount})`;

                        } else {
                            showError('Invalid data format received from server');
                        }
                        document.getElementById('rainbow-loader').style.display = 'none';
                    })
                    .catch(error => {
                        showError('Failed to load product data: ' + error.message);
                        document.getElementById('rainbow-loader').style.display = 'none';
                    });
            }

            // Modified renderTable function to respect column permissions
            function renderTable(data) {
                const tbody = document.getElementById('table-body');
                tbody.innerHTML = '';

                if (data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="17" class="text-center">No products found</td></tr>';
                    return;
                }

                const hasEditPermission = productPermissions.includes('edit');
                const hasDeletePermission = productPermissions.includes('delete');
                const selectionMode = document.querySelector('.checkbox-column')?.style.display === 'table-cell';

                // Get columns to hide for current user
                const hiddenColumns = getUserHiddenColumns();

                // All available columns
                const allColumns = [
                    "Parent", "SKU", "UPC", "INV", "OV L30", "STATUS", "Unit", "LP", "CP$",
                    "FRGHT", "SHIP", "TEMU SHIP", "EBAY2 SHIP", "INITIAL QUANTITY", "Label QTY", "WT ACT", "WT DECL", "L", "W", "H",
                    "CBM", "L(2)", "Action"
                ];

                // Filter to get visible columns
                const visibleColumns = allColumns.filter(col => !hiddenColumns.includes(col));

                // Update table header to show only visible columns
                updateTableHeader(hiddenColumns);

                // Combine search results and selected items (if in selection mode)
                let displayItems = [...data];
                if (selectionMode && Object.keys(selectedItems).length > 0) {
                    const selectedItemsList = tableData.filter(item => selectedItems[item.SKU]);
                    const combinedItems = [...data];

                    selectedItemsList.forEach(item => {
                        if (!data.some(d => d.SKU === item.SKU)) {
                            combinedItems.push(item);
                        }
                    });

                    displayItems = combinedItems;
                }

                // Sort selected items first
                if (selectionMode) {
                    displayItems.sort((a, b) => {
                        const aSelected = selectedItems[a.SKU];
                        const bSelected = selectedItems[b.SKU];
                        if (aSelected && !bSelected) return -1;
                        if (!aSelected && bSelected) return 1;
                        return 0;
                    });
                }

                // Before rendering rows, calculate totals for each parent
                const parentTotals = {};
                data.forEach(item => {
                    if (item.Parent && !String(item.SKU).toUpperCase().includes('PARENT')) {
                        if (!parentTotals[item.Parent]) {
                            parentTotals[item.Parent] = {
                                inv: 0,
                                ovl30: 0
                            };
                        }
                        parentTotals[item.Parent].inv += Number(item.shopify_inv) || 0;
                        parentTotals[item.Parent].ovl30 += Number(item.shopify_quantity) || 0;
                    }
                });

                // Render rows
                displayItems.forEach(item => {
                    const row = document.createElement('tr');

                    // Parent row
                    if (item.SKU && item.SKU.toUpperCase().includes('PARENT')) {
                        row.style.backgroundColor = 'rgba(13, 110, 253, 0.2)';
                        row.style.fontWeight = '500';
                        const totals = parentTotals[item.Parent] || {
                            inv: 0,
                            ovl30: 0
                        };

                        visibleColumns.forEach(col => {
                            let cell = document.createElement('td');
                            switch (col) {
                                case "Parent":
                                    cell.textContent = escapeHtml(item.Parent) || '-';
                                    break;
                                case "SKU":
                                    cell.textContent = escapeHtml(item.SKU) || '-';
                                    break;
                                case "UPC":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.upc, 0);
                                    break;
                                case "INV":
                                    cell.innerHTML = `<b>${totals.inv}</b>`;
                                    break;
                                case "OV L30":
                                    cell.innerHTML = `<b>${totals.ovl30}</b>`;
                                    break;
                                case "STATUS":
                                    cell.textContent = escapeHtml(item.status) || '-';
                                    break;
                                case "Unit":
                                    cell.textContent = item.unit || '-';
                                    break;
                                case "LP":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.lp, 2);
                                    break;
                                case "CP$":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.cp, 2);
                                    break;
                                case "FRGHT":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.frght, 2);
                                    break;
                                case "SHIP":
                                    cell.textContent = escapeHtml(item.ship) || '-';
                                    break;
                                    case "TEMU SHIP":
                                    cell.textContent = escapeHtml(item.temu_ship) || '-';
                                    break;
                                    case "EBAY2 SHIP":
                                    cell.textContent = escapeHtml(item.ebay2_ship) || '-';
                                    break;
                                    case "INITIAL QUANTITY":
                                    cell.textContent = escapeHtml(item.initial_quantity) || '-';
                                    break;
                                case "Label QTY":
                                    cell.textContent = escapeHtml(item.label_qty) || '0';
                                    break;
                                case "WT ACT":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.wt_act || 0, 2);
                                    break;
                                case "WT DECL":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.wt_decl || 0, 2);
                                    break;
                                case "L":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.l || 0, 2);
                                    break;
                                case "W":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.w || 0, 2);
                                    break;
                                case "H":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.h || 0, 2);
                                    break;
                                case "CBM":
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item.cbm, 4);
                                    break;
                                case "L(2)":
                                    cell.className = 'text-center';
                                    cell.innerHTML = item.l2_url ?
                                        `<a href="${escapeHtml(item.l2_url)}" target="_blank"><i class="fas fa-external-link-alt"></i></a>` :
                                        '-';
                                    break;
                                case "Action":
                                    cell.className = 'text-center';
                                    cell.innerHTML = `
                            <div class="d-inline-flex">
                                ${hasEditPermission ? 
                                    `<button class="btn btn-sm btn-outline-primary edit-btn me-1" data-sku="${escapeHtml(item.SKU)}">
                                                            <i class="bi bi-pencil-square"></i>
                                                        </button>` 
                                    : ''
                                }
                                ${hasDeletePermission ? 
                                    `<button class="btn btn-sm btn-outline-danger delete-btn" data-id="${escapeHtml(item.id)}" data-sku="${escapeHtml(item.SKU)}">
                                                            <i class="bi bi-trash"></i>
                                                        </button>` 
                                    : ''
                                }
                                ${(!hasEditPermission && !hasDeletePermission) ? '-' : ''}
                            </div>
                        `;
                                    break;
                                default:
                                    cell.textContent = '-';
                            }
                            row.appendChild(cell);
                        });

                        tbody.appendChild(row);
                        return;
                    }

                    // Add checkbox cell if selection mode is active
                    if (selectionMode) {
                        const isChecked = selectedItems[item.SKU] ? 'checked' : '';
                        const checkboxCell = document.createElement('td');
                        checkboxCell.className = 'checkbox-cell';
                        checkboxCell.style.textAlign = 'center';
                        checkboxCell.innerHTML = `
                        <input type="checkbox" class="row-checkbox" 
                            data-sku="${escapeHtml(item.SKU || '')}" 
                            data-id="${escapeHtml(item.id || '')}" ${isChecked}>
                    `;
                        row.appendChild(checkboxCell);
                    }

                    // Calculate CBM and FRGHT using the formulas
                    const l = parseFloat(item.l);
                    const w = parseFloat(item.w);
                    const h = parseFloat(item.h);
                    let cbm = '';
                    let frght = '';
                    if (!isNaN(l) && !isNaN(w) && !isNaN(h)) {
                        cbm = (((l * 2.54) * (w * 2.54) * (h * 2.54)) / 1000000);
                        frght = cbm * 200;
                        cbm = cbm.toFixed(4);
                        frght = frght.toFixed(2);
                    }

                    // Render only visible columns
                    visibleColumns.forEach(col => {
                        let cell = document.createElement('td');
                        switch (col) {
                            case "Parent":
                                cell.textContent = escapeHtml(item.Parent) || '-';
                                break;
                            case "SKU":
                                cell.innerHTML = `
                                    <span class="sku-hover" 
                                        data-sku="${escapeHtml(item.SKU) || ''}" 
                                        data-image="${item.image_path ? item.image_path : ''}">
                                        ${escapeHtml(item.SKU) || '-'}
                                    </span>
                                `;
                                break;
                            case "UPC":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.upc, 0);
                                break;
                            case "INV":
                                cell.textContent = escapeHtml(item.shopify_inv) || '-';
                                break;
                            case "OV L30":
                                cell.textContent = escapeHtml(item.shopify_quantity) || '-';
                                break;
                            case "STATUS":
                                cell.textContent = escapeHtml(item.status) || '-';
                                break;
                            case "Unit":
                                cell.textContent = item.unit || '-';
                                break;
                            case "LP":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.lp, 2);
                                break;
                            case "CP$":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.cp, 2);
                                break;
                            case "FRGHT":
                                cell.className = 'text-center';
                                cell.textContent = frght || '-';
                                break;
                            case "SHIP":
                                cell.textContent = escapeHtml(item.ship) || '-';
                                break;
                            case "TEMU SHIP":
                                cell.textContent = escapeHtml(item.temu_ship) || '-';
                                break;
                            case "EBAY2 SHIP":
                                cell.textContent = escapeHtml(item.ebay2_ship) || '-';
                                break;
                            case "INITIAL QUANTITY":
                                cell.textContent = escapeHtml(item.initial_quantity) || '-';
                                break;
                            case "Label QTY":
                                cell.textContent = escapeHtml(item.label_qty) || '0';
                                break;
                            case "WT ACT":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.wt_act || 0, 2);
                                break;
                            case "WT DECL":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.wt_decl || 0, 2);
                                break;
                            case "L":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.l || 0, 2);
                                break;
                            case "W":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.w || 0, 2);
                                break;
                            case "H":
                                cell.className = 'text-center';
                                cell.textContent = formatNumber(item.h || 0, 2);
                                break;
                            case "CBM":
                                cell.className = 'text-center';
                                cell.textContent = cbm || '-';
                                break;
                            case "L(2)":
                                cell.className = 'text-center';
                                cell.innerHTML = item.l2_url ?
                                    `<a href="${escapeHtml(item.l2_url)}" target="_blank"><i class="fas fa-external-link-alt"></i></a>` :
                                    '-';
                                break;
                            case "Action":
                                cell.className = 'text-center';
                                cell.innerHTML = `
                        <div class="d-inline-flex">
                            ${hasEditPermission ? 
                                `<button class="btn btn-sm btn-outline-primary edit-btn me-1" data-sku="${escapeHtml(item.SKU)}">
                                                        <i class="bi bi-pencil-square"></i>
                                                    </button>` 
                                : ''
                            }
                            ${hasDeletePermission ? 
                                `<button class="btn btn-sm btn-outline-danger delete-btn" data-id="${escapeHtml(item.id)}" data-sku="${escapeHtml(item.SKU)}">
                                                        <i class="bi bi-trash"></i>
                                                    </button>` 
                                : ''
                            }
                            ${(!hasEditPermission && !hasDeletePermission) ? '-' : ''}
                        </div>
                    `;
                                break;
                            default:
                                cell.textContent = '-';
                        }
                        row.appendChild(cell);
                    });

                    tbody.appendChild(row);
                });

                if (hasEditPermission) {
                    setupEditButtons();
                    setupDeleteButtons();
                }



                // bindRowCheckboxes();
                bindSelectAllCheckbox();

                updateSelectionCount();
                // restoreSelectAllState();
            }


            // Updated function to show all columns except those in the hidden list
            function updateTableHeader(hiddenColumns) {
                const thead = document.querySelector('#row-callback-datatable thead tr');

                // Store the checkbox column if it exists
                const checkboxTh = thead.querySelector('.checkbox-column');

                // Clear current header
                thead.innerHTML = '';

                // Re-add checkbox column if it exists
                if (checkboxTh) {
                    thead.appendChild(checkboxTh.cloneNode(true));
                }

                // All available columns
                const allColumns = [
                    "Parent", "SKU", "UPC", "INV", "OV L30", "STATUS", "Unit", "LP", "CP$",
                    "FRGHT", "SHIP", "TEMU SHIP", "EBAY2 SHIP", "INITIAL QUANTITY", "Label QTY", "WT ACT", "WT DECL", "L", "W", "H",
                    "CBM", "L(2)", "Action"
                ];

                // Add only columns that are not in the hidden list
                allColumns.forEach(colName => {
                    if (!hiddenColumns.includes(colName)) {
                        const th = document.createElement('th');

                        if (colName === "Parent" || colName === "SKU") {
                            th.innerHTML = `
                    <div style="display: flex; align-items: center; gap: 10px;">
                        <span>${colName}</span>
                        <span id="${colName.toLowerCase()}Count">(0)</span>
                    </div>
                    <input type="text" id="${colName.toLowerCase()}Search" class="form-control-sm" placeholder="Search ${colName}">
                `;
                        } else {
                            th.textContent = colName;
                        }

                        thead.appendChild(th);
                    }
                });

                // Ensure Action column is always visible if user has permissions
                const hasEditPermission = productPermissions.includes('edit');
                const hasDeletePermission = productPermissions.includes('delete');

                if ((hasEditPermission || hasDeletePermission) && !thead.querySelector('th').textContent.includes(
                        "Action")) {
                    let actionExists = false;
                    for (let i = 0; i < thead.children.length; i++) {
                        if (thead.children[i].textContent.trim() === "Action") {
                            actionExists = true;
                            break;
                        }
                    }

                    if (!actionExists) {
                        const actionTh = document.createElement('th');
                        actionTh.textContent = "Action";
                        thead.appendChild(actionTh);
                    }
                }

                // Update parent and SKU counts if they're visible
                const parentCount = document.getElementById('parentCount');
                const skuCount = document.getElementById('skuCount');

                if (parentCount && skuCount) {
                    const parentSet = new Set();
                    let skuCount = 0;
                    tableData.forEach(item => {
                        if (item.Parent) parentSet.add(item.Parent);
                        // Only count SKUs that do NOT contain 'PARENT'
                        if (item.SKU && !String(item.SKU).toUpperCase().includes('PARENT'))
                            skuCount++;
                    });

                    document.getElementById('parentCount').textContent = `(${parentSet.size})`;
                    document.getElementById('skuCount').textContent = `(${skuCount})`;
                }
            }

            // Modified row rendering to respect column permissions
            function renderRow(item, allowedColumns, isParent = false) {
                const row = document.createElement('tr');

                if (isParent) {
                    row.style.backgroundColor = 'rgba(13, 110, 253, 0.2)';
                    row.style.fontWeight = '500';
                }

                // If in selection mode, add checkbox first
                if (selectionMode && !isParent) {
                    const isChecked = selectedItems[item.SKU] ? 'checked' : '';
                    const checkboxCell = document.createElement('td');
                    checkboxCell.className = 'checkbox-cell';
                    checkboxCell.style.textAlign = 'center';
                    checkboxCell.innerHTML = `
            <input type="checkbox" class="row-checkbox" 
                data-sku="${escapeHtml(item.SKU || '')}" 
                data-id="${escapeHtml(item.id || '')}" ${isChecked}>
        `;
                    row.appendChild(checkboxCell);
                }

                // Add only allowed columns
                allowedColumns.forEach(colName => {
                    const cell = document.createElement('td');

                    // Add appropriate cell content based on column name
                    switch (colName) {
                        case "Parent":
                            cell.textContent = escapeHtml(item.Parent) || '-';
                            break;

                        case "SKU":
                            cell.innerHTML = `
                    <span class="sku-hover" 
                        data-sku="${escapeHtml(item.SKU) || ''}" 
                        data-image="${item.image_path ? item.image_path : ''}">
                        ${escapeHtml(item.SKU) || '-'}
                    </span>
                `;
                            break;

                            // Add cases for other columns...

                        case "Action":
                            cell.className = 'text-center';
                            cell.innerHTML = `
                    <div class="d-inline-flex">
                        ${hasEditPermission ? 
                            `<button class="btn btn-sm btn-outline-primary edit-btn me-1" data-sku="${escapeHtml(item.SKU)}">
                                                                            <i class="bi bi-pencil-square"></i>
                                                                        </button>` 
                            : ''
                        }
                        ${hasDeletePermission ? 
                            `<button class="btn btn-sm btn-outline-danger delete-btn" data-id="${escapeHtml(item.id)}" data-sku="${escapeHtml(item.SKU)}">
                                                                            <i class="bi bi-trash"></i>
                                                                        </button>` 
                            : ''
                        }
                        ${(!hasEditPermission && !hasDeletePermission) ? '-' : ''}
                    </div>
                `;
                            break;

                        default:
                            // Handle any other column using the columnDefs mapping
                            if (columnDefs[colName]) {
                                const key = columnDefs[colName].key;
                                // Format based on column type
                                if (["lp", "cp", "frght"].includes(key)) {
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item[key], 2);
                                } else if (["wt_act", "wt_decl", "l", "w", "h"].includes(key)) {
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item[key] || 0, 2);
                                } else if (key === "cbm") {
                                    cell.className = 'text-center';
                                    cell.textContent = formatNumber(item[key], 4);
                                } else if (key === "l2_url") {
                                    cell.className = 'text-center';
                                    cell.innerHTML = item[key] ?
                                        `<a href="${escapeHtml(item[key])}" target="_blank"><i class="fas fa-external-link-alt"></i></a>` :
                                        '-';
                                } else {
                                    cell.textContent = escapeHtml(item[key]) || '-';
                                }
                            }
                            break;
                    }

                    row.appendChild(cell);
                });

                return row;
            }

            // In your initializeTable function, modify:
            function initializeTable() {
                // Get columns to hide
                const hiddenColumns = getUserHiddenColumns();

                // Update header to exclude hidden columns
                updateTableHeader(hiddenColumns);

                // Rest of initialization...
                loadData();
                setupSearch();
                setupSelectFilter();
                setupHeaderColumnSearch();
                setupExcelExport();
                setupAddProductModal();
                setupProgressModal();
                setupSelectionMode();
                setupBatchProcessing();
            }

            function setupSelectFilter(){
                const fieldSelect = document.querySelector('.field-selector-wrapper select');
                let fieldInput = null; 
                let selectedField = "";

                fieldSelect.addEventListener('change', function () {
                    selectedField = this.value.trim();

                    if (fieldInput) fieldInput.remove();

                    if (selectedField) {
                        fieldInput = document.createElement('input');
                        fieldInput.type = 'text';
                        fieldInput.placeholder = `Search ${selectedField.toUpperCase()}...`;
                        fieldInput.classList.add('form-control', 'mt-2');
                        fieldInput.style.fontSize = '13px';
                        fieldInput.style.border = '1px solid #92c1ff';
                        fieldInput.style.borderRadius = '6px';

                        fieldSelect.parentElement.appendChild(fieldInput);

                        fieldInput.addEventListener('input', debounce(function () {
                            const searchTerm = fieldInput.value.toLowerCase().trim();

                            let filteredData = [...tableData];

                            if (searchTerm) {
                                filteredData = filteredData.filter(item => {
                                    const value = String(item[selectedField] ?? '').toLowerCase();
                                    return value === searchTerm;
                                });
                            }

                            renderTable(filteredData);
                        }, 300));
                    }
                });
            }
            
            function setupSearch() {
                const searchInput = document.getElementById('customSearch');
                
                const skuSearchInput = document.getElementById('skuSearch');
                
                const parentSearchInput = document.getElementById('parentSearch');

                const clearButton = document.getElementById('clearSearch');

                // Global search (search all columns)
                if (searchInput) {
                    searchInput.addEventListener('input', debounce(function () {
                        const searchTerm = searchInput.value.toLowerCase().trim();

                        let filteredData = [...tableData];
                        if (searchTerm) {
                            filteredData = filteredData.filter(item =>
                                Object.values(item).some(value =>
                                    String(value).toLowerCase().includes(searchTerm)
                                )
                            );
                        }

                        renderTable(filteredData);
                    }, 300));
                }

                // SKU search (only filter by SKU)
                if (skuSearchInput) {
                    
                    skuSearchInput.addEventListener('input', debounce(function () {
                        const skuValue = skuSearchInput.value.toLowerCase().trim();

                        let filteredData = [...tableData];
                        if (skuValue) {
                            filteredData = filteredData.filter(item =>
                                (item.SKU || item.sku || '').toLowerCase().includes(skuValue)
                            );
                        }

                        renderTable(filteredData);
                    }, 300));
                }

                // Parent search (only filter by Parent)
                if (parentSearchInput) {
                    parentSearchInput.addEventListener('input', debounce(function () {
                        const parentValue = parentSearchInput.value.toLowerCase().trim();

                        let filteredData = [...tableData];
                        if (parentValue) {
                            filteredData = filteredData.filter(item =>
                                (item.Parent || item.parent || '').toLowerCase().includes(parentValue)
                            );
                        }

                        renderTable(filteredData);
                    }, 300));
                }

                // Clear all searches
                if (clearButton) {
                    clearButton.addEventListener('click', function () {
                        if (searchInput) searchInput.value = '';
                        if (skuSearchInput) skuSearchInput.value = '';
                        if (parentSearchInput) parentSearchInput.value = '';
                        renderTable(tableData);
                    });
                }
            }

            // Setup header column search
            function setupHeaderColumnSearch() {
                const parentSearch = document.getElementById('parentSearch');
                const skuSearch = document.getElementById('skuSearch');
                const globalSearch = document.getElementById('customSearch');

                function applyFilters() {
                    const parentValue = parentSearch ? parentSearch.value.toLowerCase().trim() : '';
                    const skuValue = skuSearch ? skuSearch.value.toLowerCase().trim() : '';
                    const globalValue = globalSearch ? globalSearch.value.toLowerCase().trim() : '';

                    console.log('Header Parent Search:', parentValue);
                    console.log('Header SKU Search:', skuValue);
                    console.log('Header Global Search:', globalValue);

                    let filteredData = [...tableData];

                    if (parentValue) {
                        filteredData = filteredData.filter(item =>
                            (item.Parent || item.parent || '').toLowerCase().includes(parentValue)
                        );
                    }

                    if (skuValue) {
                        filteredData = filteredData.filter(item =>
                            (item.SKU || item.sku || '').toLowerCase().includes(skuValue)
                        );
                    }

                    if (globalValue) {
                        filteredData = filteredData.filter(item =>
                            Object.values(item).some(value =>
                                String(value).toLowerCase().includes(globalValue)
                            )
                        );
                    }

                    renderTable(filteredData);
                }

                if (parentSearch) parentSearch.addEventListener('input', debounce(applyFilters, 300));
                if (skuSearch) skuSearch.addEventListener('input', debounce(applyFilters, 300));
            }


            // Function to get columns to hide for current user
            function getUserHiddenColumns() {
                // Default columns to hide if user has no specific permissions
                const defaultHiddenColumns = [];

                if (currentUserEmail && emailColumnMap[currentUserEmail]) {
                    return emailColumnMap[currentUserEmail];
                }

                return defaultHiddenColumns;
            }

            // Update Excel export function to exclude hidden columns
            function setupExcelExport() {
                document.getElementById('downloadExcel').addEventListener('click', function() {
                    const hiddenColumns = getUserHiddenColumns();
                    const allColumns = [
                        "Parent", "SKU", "UPC", "INV", "OV L30", "STATUS", "Unit", "LP", "CP$",
                        "FRGHT", "SHIP", "TEMU SHIP", "EBAY2 SHIP", "INITIAL QUANTITY", "Label QTY", "WT ACT", "WT DECL", "L", "W", "H",
                        "CBM", "L(2)", "DC", "Pcs/Box", "L1", "B", "H1", "Weight", "MSRP", "MAP", "UPC"
                    ];

                    // Filter out hidden columns
                    const visibleColumns = allColumns.filter(col => !hiddenColumns.includes(col));

                    // Column definitions with their data keys
                    const columnDefs = {
                        "Parent": {
                            key: "Parent"
                        },
                        "SKU": {
                            key: "SKU"
                        },
                        "UPC": {
                            key: "upc"
                        },
                        "INV": {
                            key: "shopify_inv"
                        },
                        "OV L30": {
                            key: "shopify_quantity"
                        },
                        "STATUS": {
                            key: "status"
                        },
                        "Unit": {
                            key: "unit"
                        },
                        "LP": {
                            key: "lp"
                        },
                        "CP$": {
                            key: "cp"
                        },
                        "FRGHT": {
                            key: "frght"
                        },
                        "SHIP": {
                            key: "ship"
                        },
                        "TEMU SHIP": {
                            key: "temu_ship"
                        },
                        "EBAY2 SHIP": {
                            key: "ebay2_ship"
                        },
                        "INITIAL QUANTITY": {
                            key: "initial_quantity"
                        },
                        "Label QTY": {
                            key: "label_qty"
                        },
                        "WT ACT": {
                            key: "wt_act"
                        },
                        "WT DECL": {
                            key: "wt_decl"
                        },
                        "L": {
                            key: "l"
                        },
                        "W": {
                            key: "w"
                        },
                        "H": {
                            key: "h"
                        },
                        "CBM": {
                            key: "cbm"
                        },
                        "L(2)": {
                            key: "l2_url"
                        },
                        "DC": {
                            key: "dc"
                        },
                        "Pcs/Box": {
                            key: "pcs_per_box"
                        },
                        "L1": {
                            key: "l1"
                        },
                        "B": {
                            key: "b"
                        },
                        "H1": {
                            key: "h1"
                        },
                        "Weight": {
                            key: "weight"
                        },
                        "MSRP": {
                            key: "msrp"
                        },
                        "MAP": {
                            key: "map"
                        },
                        "UPC": {
                            key: "upc"
                        }
                    };

                    // Show loader or indicate download is in progress
                    document.getElementById('downloadExcel').innerHTML =
                        '<i class="fas fa-spinner fa-spin"></i> Generating...';
                    document.getElementById('downloadExcel').disabled = true;

                    // Use setTimeout to avoid UI freeze for large datasets
                    setTimeout(() => {
                        try {
                            // Create worksheet data array
                            const wsData = [];

                            // Add header row
                            wsData.push(visibleColumns);

                            // Add data rows
                            tableData.forEach(item => {
                                const row = [];
                                visibleColumns.forEach(col => {
                                    const colDef = columnDefs[col];
                                    if (colDef) {
                                        const key = colDef.key;
                                        let value = item[key] !== undefined && item[
                                            key] !== null ? item[key] : '';

                                        // Format special columns
                                        if (["lp", "cp", "frght"].includes(key)) {
                                            value = parseFloat(value) || 0;
                                        } else if (["wt_act", "wt_decl", "l", "w",
                                                "h"
                                            ].includes(key)) {
                                            value = parseFloat(value) || 0;
                                        } else if (key === "cbm") {
                                            value = parseFloat(value) || 0;
                                        }

                                        row.push(value);
                                    } else {
                                        row.push('');
                                    }
                                });
                                wsData.push(row);
                            });

                            // Create workbook and worksheet
                            const wb = XLSX.utils.book_new();
                            const ws = XLSX.utils.aoa_to_sheet(wsData);

                            // Set column widths
                            const wscols = visibleColumns.map(col => {
                                // Adjust width based on column type
                                if (["Parent", "SKU"].includes(col)) {
                                    return {
                                        wch: 20
                                    }; // Wider for text columns
                                } else if (["STATUS", "Unit"].includes(col)) {
                                    return {
                                        wch: 12
                                    };
                                } else {
                                    return {
                                        wch: 10
                                    }; // Default width for numeric columns
                                }
                            });
                            ws['!cols'] = wscols;

                            // Style the header row
                            const headerRange = XLSX.utils.decode_range(ws['!ref']);
                            for (let C = headerRange.s.c; C <= headerRange.e.c; ++C) {
                                const cell = XLSX.utils.encode_cell({
                                    r: 0,
                                    c: C
                                });
                                if (!ws[cell]) continue;

                                // Add header style
                                ws[cell].s = {
                                    fill: {
                                        fgColor: {
                                            rgb: "2C6ED5"
                                        }
                                    },
                                    font: {
                                        bold: true,
                                        color: {
                                            rgb: "FFFFFF"
                                        }
                                    },
                                    alignment: {
                                        horizontal: "center"
                                    }
                                };
                            }

                            // Add the worksheet to the workbook
                            XLSX.utils.book_append_sheet(wb, ws, "Product Master");

                            // Generate Excel file and trigger download
                            XLSX.writeFile(wb, "product_master_export.xlsx");

                            // Show success toast
                            showToast('success', 'Excel file downloaded successfully!');
                        } catch (error) {
                            console.error("Excel export error:", error);
                            showToast('danger', 'Failed to export Excel file.');
                        } finally {
                            // Reset button state
                            document.getElementById('downloadExcel').innerHTML =
                                '<i class="fas fa-file-excel me-1"></i> Download Excel';
                            document.getElementById('downloadExcel').disabled = false;
                        }
                    }, 100); // Small timeout to allow UI to update
                });
            }

            // Initialize the add product modal
            function setupAddProductModal() {
                const modal = document.getElementById('addProductModal');
                const saveBtn = document.getElementById('saveProductBtn');
                const refreshParentsBtn = document.getElementById('refreshParents');

                // Setup event listeners for calculations
                document.getElementById('w')?.addEventListener('input', calculateCBM);
                document.getElementById('l')?.addEventListener('input', calculateCBM);
                document.getElementById('h')?.addEventListener('input', calculateCBM);
                document.getElementById('cp')?.addEventListener('input', calculateLP);

                refreshParentsBtn.addEventListener('click', updateParentOptions);

                saveBtn.addEventListener('click', async function() {
                    if (!validateProductForm()) return;

                    const formData = getFormData();
                    formData.append('operation', 'create');

                    try {
                        const response = await fetch('/product_master/store', {
                            method: 'POST',
                            // Do NOT set Content-Type when using FormData!
                            headers: {
                                'X-CSRF-TOKEN': csrfToken
                            },
                            body: formData
                        });
                        const data = await response.json();

                        if (!response.ok) throw new Error(data.message ||
                            `Server returned status ${response.status}`);

                        showAlert('success', 'Product saved to database!');
                        bootstrap.Modal.getInstance(modal).hide();
                        loadData();
                        resetProductForm();
                    } catch (error) {
                        showAlert('danger', error.message);
                    }
                });

                modal.addEventListener('hidden.bs.modal', resetProductForm);
            }

            // Image preview on file select
            document.getElementById('productImage').addEventListener('change', function(e) {
                const preview = document.getElementById('imagePreview');
                preview.innerHTML = '';
                if (this.files && this.files[0]) {
                    const reader = new FileReader();
                    reader.onload = function(ev) {
                        preview.innerHTML =
                            `<img src="${ev.target.result}" alt="Preview" style="max-width:120px;max-height:120px;border-radius:8px;">`;
                    };
                    reader.readAsDataURL(this.files[0]);
                }
            });

            // Calculate CBM based on dimensions
            function calculateCBM() {
                const w = parseFloat(document.getElementById('w').value) || 0;
                const l = parseFloat(document.getElementById('l').value) || 0;
                const h = parseFloat(document.getElementById('h').value) || 0;

                // Convert to cm then to m³ per user formula: ((L*2.54)*(W*2.54)*(H*2.54))/1000000
                let cbm = 0;
                if (w > 0 && l > 0 && h > 0) {
                    cbm = ((l * 2.54) * (w * 2.54) * (h * 2.54)) / 1000000;
                }
                document.getElementById('cbm').value = cbm ? cbm.toFixed(4) : '';

                // FRGHT formula: CBM * 200
                const frght = cbm * 200;
                document.getElementById('freght').value = cbm ? frght.toFixed(2) : '';

                // Recalculate LP as well
                calculateLP();
            }

            // Calculate LP based on CP and FRGHT
            function calculateLP() {
                const cp = parseFloat(document.getElementById('cp').value) || 0;
                const frght = parseFloat(document.getElementById('freght').value) || 0;
                // LP formula: CP + FRGHT
                const lp = cp + frght;
                document.getElementById('lp').value = lp.toFixed(2);
            }

            // Validate the product form
            function validateProductForm() {
                const sku = document.getElementById('sku').value;
                // If SKU contains 'PARENT', skip required validation
                if (sku && sku.toUpperCase().includes('PARENT')) {
                    // Clear any previous errors
                    document.querySelectorAll('.is-invalid').forEach(el => el.classList.remove('is-invalid'));
                    document.querySelectorAll('.invalid-feedback').forEach(el => el.textContent = '');
                    document.getElementById('form-errors').innerHTML = '';
                    return true;
                }

                let isValid = true;
                const requiredFields = ['sku', 'labelQty', 'cp', 'ship', 'wtAct', 'wtDecl', 'w', 'l', 'h', 'unit'];

                requiredFields.forEach(id => {
                    const field = document.getElementById(id);
                    if (!field.value.trim()) {
                        showFieldError(field, 'This field is required');
                        isValid = false;
                    } else if (isNaN(field.value) && id !== 'sku' && id !== 'unit') {
                        showFieldError(field, 'Must be a number');
                        isValid = false;
                    } else {
                        clearFieldError(field);
                    }
                });

                return isValid;
            }

            // Modify getFormData to use FormData for file upload
            function getFormData() {
                const formElement = document.getElementById('addProductForm');
                const formData = new FormData(formElement);

                // Add all fields as before
                formData.append('parent', document.getElementById('parent').value || '');
                formData.append('sku', document.getElementById('sku').value);

                // Build Values JSON
                const values = {
                    lp: document.getElementById('lp').value || null,
                    cp: document.getElementById('cp').value || null,
                    frght: document.getElementById('freght').value || null,
                    lps: document.getElementById('lps').value || null,
                    ship: document.getElementById('ship').value || null,
                    temu_ship: document.getElementById('temu_ship').value || null,
                    ebay2_ship: document.getElementById('ebay2_ship').value || null,
                    initial_quantity: document.getElementById('initial_quantity').value || null,
                    label_qty: document.getElementById('labelQty').value || null,
                    wt_act: document.getElementById('wtAct').value || null,
                    wt_decl: document.getElementById('wtDecl').value || null,
                    l: document.getElementById('l').value || null,
                    w: document.getElementById('w').value || null,
                    h: document.getElementById('h').value || null,
                    cbm: document.getElementById('cbm').value || null,
                    dc: document.getElementById('dc').value || null,
                    l2_url: document.getElementById('l2Url').value || null,
                    pcs_per_box: document.getElementById('pcbox').value || null,
                    l1: document.getElementById('l1').value || null,
                    b: document.getElementById('b').value || null,
                    h1: document.getElementById('h1').value || null,
                    weight: document.getElementById('weight').value || null,
                    msrp: document.getElementById('msrp').value || null,
                    map: document.getElementById('map').value || null,
                    status: document.getElementById('status').value || null,
                    unit: document.getElementById('unit').value || null,
                    upc: document.getElementById('upc').value || null,
                };

                formData.append('Values', JSON.stringify(values));
                // The image file is already included by <input name="image">

                return formData;
            }


            // Update parent options in datalist
            function updateParentOptions() {
                const parentOptions = document.getElementById('parentOptions');
                parentOptions.innerHTML = '';

                const parentSKUs = new Set();
                tableData.forEach(item => {
                    // Only add Parent values that do NOT contain 'PARENT'
                    if (item.Parent && !item.Parent.toUpperCase().includes('PARENT')) {
                        parentSKUs.add(item.Parent);
                    }
                });

                parentSKUs.forEach(sku => {
                    const option = document.createElement('option');
                    option.value = sku;
                    parentOptions.appendChild(option);
                });
            }

            // Setup edit buttons
            function setupEditButtons() {
                document.querySelectorAll('.edit-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const sku = this.getAttribute('data-sku');
                        const product = tableData.find(p => p.SKU === sku);
                        if (product) {
                            editProduct(product);
                        }
                    });
                });
            }

            // Edit product
            function editProduct(product) {
                const modal = new bootstrap.Modal(document.getElementById('addProductModal'));
                const saveBtn = document.getElementById('saveProductBtn');

                // Clone and replace the save button to prevent multiple event listeners
                const newSaveBtn = saveBtn.cloneNode(true);
                saveBtn.parentNode.replaceChild(newSaveBtn, saveBtn);

                newSaveBtn.setAttribute('data-original-sku', product.SKU || '');
                newSaveBtn.setAttribute('data-original-parent', product.Parent || '');
                newSaveBtn.innerHTML = '<i class="fas fa-save me-2"></i> Update Product';

                newSaveBtn.addEventListener('click', async function() {
                    if (!validateProductForm()) return;

                    const formData = getFormData();
                    formData.append('operation', 'update');
                    formData.append('original_sku', this.getAttribute('data-original-sku'));
                    formData.append('original_parent', this.getAttribute('data-original-parent'));

                    try {
                        const response = await fetch('/product_master/store', {
                            method: 'POST',
                            headers: {
                                'X-CSRF-TOKEN': csrfToken
                            },
                            body: formData
                        });
                        const data = await response.json();

                        if (!response.ok) {
                            throw new Error(data.message ||
                                `Server returned status ${response.status}`);
                        }

                        showAlert('success', 'Product updated successfully!');
                        modal.hide();
                        loadData();
                        resetProductForm();
                    } catch (error) {
                        showAlert('danger', error.message);
                    }
                });

                // Normalize status value to match select options exactly
                let normalizedStatus = '';
                switch ((product.status || '').toLowerCase()) {
                    case 'active':
                        normalizedStatus = 'active';
                        break;
                    case 'inactive':
                        normalizedStatus = 'inactive';
                        break;
                    case 'dc':
                        normalizedStatus = 'DC';
                        break;
                    case 'upcoming':
                        normalizedStatus = 'upcoming';
                        break;
                    case '2bdc':
                        normalizedStatus = '2BDC';
                        break;
                    default:
                        normalizedStatus = product.status || '';
                        break;
                }

                // Populate form fields (including disabled)
                const fields = {
                    sku: product.SKU || '',
                    parent: product.Parent || '',
                    labelQty: product.label_qty || '1',
                    cp: product.cp || '',
                    ship: product.ship || '',
                    temu_ship: product.temu_ship || '',
                    ebay2_ship: product.ebay2_ship || '',
                    initial_quantity: product.initial_quantity || '',
                    wtAct: product.wt_act || '',
                    wtDecl: product.wt_decl || '',
                    w: product.w || '',
                    l: product.l || '',
                    h: product.h || '',
                    l2Url: product.l2_url || '',
                    pcbox: product.pcs_per_box || '',
                    l1: product.l1 || '',
                    b: product.b || '',
                    h1: product.h1 || '',
                    upc: product.upc || '',
                    unit: product.unit || '',
                 
                    status: normalizedStatus,
                    cbm: product.cbm || '',
                    dc: product.dc || '',
                    weight: product.weight || '',
                    msrp: product.msrp || '',
                    map: product.map || ''
                };

                Object.entries(fields).forEach(([id, value]) => {
                    const element = document.getElementById(id);
                    if (element) element.value = value;
                });

                // Show image preview if image_path exists
                const imagePreview = document.getElementById('imagePreview');
                if (product.image_path) {
                    imagePreview.innerHTML =
                        `<img src="${product.image_path}" alt="Product Image" style="max-width:120px;max-height:120px;border-radius:8px;">`;
                } else {
                    imagePreview.innerHTML = '<span class="text-muted">No image</span>';
                }

                // Calculate derived fields
                calculateCBM();
                calculateLP();
                modal.show();
            }

            // Reset the product form
            function resetProductForm() {
                document.getElementById('addProductForm').reset();
                document.querySelectorAll('.is-invalid').forEach(el => el.classList.remove('is-invalid'));
                document.querySelectorAll('.invalid-feedback').forEach(el => el.textContent = '');
                document.getElementById('form-errors').innerHTML = '';
            }

            // Initialize progress modal
            function setupProgressModal() {
                const progressModal = new bootstrap.Modal(document.getElementById('progressModal'));
                const cancelUploadBtn = document.getElementById('cancelUploadBtn');
                const doneBtn = document.getElementById('doneBtn');
                let uploadInProgress = false;
                let currentUpload = null;

                cancelUploadBtn.addEventListener('click', function() {
                    if (uploadInProgress && currentUpload) {
                        currentUpload.abort();
                    }
                    progressModal.hide();
                });

                doneBtn.addEventListener('click', function() {
                    progressModal.hide();
                });

                window.showUploadProgress = function(sheets) {
                    const progressContainer = document.getElementById('progress-container');
                    const errorContainer = document.getElementById('error-container');

                    progressContainer.innerHTML = '';
                    errorContainer.innerHTML = '';
                    document.getElementById('success-alert').style.display = 'none';
                    doneBtn.style.display = 'none';
                    cancelUploadBtn.disabled = false;
                    uploadInProgress = true;

                    sheets.forEach(sheet => {
                        progressContainer.innerHTML += `
                            <div class="progress-item mb-3" id="${sheet.id}-container">
                                <h6 class="d-flex align-items-center">
                                    <i class="fas fa-file-excel text-primary me-2"></i>
                                    ${sheet.displayName}
                                    <span id="${sheet.id}-icon" class="ms-auto">
                                        <i class="fas fa-circle-notch fa-spin"></i>
                                    </span>
                                </h6>
                                <div class="progress">
                                    <div id="${sheet.id}-progress" class="progress-bar progress-bar-striped progress-bar-animated" 
                                        role="progressbar" style="width: 0%"></div>
                                </div>
                                <div id="${sheet.id}-status" class="small text-muted mt-1">Initializing...</div>
                                <div id="${sheet.id}-error" class="small text-danger mt-1"></div>
                            </div>
                        `;
                    });

                    progressModal.show();
                };

                window.updateUploadProgress = function(sheetId, progress, status, isSuccess, errorMessage) {
                    const progressEl = document.getElementById(`${sheetId}-progress`);
                    const statusEl = document.getElementById(`${sheetId}-status`);
                    const iconEl = document.getElementById(`${sheetId}-icon`);
                    const errorEl = document.getElementById(`${sheetId}-error`);

                    if (progressEl && statusEl && iconEl) {
                        progressEl.style.width = `${progress}%`;

                        if (isSuccess) {
                            progressEl.classList.remove('progress-bar-animated');
                            progressEl.classList.add('bg-success');
                            statusEl.textContent = status || 'Completed successfully';
                            statusEl.classList.add('text-success');
                            iconEl.innerHTML = '<i class="fas fa-check-circle text-success"></i>';
                        } else if (progress === 100) {
                            progressEl.classList.remove('progress-bar-animated');
                            progressEl.classList.add('bg-danger');
                            statusEl.textContent = status || 'Failed';
                            statusEl.classList.add('text-danger');
                            iconEl.innerHTML = '<i class="fas fa-times-circle text-danger"></i>';

                            if (errorMessage) {
                                errorEl.textContent = errorMessage;
                                document.getElementById('error-container').innerHTML += `
                                    <div class="alert alert-danger py-2 mb-2">
                                        <i class="fas fa-exclamation-triangle me-2"></i>
                                        <strong>${sheetId} Error:</strong> ${errorMessage}
                                    </div>
                                `;
                            }
                        } else {
                            statusEl.textContent = status || 'Processing...';
                        }
                    }
                };

                window.completeUpload = function(successCount, totalCount) {
                    uploadInProgress = false;
                    cancelUploadBtn.disabled = true;

                    if (successCount === totalCount) {
                        document.getElementById('success-alert').style.display = 'block';
                        doneBtn.style.display = 'block';
                    } else {
                        document.getElementById('error-container').innerHTML += `
                            <div class="alert alert-warning mt-3">
                                <i class="fas fa-info-circle me-2"></i>
                                ${successCount}/${totalCount} sheets updated successfully
                            </div>
                        `;
                        doneBtn.style.display = 'block';
                    }
                };
            }

            // Setup selection mode functionality
            function setupSelectionMode() {
                const toggleButton = document.getElementById('toggleSelection');
                const checkboxColumn = document.querySelectorAll('.checkbox-column');
                const selectAllCheckbox = document.getElementById('selectAll');
                const selectionActions = document.getElementById('selectionActions');
                const selectionCount = selectionActions.querySelector('.selection-count');
                const cancelButton = document.getElementById('cancelSelection');
                let selectionMode = false;

                // Toggle selection mode
                toggleButton.addEventListener('click', function() {
                    selectionMode = !selectionMode;

                    if (!selectionMode) {
                        // Clear selections when turning off selection mode
                        selectedItems = {};
                        document.querySelectorAll('.checkbox-column').forEach(col => col.style.display = 'none'); // ✅ hide header col
                        document.querySelectorAll('.checkbox-cell').forEach(cell => cell.style.display = 'none'); // ✅ hide body cells
                        selectionActions.style.display = 'none';
                        this.innerHTML = '<i class="fas fa-plus"></i>';

                        // Re-render table without checkboxes
                        const currentFilters = getCurrentFilters();
                        let filteredData = applyFiltersToData(currentFilters);
                        renderTable(filteredData);
                    } else {
                        // ✅ Show header checkbox column
                        document.querySelectorAll('.checkbox-column').forEach(col => col.style.display = 'table-cell');

                        // ✅ Show existing row checkboxes
                        document.querySelectorAll('.checkbox-cell').forEach(cell => cell.style.display = 'table-cell');

                        selectionActions.style.display = 'block';
                        this.innerHTML = '<i class="fas fa-times"></i>';
                        addCheckboxesToRows();
                    }

                    updateSelectionCount();
                });

                // Cancel selection
                cancelButton.addEventListener('click', function() {
                    selectedItems = {};
                    selectionMode = false;
                    checkboxColumn.forEach(col => col.style.display = 'none');
                    selectionActions.style.display = 'none';
                    toggleButton.innerHTML = '<i class="fas fa-plus"></i>';

                    // Re-render table without checkboxes
                    const currentFilters = getCurrentFilters();
                    let filteredData = applyFiltersToData(currentFilters);
                    renderTable(filteredData);
                });

                // Handle individual checkbox clicks
                document.addEventListener('change', function(e) {
                    if (e.target && e.target.classList.contains('row-checkbox')) {
                        const sku = e.target.dataset.sku;
                        const id = e.target.dataset.id;

                        if (e.target.checked) {
                            selectedItems[sku] = {
                                id: id,
                                checked: true
                            };
                        } else {
                            delete selectedItems[sku];
                        }

                        updateSelectionCount();

                        // Update "select all" checkbox state
                        const visibleCheckboxes = document.querySelectorAll('.row-checkbox');
                        const allChecked = Array.from(visibleCheckboxes).every(cb => cb.checked);
                        selectAllCheckbox.checked = allChecked && visibleCheckboxes.length > 0;
                    }
                });
            }

            function updateSelectionCount() {
                const selectionActions = document.getElementById('selectionActions');
                if (!selectionActions) return;
                const selectionCount = selectionActions.querySelector('.selection-count');
                if (selectionCount) {
                    selectionCount.textContent = `${Object.keys(selectedItems).length} items selected`;
                }
            }
            function restoreSelectAllState() {
                const selectAllCheckbox = document.getElementById('selectAll');
                if (!selectAllCheckbox) return;

                const visibleCheckboxes = document.querySelectorAll('.row-checkbox');
                if (visibleCheckboxes.length === 0) {
                    selectAllCheckbox.checked = false;
                    return;
                }

                const allChecked = Array.from(visibleCheckboxes).every(cb => cb.checked);
                selectAllCheckbox.checked = allChecked;
            }

            function bindRowCheckboxes() {
                const checkboxes = document.querySelectorAll('.row-checkbox');
                checkboxes.forEach(checkbox => {
                    const sku = checkbox.dataset.sku;
                    const id = checkbox.dataset.id;

                    // Memory se restore state
                    if (selectedItems[sku]) {
                        checkbox.checked = true;
                    }

                    checkbox.addEventListener('change', function () {
                        if (this.checked) {
                            selectedItems[sku] = { id: id, checked: true };
                        } else {
                            delete selectedItems[sku];
                        }
                        updateSelectionCount();
                    });
                });
            }

            function bindSelectAllCheckbox() {
                const selectAllCheckbox = document.getElementById('selectAll');
                if (!selectAllCheckbox) return;

                selectAllCheckbox.addEventListener('change', function () {
                    const checkboxes = document.querySelectorAll('.row-checkbox');

                    if (this.checked) {
                        checkboxes.forEach(checkbox => {
                            checkbox.checked = true;
                            const sku = checkbox.dataset.sku;
                            const id = checkbox.dataset.id;
                            selectedItems[sku] = { id: id, checked: true };
                        });
                    } else {
                        checkboxes.forEach(checkbox => {
                            checkbox.checked = false;
                            const sku = checkbox.dataset.sku;
                            delete selectedItems[sku];
                        });
                    }

                    updateSelectionCount();
                });
            }


            // Add checkboxes to existing rows when entering selection mode
            function addCheckboxesToRows() {
                const rows = document.querySelectorAll('#table-body tr');
                rows.forEach(row => {
                    if (!row.querySelector('.row-checkbox')) {
                        const firstCell = row.cells[0];
                        const skuCell = row.cells[1] || row.cells[0];
                        const sku = skuCell.textContent.trim();

                        // Find the item in tableData to get the ID
                        const item = tableData.find(item => item.SKU === sku);
                        const id = item ? item.id : '';

                        const checkboxCell = document.createElement('td');
                        checkboxCell.className = 'checkbox-cell';
                        checkboxCell.style.textAlign = 'center';

                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.className = 'row-checkbox';
                        checkbox.dataset.sku = sku;
                        checkbox.dataset.id = id;

                        if (selectedItems[sku]) {
                            checkbox.checked = true;
                        }

                        checkboxCell.appendChild(checkbox);
                        row.insertBefore(checkboxCell, firstCell);
                    }
                });
            }

            // Get current filter values
            function getCurrentFilters() {
                return {
                    parent: document.getElementById('parentSearch').value.toLowerCase().trim(),
                    sku: document.getElementById('skuSearch').value.toLowerCase().trim(),
                    global: document.getElementById('customSearch').value.toLowerCase().trim()
                };
            }

            // Apply filters to data
            function applyFiltersToData(filters) {
                let filteredData = [...tableData];

                if (filters.parent) {
                    filteredData = filteredData.filter(item =>
                        (item.Parent || '').toLowerCase().includes(filters.parent)
                    );
                }

                if (filters.sku) {
                    filteredData = filteredData.filter(item =>
                        (item.SKU || '').toLowerCase().includes(filters.sku)
                    );
                }

                if (filters.global) {
                    filteredData = filteredData.filter(item =>
                        Object.values(item).some(value =>
                            String(value).toLowerCase().includes(filters.global)
                        )
                    );
                }

                return filteredData;
            }



            // Handle the batch processing of selected items
            function setupBatchProcessing() {
                const processSelectedModal = new bootstrap.Modal(document.getElementById('processSelectedModal'));
                const addFieldBtn = document.getElementById('addFieldBtn');
                const fieldOperations = document.getElementById('fieldOperations');
                const applyChangesBtn = document.getElementById('applyChangesBtn');
                const selectedItemCount = document.getElementById('selectedItemCount');
                const batchUpdateResult = document.getElementById('batchUpdateResult');

                // List of all available fields
                const allFields = [{
                        value: 'lp',
                        text: 'LP'
                    },
                    {
                        value: 'cp',
                        text: 'CP'
                    },
                    {
                        value: 'frght',
                        text: 'FRGHT'
                    },
                    {
                        value: 'ship',
                        text: 'SHIP'
                    },
                    {
                        value: 'temu_ship',
                        text: 'TEMU SHIP'
                    },
                    {
                        value: 'ebay2_ship',
                        text: 'EBAY2 SHIP'
                    },
                    {
                        value: 'initial_quantity',
                        text: 'INITIAL QUANTITY'
                    },
                    {
                        value: 'label_qty',
                        text: 'Label QTY'
                    },
                    {
                        value: 'wt_act',
                        text: 'WT ACT'
                    },
                    {
                        value: 'wt_decl',
                        text: 'WT DECL'
                    },
                    {
                        value: 'l',
                        text: 'L'
                    },
                    {
                        value: 'w',
                        text: 'W'
                    },
                    {
                        value: 'h',
                        text: 'H'
                    },
                    {
                        value: 'status',
                        text: 'Status'
                    }
                ];

                // Open modal when Process Selected is clicked
                document.getElementById('processSelected').addEventListener('click', function() {
                    const selectedCount = Object.keys(selectedItems).length;
                    if (selectedCount === 0) {
                        alert('No items selected');
                        return;
                    }

                    // Reset the form
                    resetBatchForm();

                    // Update selected count
                    selectedItemCount.textContent = selectedCount;

                    // Show the modal
                    processSelectedModal.show();
                });

                // Add a new field operation row
                addFieldBtn.addEventListener('click', function() {
                    addFieldRow();
                    updateFieldOptions();
                });

                // Handle remove field button clicks using event delegation
                fieldOperations.addEventListener('change', function(e) {
                    if (e.target.classList.contains('field-selector')) {
                        updateFieldOptions();

                        // If status is selected, change input to dropdown and hide operation
                        const selectedField = e.target.value;
                        const row = e.target.closest('.field-operation');
                        const valueInput = row.querySelector('.field-value');
                        const operationSelector = row.querySelector('.operation-selector');

                        if (selectedField === 'status') {
                            // Replace text input with status dropdown
                            valueInput.outerHTML = `
                            <select class="form-select field-value">
                                <option value="">Select Status</option>
                                <option value="active">Active</option>
                                <option value="inactive">Inactive</option>
                                <option value="DC">DC</option>
                                <option value="upcoming">Upcoming</option>
                                <option value="2BDC">2BDC</option>
                            </select>
                            `;

                            // Hide operation selector for status and set it to "set" (=)
                            operationSelector.style.display = 'none';
                            operationSelector.value = 'set';
                        } else {
                            // For non-status fields, show operation selector
                            operationSelector.style.display = '';

                            // Replace dropdown with text input if it's not status
                            if (valueInput.tagName === 'SELECT' && selectedField !== 'status') {
                                valueInput.outerHTML = `
                                <input type="text" class="form-control field-value" placeholder="Enter value">
                                `;
                            }
                        }
                    }
                });

                // Apply changes to selected items
                applyChangesBtn.addEventListener('click', async function() {
                    // Validate form

                    const operations = getFieldOperations();
                    if (operations.length === 0) {
                        showBatchResult('warning', 'Please select at least one field to update');
                        return;
                    }

                    // Prepare data for update with item IDs
                    const updateData = {
                        // Convert selectedItems to array of {sku, id} objects
                        items: Object.entries(selectedItems).map(([sku, item]) => {
                            // Find the full item data to get the ID
                            const fullItem = tableData.find(i => i.SKU === sku);
                            return {
                                sku: sku,
                                id: fullItem ? fullItem.id : null
                            };
                        }),
                        operations: operations
                    };

                    // Disable button and show loading
                    applyChangesBtn.disabled = true;
                    applyChangesBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';

                    try {
                        // Send to server
                        const response = await makeRequest('/product-master/batch-update', 'POST',
                            updateData);
                        const result = await response.json();

                        if (!response.ok) {
                            throw new Error(result.message || 'Failed to update products');
                        }



                        // Show success message
                        showBatchResult('success',
                            `Successfully updated ${Object.keys(selectedItems).length} products`);

                        // After successful update, refresh the table
                        setTimeout(() => {
                            processSelectedModal.hide();
                            loadData();
                        }, 1500);

                    } catch (error) {
                        console.error("Error during batch update:", error);
                        showBatchResult('danger', `Error: ${error.message}`);
                    } finally {
                        applyChangesBtn.disabled = false;
                        applyChangesBtn.innerHTML = 'Apply Changes';
                    }
                });

                // Helper function to add a new field row
                function addFieldRow() {
                    const row = document.createElement('div');
                    row.className = 'field-operation mb-3';
                    row.innerHTML = `
                        <div class="row g-2 align-items-center">
                            <div class="col-3">
                                <select class="form-select field-selector">
                                    <option value="">Select Field</option>
                                    ${generateFieldOptions([])}
                                </select>
                            </div>
                            <div class="col-3">
                                <select class="form-select operation-selector">
                                    <option value="set">=</option>
                                    <option value="add">+</option>
                                    <option value="subtract">-</option>
                                    <option value="multiply">×</option>
                                    <option value="divide">÷</option>
                                </select>
                            </div>
                            <div class="col-4">
                                <input type="text" class="form-control field-value" placeholder="Enter value">
                            </div>
                            <div class="col-2">
                                <button type="button" class="btn btn-outline-danger remove-field">
                                    <i class="fas fa-trash"></i>
                                </button>
                            </div>
                        </div>
                    `;
                    fieldOperations.appendChild(row);
                }

                // Helper function to reset the batch form
                function resetBatchForm() {
                    fieldOperations.innerHTML = '';
                    addFieldRow();
                    batchUpdateResult.innerHTML = '';
                }

                // Helper function to get field operations from the form
                function getFieldOperations() {
                    const operations = [];
                    const rows = fieldOperations.querySelectorAll('.field-operation');

                    rows.forEach(row => {
                        const field = row.querySelector('.field-selector').value;
                        const operation = row.querySelector('.operation-selector').value;
                        const value = row.querySelector('.field-value').value;

                        if (field && value) {
                            operations.push({
                                field: field,
                                operation: operation,
                                value: value
                            });
                        }
                    });

                    return operations;
                }

                // Helper function to show result messages
                function showBatchResult(type, message) {
                    batchUpdateResult.innerHTML = `
                        <div class="alert alert-${type} alert-dismissible fade show" role="alert">
                            ${message}
                            <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                        </div>
                    `;
                }

                // Helper function to generate field options HTML, excluding already selected fields
                function generateFieldOptions(excludeFields) {
                    return allFields
                        .filter(field => !excludeFields.includes(field.value))
                        .map(field => `<option value="${field.value}">${field.text}</option>`)
                        .join('');
                }

                // Helper function to update field options in all selectors
                function updateFieldOptions() {
                    const rows = fieldOperations.querySelectorAll('.field-operation');
                    const selectedFields = getSelectedFields();

                    rows.forEach(row => {
                        const selector = row.querySelector('.field-selector');
                        const currentValue = selector.value;

                        // Get fields to exclude (all selected fields except current one)
                        const fieldsToExclude = selectedFields.filter(field => field !== currentValue);

                        // Store current selection
                        const currentSelection = selector.value;

                        // Update options, excluding already selected fields
                        selector.innerHTML = `
                            <option value="">Select Field</option>
                            ${generateFieldOptions(fieldsToExclude)}
                        `;

                        // Restore current selection
                        selector.value = currentSelection;
                    });

                    // Disable/enable add field button based on available options
                    const unusedFields = allFields.filter(field => !selectedFields.includes(field.value));
                    addFieldBtn.disabled = unusedFields.length === 0;
                }

                // Helper function to get currently selected fields
                function getSelectedFields() {
                    const selectedFields = [];
                    const selectors = fieldOperations.querySelectorAll('.field-selector');

                    selectors.forEach(selector => {
                        if (selector.value) {
                            selectedFields.push(selector.value);
                        }
                    });

                    return selectedFields;
                }
            }

            // Initialize import from API functionality
            document.getElementById('importFromApiBtn').addEventListener('click', function() {
                const importBtn = this;
                importBtn.disabled = true;
                importBtn.innerHTML = '<i class="fas fa-spinner fa-spin me-1"></i> Importing...';

                makeRequest('/product-master-data-view', 'GET')
                    .then(res => {
                        if (!res.ok) {
                            throw new Error(`HTTP error! status: ${res.status}`);
                        }
                        return res.json();
                    })
                    .then(apiResponse => {
                        if (apiResponse && apiResponse.data) {
                            return makeRequest('/product-master/import-from-sheet', 'POST', {
                                data: apiResponse.data
                            });
                        }
                        throw new Error('Failed to fetch API data.');
                    })
                    .then(res => {
                        if (!res.ok) {
                            throw new Error(`HTTP error! status: ${res.status}`);
                        }
                        return res.json();
                    })
                    .then(result => {
                        alert(
                            `Import complete!\nImported: ${result.imported ?? 0}\nErrors: ${result.errors?.length ? result.errors.join('\n') : 'None'}`
                        );
                        loadData(); // Refresh the table after import
                    })
                    .catch(err => {
                        console.error('Import failed:', err);
                        alert('Import failed: ' + err.message);
                    })
                    .finally(() => {
                        importBtn.disabled = false;
                        importBtn.innerHTML =
                            '<i class="fas fa-cloud-download-alt me-1"></i> Import from API Sheet';
                    });
            });

            // Utility functions
            function escapeHtml(str) {
                if (!str) return '';
                return String(str)
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            }

            function formatNumber(num, decimals) {
                if (num === undefined || num === null) return '-';
                const n = parseFloat(num);
                return isNaN(n) ? '-' : n.toFixed(decimals);
            }

            function debounce(func, wait) {
                let timeout;
                return function() {
                    const context = this,
                        args = arguments;
                    clearTimeout(timeout);
                    timeout = setTimeout(() => func.apply(context, args), wait);
                };
            }

            function showError(message) {
                document.getElementById('rainbow-loader').innerHTML = `
                    <div class="alert alert-danger">
                        <i class="fas fa-exclamation-triangle me-2"></i>
                        ${escapeHtml(message)}
                    </div>
                `;
            }

            function showAlert(type, message) {
                const alert = document.createElement('div');
                alert.className = `alert alert-${type} alert-dismissible fade show`;
                alert.innerHTML = `
                    ${message}
                    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
                `;

                const container = document.getElementById('form-errors');
                container.innerHTML = '';
                container.appendChild(alert);
            }

            function showFieldError(field, message) {
                const formGroup = field.closest('.form-group');
                if (!formGroup) return;

                let errorElement = formGroup.querySelector('.invalid-feedback');
                if (!errorElement) {
                    errorElement = document.createElement('div');
                    errorElement.className = 'invalid-feedback';
                    formGroup.appendChild(errorElement);
                }

                field.classList.add('is-invalid');
                errorElement.textContent = message;
            }

            function clearFieldError(field) {
                const formGroup = field.closest('.form-group');
                if (!formGroup) return;

                const errorElement = formGroup.querySelector('.invalid-feedback');
                if (errorElement) {
                    field.classList.remove('is-invalid');
                    errorElement.textContent = '';
                }
            }

            function setupDeleteButtons() {
                document.querySelectorAll('.delete-btn').forEach(btn => {
                    btn.addEventListener('click', function() {
                        const id = this.getAttribute('data-id');
                        const sku = this.getAttribute('data-sku');
                        // Use Bootstrap modal for confirmation instead of window.confirm
                        const confirmModal = document.createElement('div');
                        confirmModal.innerHTML = `
                            <div class="modal fade" id="deleteConfirmModal" tabindex="-1" aria-labelledby="deleteConfirmModalLabel" aria-hidden="true">
                            <div class="modal-dialog modal-dialog-centered">
                                <div class="modal-content border-0" style="border-radius: 18px; overflow: hidden;">
                                <div class="modal-header" style="background: linear-gradient(135deg, #ff6b6b 0%, #ff0000 100%); color: #fff;">
                                    <div class="d-flex align-items-center w-100">
                                    <div class="me-3" style="font-size: 2.5rem;">
                                        <i class="fas fa-exclamation-triangle fa-shake"></i>
                                    </div>
                                    <div>
                                        <h5 class="modal-title mb-0" id="deleteConfirmModalLabel" style="font-weight: 800; letter-spacing: 1px;">
                                        Delete Product?
                                        </h5>
                                        <small class="text-white-50">This action cannot be undone!</small>
                                    </div>
                                    <button type="button" class="btn-close btn-close-white ms-auto" data-bs-dismiss="modal" aria-label="Close"></button>
                                    </div>
                                </div>
                                <div class="modal-body text-center py-4">
                                    <div class="mb-3" style="font-size: 1.2rem;">
                                    Are you sure you want to <span class="fw-bold text-danger">delete</span> product<br>
                                    <span class="badge bg-danger fs-6 px-3 py-2 mt-2" style="font-size:1.1rem;">SKU: ${escapeHtml(sku)}</span>?
                                    </div>
                                    <div class="mb-2 text-warning" style="font-size: 1rem;">
                                    <i class="fas fa-info-circle me-1"></i>
                                    This will move the product to trash (soft delete).
                                    </div>
                                </div>
                                <div class="modal-footer justify-content-center" style="background: #fff;">
                                    <button type="button" class="btn btn-outline-secondary px-4" data-bs-dismiss="modal">
                                    <i class="fas fa-times me-1"></i>Cancel
                                    </button>
                                    <button type="button" class="btn btn-danger px-4" id="confirmDeleteBtn">
                                    <i class="fas fa-trash me-1"></i>Yes, Delete
                                    </button>
                                </div>
                                </div>
                            </div>
                            </div>
                        `;
                        document.body.appendChild(confirmModal);

                        const modal = new bootstrap.Modal(confirmModal.querySelector(
                            '#deleteConfirmModal'));
                        modal.show();

                        confirmModal.querySelector('#confirmDeleteBtn').addEventListener('click',
                            () => {
                                makeRequest('/product_master/delete', 'DELETE', {
                                        ids: [id]
                                    })
                                    .then(response => response.json())
                                    .then(data => {
                                        if (data.success) {
                                            showToast('success', data.message ||
                                                'Product deleted successfully!');
                                            loadData();
                                        } else {
                                            showToast('danger', data.message ||
                                                'Delete failed');
                                        }
                                    })
                                    .catch(() => {
                                        showToast('danger', 'Delete failed');
                                    })
                                    .finally(() => {
                                        modal.hide();
                                        setTimeout(() => confirmModal.remove(), 500);
                                    });
                            });

                        // Remove modal from DOM after hiding
                        confirmModal.querySelectorAll('[data-bs-dismiss="modal"]').forEach(btn => {
                            btn.addEventListener('click', () => setTimeout(() =>
                                confirmModal.remove(), 500));
                        });
                    });
                });
            }

            // Add this helper function for toast notifications (place it with your other utility functions)
            function showToast(type, message) {
                // Remove any existing toast
                document.querySelectorAll('.custom-toast').forEach(t => t.remove());

                const toast = document.createElement('div');
                toast.className =
                    `custom-toast toast align-items-center text-bg-${type} border-0 show position-fixed top-0 end-0 m-4`;
                toast.style.zIndex = 2000;
                toast.setAttribute('role', 'alert');
                toast.setAttribute('aria-live', 'assertive');
                toast.setAttribute('aria-atomic', 'true');
                toast.innerHTML = `
                    <div class="d-flex">
                        <div class="toast-body">${escapeHtml(message)}</div>
                        <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="Close"></button>
                    </div>
                `;
                document.body.appendChild(toast);

                setTimeout(() => {
                    toast.classList.remove('show');
                    setTimeout(() => toast.remove(), 500);
                }, 3000);

                setTimeout(() => {
                    toast.classList.remove('show');
                    setTimeout(() => toast.remove(), 500);
                }, 3000);

                toast.querySelector('[data-bs-dismiss="toast"]').onclick = () => {
                    toast.classList.remove('show');
                    setTimeout(() => toast.remove(), 500);
                };
            }

            document.addEventListener('mouseover', function(e) {
                const target = e.target.closest('.sku-hover');
                const tooltip = document.getElementById('skuImageTooltip');
                if (target && tooltip) {
                    const image = target.getAttribute('data-image');
                    if (image) {
                        tooltip.innerHTML = `<img src="${image}" alt="Product Image">`;
                        tooltip.style.display = 'block';
                    } else {
                        tooltip.style.display = 'none';
                    }
                }
            });
            document.addEventListener('mousemove', function(e) {
                const tooltip = document.getElementById('skuImageTooltip');
                if (tooltip && tooltip.style.display === 'block') {
                    tooltip.style.left = (e.pageX + 20) + 'px';
                    tooltip.style.top = (e.pageY + 10) + 'px';
                }
            });
            document.addEventListener('mouseout', function(e) {
                const target = e.target.closest('.sku-hover');
                const tooltip = document.getElementById('skuImageTooltip');
                if (target && tooltip) {
                    tooltip.style.display = 'none';
                }
            });
        });
    </script>
@endsection
