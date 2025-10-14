@extends('layouts.vertical', ['title' => 'Amazon - Ad Running', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
    <link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
    <style>
        .tabulator .tabulator-header {
            background: linear-gradient(90deg, #D8F3F3 0%, #D8F3F3 100%);
            border-bottom: 1px solid #403f3f;
            box-shadow: 0 4px 16px rgba(37, 99, 235, 0.10);
        }

        .tabulator .tabulator-header .tabulator-col {
            text-align: center;
            background: #D8F3F3;
            border-right: 1px solid #262626;
            padding: 16px 10px;
            font-weight: 700;
            color: #1e293b;
            font-size: 1.08rem;
            letter-spacing: 0.02em;
            transition: background 0.2s;
        }

        .tabulator .tabulator-header .tabulator-col:hover {
            background: #D8F3F3;
            color: #2563eb;
        }

        .tabulator-row {
            background-color: #fff !important;
            transition: background 0.18s;
        }

        .tabulator .tabulator-cell {
            text-align: center;
            padding: 14px 10px;
            border-right: 1px solid #262626;
            border-bottom: 1px solid #262626;
            font-size: 1rem;
            color: #22223b;
            vertical-align: middle;
            transition: background 0.18s, color 0.18s;
        }

        .tabulator .tabulator-cell:focus {
            outline: 1px solid #262626;
            background: #e0eaff;
        }

        #account-health-master .tabulator {
            border-radius: 18px;
            box-shadow: 0 6px 24px rgba(37, 99, 235, 0.13);
            overflow: hidden;
            border: 1px solid #e5e7eb;
        }

        .tabulator .tabulator-row .tabulator-cell:last-child,
        .tabulator .tabulator-header .tabulator-col:last-child {
            border-right: none;
        }

        .tabulator .tabulator-footer {
            background: #f4f7fa;
            border-top: 1px solid #262626;
            font-size: 1rem;
            color: #4b5563;
            padding: 5px;
            height: 100px;
        }

        .tabulator .tabulator-footer:hover {
            background: #e0eaff;
        }

        @media (max-width: 768px) {

            .tabulator .tabulator-header .tabulator-col,
            .tabulator .tabulator-cell {
                padding: 8px 2px;
                font-size: 0.95rem;
            }
        }

        /* Pagination styling */
        .tabulator .tabulator-footer .tabulator-paginator .tabulator-page {
            padding: 8px 16px;
            margin: 0 4px;
            border-radius: 6px;
            font-size: 0.95rem;
            font-weight: 500;
            transition: all 0.2s;
        }

        .tabulator .tabulator-footer .tabulator-paginator .tabulator-page:hover {
            background: #e0eaff;
            color: #2563eb;
        }

        .tabulator .tabulator-footer .tabulator-paginator .tabulator-page.active {
            background: #2563eb;
            color: white;
        }

        .parent-row-bg{
            background-color: #c3efff !important;
        }

        .green-bg {
            color: #05bd30 !important;
        }

        .pink-bg {
            color: #ff01d0 !important;
        }

        .red-bg {
            color: #ff2727 !important;
        }
    </style>
@endsection
@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Amazon - Budget',
        'sub_title' => 'Amazon - Budget',
    ])
    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body py-3">
                    <div class="mb-4">
                        <!-- Title -->
                        <h4 class="fw-bold text-primary mb-3 d-flex align-items-center">
                            <i class="fa-solid fa-chart-line me-2"></i>
                            Ad Running
                        </h4>

                        <!-- Filters Row -->
                        <div class="row g-3 mb-3">
                            <!-- Inventory Filters -->
                            <div class="col-md-6">
                                <div class="d-flex gap-2">
                                    <select id="inv-filter" class="form-select form-select-md">
                                        <option value="">Select INV</option>
                                        <option value="ALL">ALL</option>
                                        <option value="INV_0">0 INV</option>
                                        <option value="OTHERS">OTHERS</option>
                                    </select>

                                    <select id="nrl-filter" class="form-select form-select-md d-none">
                                        <option value="">Select NRL</option>
                                        <option value="NRL">NRL</option>
                                        <option value="RL">RL</option>
                                    </select>

                                    <select id="nra-filter" class="form-select form-select-md d-none">
                                        <option value="">Select NRA</option>
                                        <option value="NRA">NRA</option>
                                        <option value="RA">RA</option>
                                        <option value="LATER">LATER</option>
                                    </select>

                                    <select id="fba-filter" class="form-select form-select-md d-none">
                                        <option value="">Select FBA</option>
                                        <option value="FBA">FBA</option>
                                        <option value="FBM">FBM</option>
                                        <option value="BOTH">BOTH</option>
                                    </select>

                                </div>
                            </div>

                            <!-- Stats -->
                            <div class="col-md-6">
                                <div class="d-flex gap-2 justify-content-end">
                                    <a href="javascript:void(0)" id="export-btn" class="btn btn-sm btn-success d-flex align-items-center justify-content-center">
                                        <i class="fas fa-file-export me-1"></i> Export Excel/CSV
                                    </a>
                                    <button class="btn btn-success btn-md">
                                        <i class="fa fa-arrow-up me-1"></i>
                                        Need to increase bids: <span id="total-campaigns" class="fw-bold ms-1 fs-4">0</span>
                                    </button>
                                    <button class="btn btn-primary btn-md">
                                        <i class="fa fa-percent me-1"></i>
                                        of Total: <span id="percentage-campaigns" class="fw-bold ms-1 fs-4">0%</span>
                                    </button>
                                </div>
                            </div>
                        </div>

                        <!-- Search and Controls Row -->
                        <div class="row g-3">
                            <div class="col-md-6">
                                <div class="d-flex gap-2">
                                    <div class="input-group">
                                        <input type="text" id="global-search" class="form-control form-control-md"
                                            placeholder="Search campaign...">
                                    </div>
                                    <select id="status-filter" class="form-select form-select-md" style="width: 140px;">
                                        <option value="">All Status</option>
                                        <option value="ENABLED">Enabled</option>
                                        <option value="PAUSED">Paused</option>
                                        <option value="ARCHIVED">Archived</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Table Section -->
                    <div id="budget-under-table"></div>
                </div>
            </div>
        </div>
    </div>
@endsection

@section('script')
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <!-- SheetJS for Excel Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script>
        document.addEventListener("DOMContentLoaded", function() {

            const getDilColor = (value) => {
                const percent = parseFloat(value) * 100;
                if (percent < 16.66) return 'red';
                if (percent >= 16.66 && percent < 25) return 'yellow';
                if (percent >= 25 && percent < 50) return 'green';
                return 'pink';
            };

            var table = new Tabulator("#budget-under-table", {
                index: "sku",
                ajaxURL: "/amazon/ad-running/data",
                layout: "fitDataFill",
                movableColumns: true,
                resizableColumns: true,
                height: "700px",             
                virtualDom: true,
                rowFormatter: function(row) {
                    const data = row.getData();
                    const sku = (data.sku || "").toLowerCase().trim();

                    if (sku.includes("parent ")) {
                        row.getElement().classList.add("parent-row-bg");
                    }
                },
                columns: [
                    {
                        title: "Parent",
                        field: "parent"
                    },
                    {
                        title: "SKU",
                        field: "sku",
                        formatter: function(cell) {
                            let sku = cell.getValue();
                            return `
                                <span>${sku}</span>
                                <i class="fa fa-info-circle text-primary toggle-cols-btn" 
                                data-sku="${sku}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "INV",
                        field: "INV",
                        visible: false
                    },
                    {
                        title: "OV L30",
                        field: "L30",
                        visible: false
                    },
                    {
                        title: "DIL %",
                        field: "DIL %",
                        formatter: function(cell) {
                            const data = cell.getData();
                            const l30 = parseFloat(data.L30);
                            const inv = parseFloat(data.INV);

                            if (!isNaN(l30) && !isNaN(inv) && inv !== 0) {
                                const dilDecimal = (l30 / inv);
                                const color = getDilColor(dilDecimal);
                                return `<div class="text-center"><span class="dil-percent-value ${color}">${Math.round(dilDecimal * 100)}%</span></div>`;
                            }
                            return `<div class="text-center"><span class="dil-percent-value red">0%</span></div>`;
                        },
                        visible: false
                    },
                    {
                        title: "AL 30",
                        field: "A_L30",
                        visible: false
                    },
                    {
                        title: "A DIL %",
                        field: "A DIL %",
                        formatter: function(cell) {
                            const data = cell.getData();
                            const al30 = parseFloat(data.A_L30);
                            const inv = parseFloat(data.INV);

                            if (!isNaN(al30) && !isNaN(inv) && inv !== 0) {
                                const dilDecimal = (al30 / inv);
                                const color = getDilColor(dilDecimal);
                                return `<div class="text-center"><span class="dil-percent-value ${color}">${Math.round(dilDecimal * 100)}%</span></div>`;
                            }
                            return `<div class="text-center"><span class="dil-percent-value red">0%</span></div>`;
                        },
                        visible: false
                    },
                    {
                        title: "NRL",
                        field: "NRL",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue();

                            let bgColor = "";
                            if (value === "NRL") {
                                bgColor = "background-color:#dc3545;color:#fff;"; // red
                            } else if (value === "RL") {
                                bgColor = "background-color:#28a745;color:#fff;"; // green
                            }

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="NRL"
                                        style="width: 90px; ${bgColor}">
                                    <option value="RL" ${value === 'RL' ? 'selected' : ''}>RL</option>
                                    <option value="NRL" ${value === 'NRL' ? 'selected' : ''}>NRL</option>
                                </select>
                            `;
                        },
                        visible: false,
                        hozAlign: "center"
                    },
                    {
                        title: "NRA",
                        field: "NRA",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue()?.trim();

                            let bgColor = "";
                            if (value === "NRA") {
                                bgColor = "background-color:#dc3545;color:#fff;"; // red
                            } else if (value === "RA") {
                                bgColor = "background-color:#28a745;color:#fff;"; // green
                            } else if (value === "LATER") {
                                bgColor = "background-color:#ffc107;color:#000;"; // yellow
                            }

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="NRA"
                                        style="width: 100px; ${bgColor}">
                                    <option value="RA" ${value === 'RA' ? 'selected' : ''}>RA</option>
                                    <option value="NRA" ${value === 'NRA' ? 'selected' : ''}>NRA</option>
                                    <option value="LATER" ${value === 'LATER' ? 'selected' : ''}>LATER</option>
                                </select>
                            `;
                        },
                        hozAlign: "center",
                        visible: false
                    },
                    {
                        title: "FBA",
                        field: "FBA",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue();

                            let bgColor = "";
                            if (value === "FBA") {
                                bgColor = "background-color:#007bff;color:#fff;"; // blue
                            } else if (value === "FBM") {
                                bgColor = "background-color:#6f42c1;color:#fff;"; // purple
                            } else if (value === "BOTH") {
                                bgColor = "background-color:#90ee90;color:#000;"; // light green
                            }

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="FBA"
                                        style="width: 90px; ${bgColor}">
                                    <option value="FBA" ${value === 'FBA' ? 'selected' : ''}>FBA</option>
                                    <option value="FBM" ${value === 'FBM' ? 'selected' : ''}>FBM</option>
                                    <option value="BOTH" ${value === 'BOTH' ? 'selected' : ''}>BOTH</option>
                                </select>
                            `;
                        },
                        hozAlign: "center",
                        visible: false
                    },
                    {
                        title: "SPEND L30",
                        field: "SPEND_L30",
                        formatter: function(cell) {
                            let SPEND_L30 = cell.getValue();
                            return `
                                <span>${SPEND_L30.toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-spendL30-btn" 
                                data-spend-l30="${SPEND_L30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Spend L30",
                        field: "kw_spend_L30",
                        visible: false,
                        formatter: function(cell) {
                            let KW_SPEND_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_SPEND_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Spend L30",
                        field: "pt_spend_L30",
                        visible: false,
                        formatter: function(cell) {
                            let PT_SPEND_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_SPEND_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Spend L30",
                        field: "hl_spend_L30",
                        visible: false,
                        formatter: function(cell) {
                            let HL_SPEND_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_SPEND_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "SPEND L7",
                        field: "SPEND_L7",
                        formatter: function(cell) {
                            let SPEND_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(SPEND_L7).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-spendL7-btn" 
                                data-spend-l7="${SPEND_L7}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Spend L7",
                        field: "kw_spend_L7",
                        visible: false,
                        formatter: function(cell) {
                            let KW_SPEND_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_SPEND_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Spend L7",
                        field: "pt_spend_L7",
                        visible: false,
                        formatter: function(cell) {
                            let PT_SPEND_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_SPEND_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Spend L7",
                        field: "hl_spend_L7",
                        visible: false,
                        formatter: function(cell) {
                            let HL_SPEND_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_SPEND_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "SOLD L30",
                        field: "SOLD_L30",
                        formatter: function(cell) {
                            let SOLD_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(SOLD_L30).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-soldL30-btn" 
                                data-sold-l30="${SOLD_L30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Sold L30",
                        field: "kw_sold_L30",
                        visible: false,
                        formatter: function(cell) {
                            let KW_SOLD_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_SOLD_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Sold L30",
                        field: "pt_sold_L30",
                        visible: false,
                        formatter: function(cell) {
                            let PT_SOLD_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_SOLD_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Sold L30",
                        field: "hl_sold_L30",
                        visible: false,
                        formatter: function(cell) {
                            let HL_SOLD_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_SOLD_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "SOLD L7",
                        field: "SOLD_L7",
                        formatter: function(cell) {
                            let SOLD_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(SOLD_L7).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-soldL7-btn" 
                                data-sold-l7="${SOLD_L7}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Sold L7",
                        field: "kw_sold_L7",
                        visible: false,
                        formatter: function(cell) {
                            let KW_SOLD_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_SOLD_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Sold L7",
                        field: "pt_sold_L7",
                        visible: false,
                        formatter: function(cell) {
                            let PT_SOLD_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_SOLD_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Sold L7",
                        field: "hl_sold_L7",
                        visible: false,
                        formatter: function(cell) {
                            let HL_SOLD_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_SOLD_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "SALES L30",
                        field: "SALES_L30",
                        formatter: function(cell) {
                            let SALES_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(SALES_L30).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-salesL30-btn" 
                                data-sales-l30="${SALES_L30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Sales L30",
                        field: "kw_sales_L30",
                        visible: false,
                        formatter: function(cell) {
                            let KW_SALES_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_SALES_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Sales L30",
                        field: "pt_sales_L30",
                        visible: false,
                        formatter: function(cell) {
                            let PT_SALES_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_SALES_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Sales L30",
                        field: "hl_sales_L30",
                        visible: false,
                        formatter: function(cell) {
                            let HL_SALES_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_SALES_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "SALES L7",
                        field: "SALES_L7",
                        formatter: function(cell) {
                            let SALES_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(SALES_L7).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-salesL7-btn" 
                                data-sales-l30="${SALES_L7}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Sales L7",
                        field: "kw_sales_L7",
                        visible: false,
                        formatter: function(cell) {
                            let KW_SALES_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_SALES_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Sales L7",
                        field: "pt_sales_L7",
                        visible: false,
                        formatter: function(cell) {
                            let PT_SALES_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_SALES_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Sales L7",
                        field: "hl_sales_L7",
                        visible: false,
                        formatter: function(cell) {
                            let HL_SALES_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_SALES_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "CLICKS L30",
                        field: "CLICKS_L30",
                        formatter: function(cell) {
                            let CLICKS_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(CLICKS_L30).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-clicksL30-btn" 
                                data-clicks-l30="${CLICKS_L30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Clicks L30",
                        field: "kw_clicks_L30",
                        visible: false,
                        formatter: function(cell) {
                            let KW_CLICKS_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_CLICKS_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Clicks L30",
                        field: "pt_clicks_L30",
                        visible: false,
                        formatter: function(cell) {
                            let PT_CLICKS_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_CLICKS_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Clicks L30",
                        field: "hl_clicks_L30",
                        visible: false,
                        formatter: function(cell) {
                            let HL_CLICKS_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_CLICKS_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "CLICKS L7",
                        field: "CLICKS_L7",
                        formatter: function(cell) {
                            let CLICKS_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(CLICKS_L7).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-clicksL7-btn" 
                                data-clicks-l7="${CLICKS_L7}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW Clicks L7",
                        field: "kw_clicks_L7",
                        visible: false,
                        formatter: function(cell) {
                            let KW_CLICKS_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_CLICKS_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT Clicks L7",
                        field: "pt_clicks_L7",
                        visible: false,
                        formatter: function(cell) {
                            let PT_CLICKS_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_CLICKS_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL Clicks L7",
                        field: "hl_clicks_L7",
                        visible: false,
                        formatter: function(cell) {
                            let HL_CLICKS_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_CLICKS_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "IMP L30",
                        field: "IMP_L30",
                        formatter: function(cell) {
                            let IMP_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(IMP_L30).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-impL30-btn" 
                                data-clicks-l7="${IMP_L30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW IMP L30",
                        field: "kw_impr_L30",
                        visible: false,
                        formatter: function(cell) {
                            let KW_IMP_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_IMP_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT IMP L30",
                        field: "pt_impr_L30",
                        visible: false,
                        formatter: function(cell) {
                            let PT_IMP_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_IMP_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL IMP L30",
                        field: "hl_impr_L30",
                        visible: false,
                        formatter: function(cell) {
                            let HL_IMP_L30 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_IMP_L30).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "IMP L7",
                        field: "IMP_L7",
                        formatter: function(cell) {
                            let IMP_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(IMP_L7).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary toggle-impL7-btn" 
                                data-clicks-l7="${IMP_L7}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "KW IMP L7",
                        field: "kw_impr_L7",
                        visible: false,
                        formatter: function(cell) {
                            let KW_IMP_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(KW_IMP_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "PT IMP L7",
                        field: "pt_impr_L7",
                        visible: false,
                        formatter: function(cell) {
                            let PT_IMP_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(PT_IMP_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "HL IMP L7",
                        field: "hl_impr_L7",
                        visible: false,
                        formatter: function(cell) {
                            let HL_IMP_L7 = cell.getValue() || 0;
                            return `
                                <span>${parseFloat(HL_IMP_L7).toFixed(0)}</span>
                            `;
                        }
                    },
                    {
                        title: "START AD",
                        field: "start_ad",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue();

                            let bgColor = "";
                            if (value === "KW") {
                                bgColor = "background-color:#28a745;color:#fff;";
                            } else if (value === "HL") {
                                bgColor = "background-color:#28a745;color:#fff;";
                            } else if (value === "PT") {
                                bgColor = "background-color:#28a745;color:#000;";
                            } else if (value === "ALL") {
                                bgColor = "background-color:#28a745;color:#000;";
                            }

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="start_ad"
                                        style="${bgColor}">
                                    <option value=""></option>
                                    <option value="KW" ${value === 'KW' ? 'selected' : ''}>KW</option>
                                    <option value="HL" ${value === 'HL' ? 'selected' : ''}>HL</option>
                                    <option value="PT" ${value === 'PT' ? 'selected' : ''}>PT</option>
                                    <option value="ALL" ${value === 'ALL' ? 'selected' : ''}>ALL</option>
                                </select>
                            `;
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "STOP AD",
                        field: "stop_ad",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue();

                            let bgColor = "";
                            if (value === "KW") {
                                bgColor = "background-color:#dc3545;color:#fff;";
                            } else if (value === "HL") {
                                bgColor = "background-color:#dc3545;color:#fff;";
                            } else if (value === "PT") {
                                bgColor = "background-color:#dc3545;color:#000;";
                            } else if (value === "ALL") {
                                bgColor = "background-color:#dc3545;color:#000;";
                            }

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="stop_ad"
                                        style="${bgColor}">
                                    <option value=""></option>
                                    <option value="KW" ${value === 'KW' ? 'selected' : ''}>KW</option>
                                    <option value="HL" ${value === 'HL' ? 'selected' : ''}>HL</option>
                                    <option value="PT" ${value === 'PT' ? 'selected' : ''}>PT</option>
                                    <option value="ALL" ${value === 'ALL' ? 'selected' : ''}>ALL</option>
                                </select>
                            `;
                        },
                        hozAlign: "center"
                    },

                ],
                ajaxResponse: function(url, params, response) {
                    return response.data;
                }
            });

            document.addEventListener("change", function(e){
                if(e.target.classList.contains("editable-select")){
                    let sku   = e.target.getAttribute("data-sku");
                    let field = e.target.getAttribute("data-field");
                    let value = e.target.value;

                    fetch('/update-amazon-nr-nrl-fba', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'X-CSRF-TOKEN': document.querySelector('meta[name="csrf-token"]').getAttribute('content')
                        },
                        body: JSON.stringify({
                            sku: sku,
                            field: field,
                            value: value
                        })
                    })
                    .then(res => res.json())
                    .then(data => {
                        console.log(data);
                    })
                    .catch(err => console.error(err));
                }
            });

            table.on("dataLoaded", function(data){
                data.forEach(function(row){
                    $.ajax({
                        url: '/amazon/save-nr', 
                        type: 'POST',
                        data: {
                            sku: row.sku,
                            spend_l30: row.SPEND_L30,
                            _token: $('meta[name="csrf-token"]').attr('content')
                        },
                        success: function(res){
                            // console.log("SPEND saved for SKU:", row.sku);
                        },
                        error: function(err){
                            console.error("SPEND save failed for SKU:", row.sku, err);
                        }
                    });
                });
            });

            table.on("tableBuilt", function() {

                function combinedFilter(data) {

                    let searchVal = $("#global-search").val()?.toLowerCase() || "";
                    if (searchVal && !(data.sku?.toLowerCase().includes(searchVal))) {
                        return false;
                    }

                    let statusVal = $("#status-filter").val();
                    if (statusVal && data.campaignStatus !== statusVal) {
                        return false;
                    }

                    let invFilterVal = $("#inv-filter").val();
                    if (invFilterVal === "ALL") {
                        if (parseFloat(data.INV) === 0) return false;
                    } else if (invFilterVal === "INV_0") {
                        if (parseFloat(data.INV) !== 0) return false;
                    } else if (invFilterVal === "OTHERS") {
                        if (parseFloat(data.INV) === 0) return false;
                    }

                    let nrlFilterVal = $("#nrl-filter").val();
                    if (nrlFilterVal) {
                        let rowSelect = document.querySelector(
                            `select[data-sku="${data.sku}"][data-field="NRL"]`
                        );
                        let rowVal = rowSelect ? rowSelect.value : "";
                        if (!rowVal) rowVal = data.NRL || "";

                        if (rowVal !== nrlFilterVal) return false;
                    }

                    let nraFilterVal = $("#nra-filter").val();
                    if (nraFilterVal) {
                        let rowSelect = document.querySelector(
                            `select[data-sku="${data.sku}"][data-field="NR"]`
                        );
                        let rowVal = rowSelect ? rowSelect.value : "";
                        if (!rowVal) rowVal = data.NR || "";

                        if (rowVal !== nraFilterVal) return false;
                    }

                    let fbaFilterVal = $("#fba-filter").val();
                    if (fbaFilterVal) {
                        let rowSelect = document.querySelector(
                            `select[data-sku="${data.sku}"][data-field="FBA"]`
                        );
                        let rowVal = rowSelect ? rowSelect.value : "";
                        if (!rowVal) rowVal = data.FBA || "";

                        if (rowVal !== fbaFilterVal) return false;
                    }

                    return true;
                }

                table.setFilter(combinedFilter);

                function updateCampaignStats() {
                    let allRows = table.getData();
                    let filteredRows = allRows.filter(combinedFilter);

                    let total = allRows.length;
                    let filtered = filteredRows.length;

                    let percentage = total > 0 ? ((filtered / total) * 100).toFixed(0) : 0;

                    document.getElementById("total-campaigns").innerText = filtered;
                    document.getElementById("percentage-campaigns").innerText = percentage + "%";
                }

                table.on("dataFiltered", updateCampaignStats);
                table.on("pageLoaded", updateCampaignStats);
                table.on("dataProcessed", updateCampaignStats);

                $("#global-search").on("keyup", function() {
                    table.setFilter(combinedFilter);
                });

                $("#status-filter,#inv-filter, #nrl-filter, #nra-filter, #fba-filter").on("change",
                    function() {
                        table.setFilter(combinedFilter);
                    });

                updateCampaignStats();
            });

            document.addEventListener("click", function(e) {
                if (e.target.classList.contains("toggle-cols-btn")) {
                    let btn = e.target;

                    let colsToToggle = ["INV", "L30", "DIL %", "A_L30", "A DIL %", "NRL", "NRA", "FBA"];

                    colsToToggle.forEach(colName => {
                        let col = table.getColumn(colName);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
            });

            document.addEventListener("click", function(e) {
                if (e.target.classList.contains("toggle-spendL30-btn")) {
                    let colsToToggle = ["kw_spend_L30", "pt_spend_L30", "hl_spend_L30"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-spendL7-btn")) {
                    let colsToToggle = ["kw_spend_L7", "pt_spend_L7", "hl_spend_L7"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-soldL30-btn")) {
                    let colsToToggle = ["kw_sold_L30", "pt_sold_L30", "hl_sold_L30"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-soldL7-btn")) {
                    let colsToToggle = ["kw_sold_L7", "pt_sold_L7", "hl_sold_L7"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-salesL30-btn")) {
                    let colsToToggle = ["kw_sales_L30", "pt_sales_L30", "hl_sales_L30"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-salesL7-btn")) {
                    let colsToToggle = ["kw_sales_L7", "pt_sales_L7", "hl_sales_L7"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-clicksL30-btn")) {
                    let colsToToggle = ["kw_clicks_L30", "pt_clicks_L30", "hl_clicks_L30"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-clicksL7-btn")) {
                    let colsToToggle = ["kw_clicks_L7", "pt_clicks_L7", "hl_clicks_L7"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-impL30-btn")) {
                    let colsToToggle = ["kw_impr_L30", "pt_impr_L30", "hl_impr_L30"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                if (e.target.classList.contains("toggle-impL7-btn")) {
                    let colsToToggle = ["kw_impr_L7", "pt_impr_L7", "hl_impr_L7"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
                
            });

            document.getElementById("export-btn").addEventListener("click", function () {
                let allData = table.getData("active"); 

                if (allData.length === 0) {
                    alert("No data available to export!");
                    return;
                }

                let exportData = allData.map(row => ({
                    sku: row.sku,
                    SPEND_L30: row.SPEND_L30.toFixed(2),
                }));

                let ws = XLSX.utils.json_to_sheet(exportData);
                let wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Campaigns");

                XLSX.writeFile(wb, "amazon_ad_running.xlsx");
            });

            document.body.style.zoom = "85%";
        });
    </script>
@endsection
