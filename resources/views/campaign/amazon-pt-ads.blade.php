


@extends('layouts.vertical', ['title' => 'Amazon PT ADS', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
    <link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.css" />
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

        .tabulator-row:nth-child(even) {
            background-color: #f8fafc !important;
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

        .tabulator-row:hover {
            background-color: #dbeafe !important;
        }

        .parent-row {
            background-color: #e0eaff !important;
            font-weight: 700;
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

        .green-bg {
            color: #05bd30 !important;
        }

        .pink-bg {
            color: #ff01d0 !important;
        }

        .red-bg {
            color: #ff2727 !important;
        }
        .parent-row-bg{
            background-color: #c3efff !important;
        }
    </style>
@endsection
@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Amazon PT ADS',
        'sub_title' => 'Amazon PT ADS',
    ])
    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm border-0">
                <div class="card-body">
                    <div class="mb-3">
                        <button id="daterange-btn" class="btn btn-outline-dark">
                            <span>Date range: Select</span> <i class="fa-solid fa-chevron-down ms-1"></i>
                        </button>
                    </div>
                    <!-- Stats Row -->
                    <div class="row text-center mb-4">
                        <!-- Clicks -->
                        <div class="col-md-3 mb-3 mb-md-0">
                            <div class="p-3 border rounded bg-light h-100">
                                <div class="text-muted small">Clicks</div>
                                <div class="h3 mb-0 fw-bold text-primary card-clicks">{{ $clicks->sum() }}</div>
                            </div>
                        </div>

                        <!-- Spend -->
                        <div class="col-md-3 mb-3 mb-md-0">
                            <div class="p-3 border rounded bg-light h-100">
                                <div class="text-muted small">Spend</div>
                                <div class="h3 mb-0 fw-bold text-success card-spend">
                                    US${{ number_format($spend->sum(), 2) }}
                                </div>
                            </div>
                        </div>

                        <!-- Orders -->
                        <div class="col-md-3 mb-3 mb-md-0">
                            <div class="p-3 border rounded bg-light h-100">
                                <div class="text-muted small">Orders</div>
                                <div class="h3 mb-0 fw-bold text-danger card-orders">{{ $orders->sum() }}</div>
                            </div>
                        </div>

                        <!-- Sales -->
                        <div class="col-md-3">
                            <div class="p-3 border rounded bg-light h-100">
                                <div class="d-flex justify-content-between align-items-center">
                                    <div>
                                        <div class="text-muted small">Sales</div>
                                        <div class="h3 mb-0 fw-bold text-info card-sales">
                                            US${{ number_format($sales->sum(), 2) }}
                                        </div>
                                    </div>
                                    <!-- Arrow button -->
                                    <button id="toggleChartBtn" class="btn btn-sm btn-info ms-2">
                                        <i id="chartArrow" class="fa-solid fa-chevron-down"></i>
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Chart (hidden by default) -->
                    <div id="chartContainer" style="display: none;">
                        <canvas id="campaignChart" height="120"></canvas>
                    </div>

                </div>
            </div>
        </div>

        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body py-3">
                    <div class="mb-4">
                        <!-- Title -->
                        <h4 class="fw-bold text-primary mb-3 d-flex align-items-center">
                            <i class="fa-solid fa-chart-line me-2"></i>
                            Amazon PT ADS
                        </h4>

                        <!-- Search and Controls Row -->
                        <div class="row g-3 mb-3">
                            <!-- Inventory Filters -->
                            <div class="col-md-4">
                                <div class="d-flex gap-2">
                                    <select id="inv-filter" class="form-select form-select-md">
                                        <option value="">Select INV</option>
                                        <option value="ALL">ALL</option>
                                        <option value="INV_0">0 INV</option>
                                        <option value="OTHERS">OTHERS</option>
                                    </select>

                                    <select id="acos-filter" class="form-select form-select-md">
                                        <option value="">Select ACOS</option>
                                        <option value="PINK">PINK</option>
                                        <option value="GREEN">GREEN</option>
                                        <option value="RED">RED</option>
                                    </select>

                                    <select id="cvr-filter" class="form-select form-select-md">
                                        <option value="">Select CVR</option>
                                        <option value="PINK">PINK</option>
                                        <option value="GREEN">GREEN</option>
                                        <option value="RED">RED</option>
                                    </select>
                                </div>
                            </div>
                            <div class="col-md-8 d-flex justify-content-end gap-2">
                                <a href="javascript:void(0)" id="export-btn" class="btn btn-sm btn-success d-flex align-items-center justify-content-center">
                                    <i class="fas fa-file-export me-1"></i> Export Excel/CSV
                                </a>
                                {{-- <button class="btn btn-info btn-md d-flex align-items-center"> --}}
                                    <span>Total Spend: <span id="total-spend" class="fw-bold ms-1 fs-4">0</span></span>
                                {{-- </button> --}}
                                {{-- <button class="btn btn-info btn-md d-flex align-items-center"> --}}
                                    <span>Total Sales: <span id="total-sales" class="fw-bold ms-1 fs-4">0</span></span>
                                {{-- </button> --}}
                                <button class="btn btn-success btn-md d-flex align-items-center">
                                    <span>Total Campaigns: <span id="total-campaigns" class="fw-bold ms-1 fs-4">0</span></span>
                                </button>
                                <button class="btn btn-primary btn-md d-flex align-items-center">
                                    <i class="fa fa-percent me-1"></i>
                                    <span>Of Total: <span id="percentage-campaigns" class="fw-bold ms-1 fs-4">0%</span></span>
                                </button>
                            </div>
                        </div>

                        <!-- Search and Controls Row -->
                        <div class="row g-3 align-items-center">
                            <!-- Left: Search & Status -->
                            <div class="col-md-6">
                                <div class="d-flex gap-2">
                                    <input type="text" id="global-search" class="form-control form-control-md" placeholder="Search campaign...">
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
    <!-- Moment.js (daterangepicker dependency) -->
    <script src="https://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
    <!-- Daterangepicker -->
    <script type="text/javascript" src="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
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
                index: "Sku",
                ajaxURL: "/amazon/pt/ads/data",
                layout: "fitData",
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
                        title: "CAMPAIGN",
                        field: "campaignName"
                    },
                    {
                        title: "7 UB%",
                        field: "spend_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var spend_l7 = parseFloat(row.spend_l7) || 0;
                            var budget = parseFloat(row.campaignBudgetAmount) || 0;
                            var ub7 = budget > 0 ? (spend_l7 / (budget * 7)) * 100 : 0;

                            var td = cell.getElement();
                            td.classList.remove('green-bg', 'pink-bg', 'red-bg');
                            if (ub7 >= 70 && ub7 <= 90) {
                                td.classList.add('green-bg');
                            } else if (ub7 > 90) {
                                td.classList.add('pink-bg');
                            } else if (ub7 < 70) {
                                td.classList.add('red-bg');
                            }

                            return ub7.toFixed(0) + "%";
                        }
                    },
                    {
                        title: "IMP L30",
                        field: "impressions_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let impressions_l30 = cell.getValue();
                            return `
                                <span>${parseFloat(cell.getValue() || 0).toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary impressions_l30_btn" 
                                    data-impression-l30="${impressions_l30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "IMP L60",
                        field: "impressions_l60",
                        hozAlign: "right",
                        formatter: function(cell) {
                            return `
                                <span>${parseFloat(cell.getValue() || 0).toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "IMP L15",
                        field: "impressions_l15",
                        hozAlign: "right",
                        formatter: function(cell) {
                            return `
                                <span>${parseFloat(cell.getValue() || 0).toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "IMP L7",
                        field: "impressions_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let impressions_l7 = cell.getValue();
                            return `
                                <span>${impressions_l7}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Clicks L30",
                        field: "clicks_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let color = value < 50 ? "red" : "green";
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}
                                </span>
                                <i class="fa fa-info-circle text-primary clicks_l30_btn" 
                                data-clicks-l30="${value}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "Clicks L60",
                        field: "clicks_l60",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let color = value < 50 ? "red" : "green";
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Clicks L15",
                        field: "clicks_l15",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let color = value < 50 ? "red" : "green";
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Clicks L7",
                        field: "clicks_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let color = value < 50 ? "red" : "green";
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Spend L30",
                        field: "spend_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary spend_l30_btn" 
                                data-spend-l30="${value}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "Spend L60",
                        field: "spend_l60",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Spend L15",
                        field: "spend_l15",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Spend L7",
                        field: "spend_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Ad Sales L30",
                        field: "ad_sales_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary ad_sales_l30_btn" 
                                    data-ad_sales-l30="${value}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "Ad Sales L60",
                        field: "ad_sales_l60",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Ad Sales L15",
                        field: "ad_sales_l15",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Ad Sales L7",
                        field: "ad_sales_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Ad Sold L30",
                        field: "ad_sold_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                                <i class="fa fa-info-circle text-primary ad_sold_l30_btn" 
                                    data-ad_sold-l30="${value}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "Ad Sold L60",
                        field: "ad_sold_l60",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Ad Sold L15",
                        field: "ad_sold_l15",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "Ad Sold L7",
                        field: "ad_sold_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            return `
                                <span>${value.toFixed(0)}</span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "ACOS L30",
                        field: "acos_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let row = cell.getRow().getData();
                            let adSales = parseFloat(row.ad_sales_l30 || 0);

                            if (adSales === 0) {
                                value = 100;
                            }

                            let color = "green";
                            if(value == 100){
                                color = "#000000";
                            }else if (value < 7) {
                                color = "#e83e8c";
                            } else if (value >= 7 && value <= 14) {
                                color = "green";
                            } else if (value > 14) {
                                color = "red";
                            }

                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}%
                                </span>
                                <i class="fa fa-info-circle text-primary acos_l30_btn" 
                                    data-acos-l30="${value}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "ACOS L60",
                        field: "acos_l60",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let row = cell.getRow().getData();
                            let adSales = parseFloat(row.ad_sales_l60 || 0);

                            if (adSales === 0) {
                                value = 100;
                            }

                            let color = "green";
                            if(value == 100){
                                color = "#000000";
                            }else if (value < 7) {
                                color = "#e83e8c";
                            }else if (value >= 7 && value <= 14) {
                                color = "green";
                            } else if (value > 14) {
                                color = "red";
                            }

                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}%
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "ACOS L15",
                        field: "acos_l15",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let row = cell.getRow().getData();
                            let adSales = parseFloat(row.ad_sales_l15 || 0);

                            if (adSales === 0) {
                                value = 100;
                            }

                            let color = "green";
                            if(value == 100){
                                color = "#000000";
                            }else if (value < 7) {
                                color = "#e83e8c";
                            }else if (value >= 7 && value <= 14) {
                                color = "green";
                            } else if (value > 14) {
                                color = "red";
                            }

                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}%
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "ACOS L7",
                        field: "acos_l7",
                        hozAlign: "right",
                        formatter: function(cell) {
                            let value = parseFloat(cell.getValue() || 0);
                            let row = cell.getRow().getData();
                            let adSales = parseFloat(row.ad_sales_l7 || 0);

                            if (adSales === 0) {
                                value = 100;
                            }

                            let color = "green";
                            if(value == 100){
                                color = "#000000";
                            }else if (value < 7) {
                                color = "#e83e8c";
                            }else if (value >= 7 && value <= 14) {
                                color = "green";
                            } else if (value > 14) {
                                color = "red";
                            }

                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${value.toFixed(0)}%
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "CPC L30",
                        field: "cpc_l30",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l30 = parseFloat(row.cpc_l30) || 0;

                            return `
                                <span>
                                    ${cpc_l30.toFixed(2)}
                                </span>
                                <i class="fa fa-info-circle text-primary cpc_l30_btn" 
                                    data-cpc-l30="${cpc_l30}" 
                                style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "CPC L60",
                        field: "cpc_l60",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l60 = parseFloat(row.cpc_l60) || 0;
                            return cpc_l60.toFixed(2);
                        },
                        visible: false
                    },
                    {
                        title: "CPC L15",
                        field: "cpc_l15",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l15 = parseFloat(row.cpc_l15) || 0;
                            return cpc_l15.toFixed(2);
                        },
                        visible: false
                    },
                    {
                        title: "CPC L7",
                        field: "cpc_l7",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l7 = parseFloat(row.cpc_l7) || 0;
                            return `
                                <span>
                                    ${cpc_l7.toFixed(2)}
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "CVR L30",
                        field: "cvr_l30",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var ad_sold_l30 = parseFloat(row.ad_sold_l30) || 0;
                            var clicks_l30 = parseFloat(row.clicks_l30) || 0;
                            
                            var cvr_l30 = (clicks_l30 > 0) ? (ad_sold_l30 / clicks_l30) * 100 : 0;
                            let color = "";
                            if (cvr_l30 < 5) {
                                color = "red";
                            } else if (cvr_l30 >= 5 && cvr_l30 <= 10) {
                                color = "green";
                            } else if (cvr_l30 > 10){
                                color = "#e83e8c";
                            }
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${cvr_l30.toFixed(0)}%
                                </span>
                                <i class="fa fa-info-circle text-primary cvr_l30_btn" 
                                    data-cvr-l30="${cvr_l30}" 
                                    style="cursor:pointer; margin-left:8px;"></i>
                            `;
                        }
                    },
                    {
                        title: "CVR L60",
                        field: "cvr_l60",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var ad_sold_l60 = parseFloat(row.ad_sold_l60) || 0;
                            var clicks_l60 = parseFloat(row.clicks_l60) || 0;
                            
                            var cvr_l60 = (clicks_l60 > 0) ? (ad_sold_l60 / clicks_l60) * 100 : 0;
                            let color = "";
                            if (cvr_l60 < 5) {
                                color = "red";
                            } else if (cvr_l60 >= 5 && cvr_l60 <= 10) {
                                color = "green";
                            } else if (cvr_l60 > 10){
                                color = "#e83e8c";
                            }
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${cvr_l60.toFixed(0)}%
                                </span>
                            `;

                        },
                        visible: false
                    },
                    {
                        title: "CVR L15",
                        field: "cvr_l15",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var ad_sold_l15 = parseFloat(row.ad_sold_l15) || 0;
                            var clicks_l15 = parseFloat(row.clicks_l15) || 0;

                            var cvr_l15 = (clicks_l15 > 0) ? (ad_sold_l15 / clicks_l15) * 100 : 0;
                            let color = "";
                            if (cvr_l15 < 5) {
                                color = "red";
                            } else if (cvr_l15 >= 5 && cvr_l15 <= 10) {
                                color = "green";
                            } else if (cvr_l15 > 10){
                                color = "#e83e8c";
                            }
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${cvr_l15.toFixed(0)}%
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "CVR L7",
                        field: "cvr_l7",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var ad_sold_l7 = parseFloat(row.ad_sold_l7) || 0;
                            var clicks_l7 = parseFloat(row.clicks_l7) || 0;

                            var cvr_l7 = (clicks_l7 > 0) ? (ad_sold_l7 / clicks_l7) * 100 : 0;
                            let color = "";
                            if (cvr_l7 < 5) {
                                color = "red";
                            } else if (cvr_l7 >= 5 && cvr_l7 <= 10) {
                                color = "green";
                            } else if (cvr_l7 > 10){
                                color = "#e83e8c";
                            }
                            return `
                                <span style="color:${color}; font-weight:600;">
                                    ${cvr_l7.toFixed(0)}%
                                </span>
                            `;
                        },
                        visible: false
                    },
                    {
                        title: "TPFT%",
                        field: "TPFT",
                        hozAlign: "center",
                        formatter: function(cell){
                            let value = parseFloat(cell.getValue()) || 0;
                            let percent = value.toFixed(0);
                            let color = "";

                            if (value < 10) {
                                color = "red";
                            } else if (value >= 10 && value < 15) {
                                color = "#ffc107";
                            } else if (value >= 15 && value < 20) {
                                color = "blue";
                            } else if (value >= 20 && value <= 40) {
                                color = "green";
                            } else if (value > 40) {
                                color = "#e83e8c";
                            }

                            return `
                                <span style="font-weight:600; color:${color};">
                                    ${percent}%
                                </span>
                            `;
                        }
                    }

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

            table.on("tableBuilt", function() {

                function combinedFilter(data) {

                    let searchVal = $("#global-search").val()?.toLowerCase() || "";

                    if (searchVal) {
                        let fieldsToSearch = ["sku", "campaignName", "campaignStatus"];
                        let match = fieldsToSearch.some(field => {
                            return data[field]?.toString().toLowerCase().includes(searchVal);
                        });
                        if (!match) return false;
                    }

                    let statusVal = $("#status-filter").val();
                    if (statusVal && data.campaignStatus !== statusVal) {
                        return false;
                    }

                    let invFilterVal = $("#inv-filter").val();
                    if (!invFilterVal) {
                        // if (parseFloat(data.INV) != 0) return false;
                    } else if (invFilterVal === "INV_0") {
                        if (parseFloat(data.INV) !== 0) return false;
                    } else if (invFilterVal === "OTHERS") {
                        if (parseFloat(data.INV) === 0) return false;
                    }

                    let acosFilterVal = $("#acos-filter").val();
                    if (acosFilterVal) {
                        let acosFields = ["acos_l30"];

                        let matched = acosFields.every(field => {
                            let val = parseFloat(data[field]) || 0;

                            if (acosFilterVal === "PINK") {
                                return val < 7;
                            }
                            if (acosFilterVal === "GREEN") {
                                return val >= 7 && val <= 14;
                            }
                            if (acosFilterVal === "RED") {
                                return val > 14;
                            }
                            return false;
                        });

                        if (!matched) return false;
                    }

                    let cvrFilterVal = $("#cvr-filter").val();
                    if (cvrFilterVal) {
                        let cvrFields = ["cvr_l30"];

                        let matched = cvrFields.every(field => {
                            let val = parseFloat(data[field]) || 0;

                            if (cvrFilterVal === "PINK") {
                                return val > 10;
                            }
                            if (cvrFilterVal === "GREEN") {
                                return val >= 5 && val <= 10;
                            }
                            if (cvrFilterVal === "RED") {
                                return val < 5;
                            }
                            return false;
                        });

                        if (!matched) return false;
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

                    let totalSales = filteredRows.reduce((sum, row) => sum + (parseFloat(row.ad_sales_l30) || 0), 0);
                    let totalSpend = filteredRows.reduce((sum, row) => sum + (parseFloat(row.spend_l30) || 0), 0);
                    // let totalOrders = filteredRows.reduce((sum, row) => sum + (parseFloat(row.ad_sold_l30) || 0), 0);

                    const totalSalesEl = document.getElementById("total-sales");
                    const totalSpendEl = document.getElementById("total-spend");
                    const totalEl = document.getElementById("total-campaigns");
                    const percentageEl = document.getElementById("percentage-campaigns");

                    if (totalSalesEl) totalSalesEl.innerText = totalSales.toFixed(0);
                    if (totalSpendEl) totalSpendEl.innerText = totalSpend.toFixed(0);
                    if (totalEl) totalEl.innerText = filtered;
                    if (percentageEl) percentageEl.innerText = percentage + "%";
                }

                table.on("dataFiltered", updateCampaignStats);
                table.on("pageLoaded", updateCampaignStats);
                table.on("dataProcessed", updateCampaignStats);

                $("#global-search").on("keyup", function() {
                    table.setFilter(combinedFilter);
                });

                $("#status-filter,#inv-filter,#acos-filter,#cvr-filter").on("change", function() {
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

                if (e.target.classList.contains("impressions_l30_btn")) {
                    let colsToToggle = ["impressions_l60", "impressions_l15", "impressions_l7"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("clicks_l30_btn")) {
                    let colsToToggle = ["clicks_l15", "clicks_l7", "clicks_l60"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("spend_l30_btn")) {
                    let colsToToggle = ["spend_l15", "spend_l7", "spend_l60"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("ad_sales_l30_btn")) {
                    let colsToToggle = ["ad_sales_l15", "ad_sales_l7", "ad_sales_l60"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("ad_sold_l30_btn")) {
                    let colsToToggle = ["ad_sold_l15", "ad_sold_l7", "ad_sold_l60"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("acos_l30_btn")) {
                    let colsToToggle = ["acos_l15", "acos_l7", "acos_l60"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("cpc_l30_btn")) {
                    let colsToToggle = ["cpc_l15", "cpc_l7", "cpc_l60"];

                    colsToToggle.forEach(colField => {
                        let col = table.getColumn(colField);
                        if (col) {
                            col.toggle();
                        }
                    });
                }

                if (e.target.classList.contains("cvr_l30_btn")) {
                    let colsToToggle = ["cvr_l15", "cvr_l7", "cvr_l60"];

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

                let exportData = allData.map(row => ({ ...row }));

                let ws = XLSX.utils.json_to_sheet(exportData);
                let wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Campaigns");

                XLSX.writeFile(wb, "amazon_pt_ads_report.xlsx");
            });

            document.body.style.zoom = "78%";
        });
    </script>
    <script>
        const ctx = document.getElementById('campaignChart').getContext('2d');

        const chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: {!! json_encode($dates) !!},
                datasets: [
                    {
                        label: 'Clicks',
                        data: {!! json_encode($clicks) !!},
                        borderColor: 'purple',
                        backgroundColor: 'rgba(128, 0, 128, 0.1)',
                        yAxisID: 'y1',
                        tension: 0.4,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        fill: false,
                    },
                    {
                        label: 'Spend (USD)',
                        data: {!! json_encode($spend) !!},
                        borderColor: 'teal',
                        backgroundColor: 'rgba(0, 128, 128, 0.1)',
                        yAxisID: 'y2',
                        tension: 0.4,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        fill: false,
                    },
                    {
                        label: 'Orders',
                        data: {!! json_encode($orders) !!},
                        borderColor: 'magenta',
                        backgroundColor: 'rgba(255, 0, 255, 0.1)',
                        yAxisID: 'y1',
                        tension: 0.4,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        fill: false,
                    },
                    {
                        label: 'Sales (USD)',
                        data: {!! json_encode($sales) !!},
                        borderColor: 'blue',
                        backgroundColor: 'rgba(0, 0, 255, 0.1)',
                        yAxisID: 'y2',
                        tension: 0.4,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        fill: false,
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    tooltip: {
                        backgroundColor: "#fff",
                        titleColor: "#111",
                        bodyColor: "#333",
                        borderColor: "#ddd",
                        borderWidth: 1,
                        padding: 12,
                        titleFont: { size: 14, weight: 'bold' },
                        bodyFont: { size: 13 },
                        usePointStyle: true,
                        callbacks: {
                            label: function(context) {
                                let value = context.raw;
                                if (context.dataset.label.includes("Spend") || context.dataset.label.includes("Sales")) {
                                    return `${context.dataset.label}: $${Number(value).toFixed(2)}`;
                                }
                                return `${context.dataset.label}: ${value}`;
                            }
                        }
                    },
                    legend: {
                        labels: {
                            usePointStyle: true,
                            boxWidth: 10,
                            padding: 20
                        },
                        onClick: (e, legendItem, legend) => {
                            const index = legendItem.datasetIndex;
                            const ci = legend.chart;
                            const meta = ci.getDatasetMeta(index);
                            meta.hidden = meta.hidden === null ? !ci.data.datasets[index].hidden : null;
                            ci.update();
                        }
                    }
                },
                scales: {
                    y1: {
                        type: 'linear',
                        position: 'left',
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Clicks / Orders'
                        }
                    },
                    y2: {
                        type: 'linear',
                        position: 'right',
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Spend / Sales (USD)'
                        },
                        grid: {
                            drawOnChartArea: false
                        }
                    }
                }
            }
        });

        document.addEventListener("DOMContentLoaded", function() {
            const toggleBtn = document.getElementById("toggleChartBtn");
            const chartContainer = document.getElementById("chartContainer");
            const arrowIcon = document.getElementById("chartArrow");

            toggleBtn.addEventListener("click", function() {
                if (chartContainer.style.display === "none") {
                    chartContainer.style.display = "block";
                    arrowIcon.classList.remove("fa-chevron-down");
                    arrowIcon.classList.add("fa-chevron-up");
                } else {
                    chartContainer.style.display = "none";
                    arrowIcon.classList.remove("fa-chevron-up");
                    arrowIcon.classList.add("fa-chevron-down");
                }
            });
        });

        $(function() {
            let picker = $('#daterange-btn').daterangepicker({
                opens: 'right',
                autoUpdateInput: false,
                alwaysShowCalendars: true,
                locale: {
                    format: "D MMM YYYY",
                    cancelLabel: 'Clear'
                },
                ranges: {
                    'Today': [moment(), moment()],
                    'Yesterday': [moment().subtract(1, 'days'), moment().subtract(1, 'days')],
                    'Last 7 Days': [moment().subtract(6, 'days'), moment()],
                    'Last 30 Days': [moment().subtract(29, 'days'), moment()],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Last Month': [moment().subtract(1, 'month').startOf('month'), moment().subtract(1, 'month').endOf('month')]
                }
            }, function(start, end) {
                const startDate = start.format("YYYY-MM-DD");
                const endDate   = end.format("YYYY-MM-DD");

                $('#daterange-btn span').html("Date range: " + startDate + " - " + endDate);
                fetchChartData(startDate, endDate);
            });

            // Reset on cancel
            $('#daterange-btn').on('cancel.daterangepicker', function(ev, picker) {
                $(this).find('span').html("Date range: Select");
                fetchChartData(); 
            });

        });

        function fetchChartData(startDate, endDate) {
            $.ajax({
                url: "{{ route('amazonPtAds.filter') }}",
                type: "GET",
                data: { startDate, endDate },
                success: function(response) {
                    const formattedDates = response.dates.map(d => moment(d).format('MMM DD'));
                    chart.data.labels = formattedDates;
                    chart.data.datasets[0].data = response.clicks;
                    chart.data.datasets[1].data = response.spend;
                    chart.data.datasets[2].data = response.orders;
                    chart.data.datasets[3].data = response.sales;
                    chart.update();

                    $('.card-clicks').text(response.totals.clicks);
                    $('.card-spend').text('$' + Number(response.totals.spend).toFixed(2));
                    $('.card-orders').text(response.totals.orders);
                    $('.card-sales').text('$' + Number(response.totals.sales).toFixed(2));
                }
            });
        }
    </script>
@endsection

