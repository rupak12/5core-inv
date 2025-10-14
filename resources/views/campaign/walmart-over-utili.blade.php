@extends('layouts.vertical', ['title' => 'Walmart - Over Utilised', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
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
    </style>
@endsection
@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Walmart - Over Utilised',
        'sub_title' => 'Walmart - Over Utilised',
    ])
    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body py-3">
                    <div class="mb-4">
                        <!-- Title -->
                        <h4 class="fw-bold text-primary mb-3 d-flex align-items-center">
                            <i class="fa-solid fa-chart-line me-2"></i>
                            Walmart - Over Utilised
                        </h4>

                        <!-- Filters Row -->
                        <div class="row g-3 mb-3">
                            <!-- Inventory Filters -->
                            <div class="col-md-6">
                                <div class="d-flex gap-2">
                                    <select id="inv-filter" class="form-select form-select-md">
                                        <option value="ALL">ALL</option>
                                        <option value="INV_0">0 INV</option>
                                        <option value="OTHERS">OTHERS</option>
                                    </select>

                                    <select id="nrl-filter" class="form-select form-select-md">
                                        <option value="">Select NRL</option>
                                        <option value="NRL">NRL</option>
                                        <option value="RL">RL</option>
                                    </select>

                                </div>
                            </div>

                            <!-- Stats -->
                            <div class="col-md-6">
                                <div class="d-flex gap-2 justify-content-end">
                                    <button id="apr-all-sbid-btn" class="btn btn-info btn-sm d-none">
                                        APR ALL SBID
                                    </button> 
                                    <a href="javascript:void(0)" id="export-btn" class="btn btn-sm btn-success d-flex align-items-center justify-content-center">
                                        <i class="fas fa-file-export me-1"></i> Export Excel/CSV
                                    </a>
                                    <button class="btn btn-success btn-md">
                                        <i class="fa fa-arrow-down me-1"></i>
                                        Need to decrease bids: <span id="total-campaigns" class="fw-bold ms-1 fs-4">0</span>
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

    <div id="progress-overlay" style="display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.7); z-index: 9999;">
        <div style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); text-align: center;">
            <div class="spinner-border text-light" role="status" style="width: 3rem; height: 3rem;">
                <span class="visually-hidden">Loading...</span>
            </div>
            <div class="mt-3" style="color: white; font-size: 1.2rem; font-weight: 500;">
                Updating campaigns...
            </div>
            <div style="color: #a3e635; font-size: 0.9rem; margin-top: 0.5rem;">
                Please wait while we process your request
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
            document.body.style.zoom = "78%";

            const invFilter  = document.querySelector("#inv-filter");
            const nrlFilter  = document.querySelector("#nrl-filter");
            const nraFilter  = document.querySelector("#nra-filter");
            const fbaFilter  = document.querySelector("#fba-filter");


            const getDilColor = (value) => {
                const percent = parseFloat(value) * 100;
                if (percent < 16.66) return 'red';
                if (percent >= 16.66 && percent < 25) return 'yellow';
                if (percent >= 25 && percent < 50) return 'green';
                return 'pink';
            };

            var table = new Tabulator("#budget-under-table", {
                index: "Sku",
                ajaxURL: "/walmart/utilized/kw/data",
                layout: "fitData",
                movableColumns: true,
                resizableColumns: true,
                rowFormatter: function(row) {
                    const data = row.getData();
                    const sku = data["Sku"] || '';

                    if (sku.toUpperCase().includes("PARENT")) {
                        row.getElement().classList.add("parent-row");
                    }
                },
                columns: [
                    {
                        formatter: "rowSelection",
                        titleFormatter: "rowSelection",
                        hozAlign: "center",
                        headerSort: false,
                        width: 50
                    },
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
                        title: "WL 30",
                        field: "WA_L30",
                        visible: false
                    },
                    {
                        title: "NRA",
                        field: "NRA",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue();

                            let bgColor = "";
                            if (value === "NRA") {
                                bgColor = "background-color:#dc3545;color:#fff;"; // red
                            } else if (value === "RA") {
                                bgColor = "background-color:#28a745;color:#fff;"; // green
                            }

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="NR"
                                        style="width: 90px; ${bgColor}">
                                    <option value="RA" ${value === 'RA' ? 'selected' : ''}>RA</option>
                                    <option value="NRA" ${value === 'NRA' ? 'selected' : ''}>NRA</option>
                                </select>
                            `;
                        },
                        visible: false,
                        hozAlign: "center"
                    },
                    {
                        title: "CAMPAIGN",
                        field: "campaignName"
                    },
                    {
                        title: "BGT",
                        field: "campaignBudgetAmount",
                        hozAlign: "right",
                        formatter: (cell) => parseFloat(cell.getValue() || 0)
                    },
                    {
                        title: "ACOS L30",
                        field: "acos_L30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            return `
                                <span>${parseFloat(cell.getValue() || 0).toFixed(0) + "%"}</span>
                            `;
                            
                        },
                        visible: false,
                    },
                    {
                        title: "Clicks L30",
                        field: "clicks_l30",
                        hozAlign: "right",
                        formatter: function(cell) {
                            return `
                                <span>${parseFloat(cell.getValue() || 0).toFixed(0)}</span>
                            `;
                        }
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
                        title: "1 UB%",
                        field: "spend_l1",
                        hozAlign: "right",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var spend_l1 = parseFloat(row.spend_l1) || 0;
                            var budget = parseFloat(row.campaignBudgetAmount) || 0;
                            var ub1 = budget > 0 ? (spend_l1 / budget) * 100 : 0;

                            var td = cell.getElement();
                            td.classList.remove('green-bg', 'pink-bg', 'red-bg');
                            if (ub1 >= 70 && ub1 <= 90) {
                                td.classList.add('green-bg');
                            } else if (ub1 > 90) {
                                td.classList.add('pink-bg');
                            } else if (ub1 < 70) {
                                td.classList.add('red-bg');
                            }

                            return ub1.toFixed(0) + "%";
                        }
                    },
                    {
                        title: "L7 CPC",
                        field: "cpc_l7",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l7 = parseFloat(row.cpc_l7) || 0;
                            return cpc_l7.toFixed(2);
                        }
                    },
                    {
                        title: "L1 CPC",
                        field: "cpc_l1",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l1 = parseFloat(row.cpc_l1) || 0;
                            return cpc_l1.toFixed(2);
                        }
                    },
                    {
                        title: "SBID",
                        field: "sbid",
                        hozAlign: "center",
                        formatter: function(cell) {
                            var row = cell.getRow().getData();
                            var cpc_l1 = parseFloat(row.cpc_l1) || 0;
                            var cpc_l7 = parseFloat(row.cpc_l7) || 0;
                            var sbid;
                            if(cpc_l1 > cpc_l7) {
                                sbid = (cpc_l1 * 0.95).toFixed(2);
                            }else{
                                sbid = (cpc_l7 * 0.95).toFixed(2);
                            }
                            return sbid;
                        },
                    },
                    {
                        title: "APR BID",
                        field: "apr_bid",
                        hozAlign: "center",
                        formatter: function(cell, formatterParams, onRendered) {
                            var value = cell.getValue() || 0;
                            return `
                                <div style="align-items:center; gap:5px;">
                                    <button class="btn btn-primary update-row-btn">APR BID</button>
                                </div>
                            `;
                        },
                        cellClick: function(e, cell) {
                            if (e.target.classList.contains("update-row-btn")) {
                                var row = cell.getRow().getData();
                                var cpc_l1 = parseFloat(row.cpc_l1) || 0;
                                var cpc_l7 = parseFloat(row.cpc_l7) || 0;
                                var sbid;
                                if(cpc_l1 > cpc_l7) {
                                    sbid = (cpc_l1 * 0.95).toFixed(2);
                                }else{
                                    sbid = (cpc_l7 * 0.95).toFixed(2);
                                }
                                updateBid(sbid, rowData.campaign_id);
                            }
                        }
                    },
                    {
                        title: "Status",
                        field: "campaignStatus",
                        hozAlign: "center",
                        formatter: function(cell) {
                            const row = cell.getRow();
                            const sku = row.getData().sku;
                            const value = cell.getValue();
                            return `
                                <select class="form-select form-select-sm editable-select" 
                                        data-sku="${sku}" 
                                        data-field="status"
                                        style="width: 110px;">
                                    <option value="PAUSED" ${value === 'PAUSED' ? 'selected' : ''}>PAUSED</option>
                                    <option value="LIVE" ${value === 'LIVE' ? 'selected' : ''}>LIVE</option>
                                </select>
                            `;
                        }
                    }
                ],
                ajaxResponse: function(url, params, response) {
                    return response.data;
                }
            });

            table.on("rowSelectionChanged", function(data, rows){
                if(data.length > 0){
                    document.getElementById("apr-all-sbid-btn").classList.remove("d-none");
                } else {
                    document.getElementById("apr-all-sbid-btn").classList.add("d-none");
                }
            });

            $(document).on("change", ".editable-select", function () {
                let select = this;
                let sku = select.getAttribute("data-sku");
                let field = select.getAttribute("data-field");
                let value = select.value;

                console.log(`SKU: ${sku}, Field: ${field}, Value: ${value}`);

                fetch('/walmart/save-nr', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRF-TOKEN': '{{ csrf_token() }}'
                    },
                    body: JSON.stringify({ sku, nr: value })
                })
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        let bgColor = "";
                        if (value === "NRA") {
                            bgColor = "background-color:#dc3545;color:#fff;"; 
                        } else if (value === "RA") {
                            bgColor = "background-color:#28a745;color:#fff;";
                        } else if (value === "LATER") {
                            bgColor = "background-color:#ffc107;color:#000;";
                        }
                        select.style = `width: 100px; ${bgColor}`;
                    } else {
                        console.error('Failed to update status');
                    }
                })
                .catch(error => {
                    console.error('Error:', error);
                });
            });

            table.on("tableBuilt", function () {

                function combinedFilter(data) {
                    var budget = parseFloat(data.campaignBudgetAmount) || 0;
                    var spend_l7 = parseFloat(data.spend_l7 || 0);
                    var spend_l1 = parseFloat(data.spend_l1 || 0);

                    var ub7 = budget > 0 ? (spend_l7 / (budget * 7)) * 100 : 0;
                    var ub1 = budget > 0 ? (spend_l1 / budget) * 100 : 0;

                    if (!(ub7 > 90)) return false;

                    let searchVal = $("#global-search").val()?.toLowerCase() || "";
                    if (searchVal) {
                        let campaignMatch = data.campaignName?.toLowerCase().includes(searchVal);
                        let skuMatch = data.sku?.toLowerCase().includes(searchVal);

                        if (!(campaignMatch || skuMatch)) {
                            return false;
                        }
                    }

                    let statusVal = $("#status-filter").val();
                    if (statusVal && data.campaignStatus !== statusVal) {
                        return false;
                    }

                    let invFilterVal = $("#inv-filter").val();
                    let invVal = parseFloat(data.INV || 0);

                    if (invFilterVal === "INV_0") {
                        if (invVal !== 0) return false;
                    } else if (invFilterVal === "OTHERS") {
                        if (invVal === 0) return false;
                    }

                    let nrlFilterVal = $("#nrl-filter").val();
                    if (nrlFilterVal) {
                        let rowSelect = getRowSelectBySkuAndField(data.sku, "NRL");
                        let rowVal = rowSelect ? rowSelect.value : "";
                        if (!rowVal) rowVal = data.NRL || "";

                        if (rowVal !== nrlFilterVal) return false;
                    }

                    return true;
                }

                function updateCampaignStats() {
                    let total = table.getDataCount();                 
                    let filtered = table.getDataCount("active");      
                    let percentage = total > 0 ? ((filtered / total) * 100).toFixed(0) : 0;

                    document.getElementById("total-campaigns").innerText = filtered;
                    document.getElementById("percentage-campaigns").innerText = percentage + "%";
                }

                function refreshFilters() {
                    table.setFilter(combinedFilter);
                    updateCampaignStats(); 
                }

                table.setFilter(combinedFilter);

                table.on("dataFiltered", updateCampaignStats);
                table.on("pageLoaded", updateCampaignStats);
                table.on("dataProcessed", updateCampaignStats);

                $("#global-search").on("keyup", refreshFilters);
                $("#status-filter, #nrl-filter, #inv-filter").on("change", refreshFilters);

                updateCampaignStats();
            });

            document.addEventListener("click", function(e) {
                if (e.target.classList.contains("toggle-cols-btn")) {
                    let btn = e.target;

                    let colsToToggle = ["INV", "L30", "DIL %", "WA_L30", "NRL"];

                    colsToToggle.forEach(colName => {
                        let col = table.getColumn(colName);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
            });

            function getRowSelectBySkuAndField(sku, field) {
                try {
                    let escapedSku = CSS.escape(sku); // escape special chars
                    return document.querySelector(`select[data-sku="${escapedSku}"][data-field="${field}"]`);
                } catch (e) {
                    console.warn("Invalid selector for SKU:", sku, e);
                    return null;
                }
            }

            document.getElementById("export-btn").addEventListener("click", function () {
                let filteredData = table.getData("active");

                let exportData = filteredData.map(row => {
                    let cpc_l1 = parseFloat(row.cpc_l1 || 0);
                    let cpc_l7 = parseFloat(row.cpc_l7 || 0);
                    let sbid = 0;

                    if(cpc_l1 > cpc_l7) {
                        sbid = (cpc_l1 * 0.95).toFixed(2);
                    }else{
                        sbid = (cpc_l7 * 0.95).toFixed(2);
                    }

                    return {
                        campaignName: row.campaignName || "",
                        sbid: sbid
                    };
                });

                if (exportData.length === 0) {
                    alert("No data available to export!");
                    return;
                }

                let ws = XLSX.utils.json_to_sheet(exportData);
                let wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Campaigns");

                XLSX.writeFile(wb, "walmart_over_utilized.xlsx");
            });

        });
    </script>
@endsection
