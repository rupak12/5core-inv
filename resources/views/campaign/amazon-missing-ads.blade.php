@extends('layouts.vertical', ['title' => 'Amazon Missing Ads', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
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

        /* .tabulator .tabulator-cell:focus {
            outline: 1px solid #262626;
            background: #e0eaff;
        } */

        /* .tabulator-row:hover {
            background-color: #dbeafe !important;
        } */

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
        .stats-box {
            padding: 12px 16px;
            min-width: 130px;
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
            border: 1px solid #e5e7eb;
            transition: all 0.2s;
        }

        .stats-box:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.12);
        }

        .stats-label {
            font-size: 13px;
            font-weight: 600;
            color: #64748b;
            margin-bottom: 4px;
        }

        .stats-value {
            font-size: 20px;
            font-weight: 700;
        }

        .stats-value.primary { color: #2563eb; }
        .stats-value.danger { color: #dc2626; }
        .stats-value.success { color: #16a34a; }
    </style>
@endsection
@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Amazon Missing Ads',
        'sub_title' => 'Amazon Missing Ads',
    ])
    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body py-3">
                    <div class="mb-4">
                        <!-- Title -->
                        <h4 class="fw-bold text-primary mb-3 d-flex align-items-center">
                            <i class="fa-solid fa-chart-line me-2"></i>
                            Amazon Missing Ads
                        </h4>

                        <!-- Filters Row -->
                        <!-- Stats Row -->
                        <div class="row g-3 mb-3">
                            <div class="col-12">
                                <div class="d-flex flex-wrap gap-2 justify-content-between align-items-center">
                                    <!-- Left side controls -->
                                    <div class="d-flex gap-2">
                                        <input type="text" id="global-search" class="form-control form-select-lg border-1 border-secondary" placeholder="Search campaign...">

                                        <select id="status-filter" class="form-select form-select-lg" style="width: 140px;">
                                            <option value="">All Status</option>
                                            <option value="ENABLED">Enabled</option>
                                            <option value="PAUSED">Paused</option>
                                            <option value="ARCHIVED">Archived</option>
                                        </select>

                                        <select id="inv-filter" class="form-select form-select-lg" style="width: 200px;">
                                            <option value="">Select INV</option>
                                            <option value="ALL">ALL</option>
                                            <option value="INV_0">0 INV</option>
                                            <option value="OTHERS">OTHERS</option>
                                        </select>

                                        <select id="missingAds-filter" class="form-select form-select-lg" style="width: 180px;">
                                            <option value="">Select Missing Ads</option>
                                            <option value="Both Running">Both Running</option>
                                            <option value="KW Missing">KW Missing</option>
                                            <option value="PT Missing">PT Missing</option>
                                            <option value="Both Missing">Both Missing</option>
                                        </select>
                                    </div>

                                    <!-- Right side - Stats Boxes -->
                                    <div class="d-flex flex-wrap gap-3">
                                        <div class="stats-box">
                                            <div class="stats-label">Total SKUs</div>
                                            <div id="total-campaigns" class="stats-value primary">0</div>
                                        </div>
                                        
                                        <div class="stats-box">
                                            <div class="stats-label">Both Missing Ads</div>
                                            <div id="both-missing" class="stats-value danger">0</div>
                                        </div>

                                        <div class="stats-box">
                                            <div class="stats-label">KW Missing</div> 
                                            <div id="kw-missing" class="stats-value danger">0</div>
                                        </div>

                                        <div class="stats-box">
                                            <div class="stats-label">PT Missing</div>
                                            <div id="pt-missing" class="stats-value danger">0</div>
                                        </div>
                                        <div class="stats-box">
                                            <div class="stats-label">KW Running</div>
                                            <div id="kw-running" class="stats-value success">0</div>
                                        </div>
                                         <div class="stats-box">
                                            <div class="stats-label">PT Running</div>
                                            <div id="pt-running" class="stats-value success">0</div>
                                        </div>
                                        <div class="stats-box">
                                            <div class="stats-label">Both Ads Running</div>
                                            <div id="both-running" class="stats-value success">0</div>
                                        </div>
                                    </div>
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
    <script>
        document.addEventListener("DOMContentLoaded", function() {

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
                index: "sku",
                ajaxURL: "/amazon/missing/ads/data",
                layout: "fitData",
                movableColumns: true,
                resizableColumns: true,
                height: "700px",             
                virtualDom: true,
                rowFormatter: function(row) {
                    const data = row.getData();
                    const sku = data["sku"] || '';

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
                        title: "Missing Ads",
                        field: "missing_ads",
                        formatter: function(cell){
                            var row = cell.getRow().getData();
                            var kwCampaign = row.kw_campaign_name || '';
                            var ptCampaign = row.pt_campaign_name || '';
                            var sku = row.sku || '';
                            const isParent = sku.toUpperCase().includes("PARENT");

                            if(!isParent){
                                if(kwCampaign && ptCampaign){
                                    return `
                                        <span style="color: green;">Both Running</span>
                                        <i class="fa fa-info-circle text-primary toggle-missingAds-btn" 
                                            style="cursor:pointer; margin-left:8px;">
                                        </i>
                                    `;
                                } else if(kwCampaign){
                                    return `
                                        <span style="color: red;">PT Missing</span>
                                        <i class="fa fa-info-circle text-primary toggle-missingAds-btn" 
                                            style="cursor:pointer; margin-left:8px;">
                                        </i>
                                    `;
                                } else if(ptCampaign){
                                    return `
                                        <span style="color: red;">KW Missing</span>
                                        <i class="fa fa-info-circle text-primary toggle-missingAds-btn" 
                                            style="cursor:pointer; margin-left:8px;">
                                        </i>
                                    `;
                                } else {
                                    return `
                                        <span style="color: red;">Both Missing</span>
                                        <i class="fa fa-info-circle text-primary toggle-missingAds-btn" 
                                            style="cursor:pointer; margin-left:8px;">
                                        </i>
                                    `;
                                }
                            }
                            
                        }
                    },
                    {
                        title: "KW Campaign",
                        field: "kw_campaign_name",
                        visible: false
                    },
                    {
                        title: "PT Campaign",
                        field: "pt_campaign_name",
                        visible: false
                    }
                ],
                ajaxResponse: function(url, params, response) {
                    return response.data;
                }
            });


            table.on("tableBuilt", function () {

                // ✅ Combined Filter Function
                function combinedFilter(data) {
                    // 🔹 Global Search
                    let searchVal = ($("#global-search").val() || "").toLowerCase().trim();
                    if (searchVal) {
                        let fieldsToSearch = [
                            data.sku,
                            data.parent,
                            data.kw_campaign_name,
                            data.pt_campaign_name
                        ].map(f => (f || "").toLowerCase());

                        if (!fieldsToSearch.some(f => f.includes(searchVal))) {
                            return false;
                        }
                    }

                    // 🔹 Status Filter
                    let statusVal = $("#status-filter").val();
                    if (statusVal && data.campaignStatus !== statusVal) return false;

                    // 🔹 INV Filter
                    let invFilterVal = $("#inv-filter").val();
                    let inv = parseFloat(data.INV) || 0;
                    if (invFilterVal === "INV_0" && inv !== 0) return false;
                    if (invFilterVal === "OTHERS" && inv === 0) return false;

                    // 🔹 Missing Ads Filter
                    let missingVal = $("#missingAds-filter").val();
                    let kw = data.kw_campaign_name || "";
                    let pt = data.pt_campaign_name || "";

                    if (missingVal === "Both Running" && !(kw && pt)) return false;
                    if (missingVal === "KW Missing" && (!pt || kw)) {
                        if (pt && !kw) return true;
                        else return false;
                    }
                    if (missingVal === "PT Missing" && (!kw || pt)) {
                        if (kw && !pt) return true;
                        else return false;
                    }
                    if (missingVal === "Both Missing" && (kw || pt)) return false;

                    return true;
                }

                // ✅ Apply Filter
                table.setFilter(combinedFilter);

                // ✅ Update Stats Based on *Visible (Filtered)* Data
                function updateCampaignStats() {
                    let visibleData = table.getData("active");

                    let bothMissing = 0;
                    let kwMissing = 0;
                    let ptMissing = 0;
                    let bothRunning = 0;
                    let kwRunning = 0;
                    let ptRunning = 0;

                    visibleData.forEach(row => {
                        let kw = row.kw_campaign_name || "";
                        let pt = row.pt_campaign_name || "";

                        // Existing counts
                        if (kw && pt) bothRunning++;
                        else if (kw && !pt) ptMissing++;
                        else if (!kw && pt) kwMissing++;
                        else bothMissing++;

                        // New running counts
                        if (kw) kwRunning++;
                        if (pt) ptRunning++;
                    });

                    // Update HTML
                    $("#total-campaigns").text(visibleData.length);
                    $("#both-missing").text(bothMissing);
                    $("#kw-missing").text(kwMissing);
                    $("#pt-missing").text(ptMissing);
                    $("#both-running").text(bothRunning);

                    // New stats
                    $("#kw-running").text(kwRunning);
                    $("#pt-running").text(ptRunning);
                }


                // ✅ Trigger Update on Every Filter / Search Change
                function reapplyFiltersAndUpdate() {
                    table.setFilter(combinedFilter);
                    updateCampaignStats();
                }

                // ✅ Events
                $("#global-search").on("keyup", reapplyFiltersAndUpdate);
                $("#status-filter, #inv-filter, #missingAds-filter").on("change", reapplyFiltersAndUpdate);

                table.on("dataFiltered", updateCampaignStats);
                table.on("pageLoaded", updateCampaignStats);
                table.on("dataProcessed", updateCampaignStats);

                // ✅ Initial Stats Load
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
                if (e.target.classList.contains("toggle-missingAds-btn")) {
                    let btn = e.target;

                    let colsToToggle = ["kw_campaign_name", "pt_campaign_name"];

                    colsToToggle.forEach(colName => {
                        let col = table.getColumn(colName);
                        if (col) {
                            col.toggle();
                        }
                    });
                }
            });


            document.body.style.zoom = "78%";
        });
    </script>
@endsection
