@extends('layouts.vertical', ['title' => 'Forecast Analysis'])

@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
    <link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
    <style>
        .tabulator .tabulator-footer {
            background: #f4f7fa;
            border-top: 1px solid #262626;
            font-size: 1rem;
            color: #4b5563;
            padding: 5px;
            height: 70px;
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

        #image-hover-preview {
            transition: opacity 0.2s ease;
        }
    </style>
@endsection

@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Forecast Analysis',
        'sub_title' => 'Forecast Analysis',
    ])

    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body">
                    <div class="mb-3 d-flex align-items-center gap-3">
                        <!-- Play/Pause Controls -->
                        <div class="d-flex align-items-center me-3">
                            <div class="btn-group time-navigation-group" role="group" aria-label="Parent navigation">
                                <button id="play-backward" class="btn btn-light rounded-circle shadow-sm me-1"
                                    style="width: 36px; height: 36px; padding: 6px;">
                                    <i class="fas fa-step-backward"></i>
                                </button>

                                <button id="play-pause" class="btn btn-light rounded-circle shadow-sm me-1"
                                    style="width: 36px; height: 36px; padding: 6px; display: none;">
                                    <i class="fas fa-pause"></i>
                                </button>

                                <button id="play-auto" class="btn btn-primary rounded-circle shadow-sm me-1"
                                    style="width: 36px; height: 36px; padding: 6px;">
                                    <i class="fas fa-play"></i>
                                </button>

                                <button id="play-forward" class="btn btn-light rounded-circle shadow-sm"
                                    style="width: 36px; height: 36px; padding: 6px;">
                                    <i class="fas fa-step-forward"></i>
                                </button>
                            </div>
                        </div>

                        <div class="d-flex align-items-center flex-wrap gap-2">
                            <!-- Column Management -->
                            <div class="dropdown">
                                <button class="btn btn-sm btn-primary dropdown-toggle d-flex align-items-center gap-1"
                                    type="button" id="hide-column-dropdown" data-bs-toggle="dropdown">
                                    <i class="bi bi-grid-3x3-gap-fill"></i>
                                    Manage Columns
                                </button>
                                <ul class="dropdown-menu p-3 shadow-lg border rounded-3" id="column-dropdown-menu"
                                    style="max-height: 300px; overflow-y: auto; min-width: 250px;">
                                    <li class="fw-semibold text-muted mb-2">Toggle Columns</li>
                                </ul>
                            </div>

                            <!-- 2 ORDER Color Filter -->
                            <div class="dropdown">
                                <button class="btn btn-sm btn-outline-warning dropdown-toggle d-flex align-items-center gap-1"
                                    type="button" id="order-color-filter-dropdown" data-bs-toggle="dropdown">
                                    <i class="bi bi-funnel-fill"></i>
                                    2 ORDER
                                </button>
                                <ul class="dropdown-menu p-2 shadow-lg border rounded-3">
                                    <li><button class="dropdown-item" data-filter="">All</button></li>
                                    <li><button class="dropdown-item" data-filter="yellow">🟡 Yellow</button></li>
                                </ul>
                            </div>

                            <!-- Yellow Count -->
                            <div id="yellow-count-container" class="d-none px-2 btn btn-sm rounded-2 shadow-sm border border-danger bg-danger">
                                <div class="d-flex align-items-center gap-1">
                                    <i class="bi bi-star-fill text-white"></i>
                                    <span id="yellow-count-box" class="fw-semibold text-white">Approval Pending: 0</span>
                                </div>
                            </div>


                            <!-- Show All Columns -->
                            <button id="show-all-columns-btn"
                                class="btn btn-sm btn-outline-success d-flex align-items-center gap-1">
                                <i class="bi bi-eye"></i>
                                Show All
                            </button>

                            <!-- Toggle NR -->
                            <button id="toggle-nr-rows" class="btn btn-sm btn-outline-secondary">
                                Show NR
                            </button>

                            <!-- Row Type Filter -->
                            <select id="row-data-type" class="form-select-sm border border-primary" style="width: 150px;">
                                <option value="all">🔁 Show All</option>
                                <option value="sku">🔹 SKU (Child)</option>
                                <option value="parent">🔸 Parent</option>
                            </select>

                            <button id="total-transit" class="btn btn-sm btn-info">
                                Show Transit
                            </button>

                            <button id="restock_needed" class="btn btn-sm btn-warning fw-semibold text-white">
                                Restock Needed: <span id = "total_restock" class="fw-semibold text-white">0</span>
                            </button>
                        </div>
                    </div>

                    <div id="forecast-table"></div>
                </div>
            </div>
        </div>
    </div>
    {{-- month view modal --}}
    <div class="modal fade" id="monthModal" tabindex="-1" aria-labelledby="monthModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content border-0 shadow-sm">
                <div class="modal-header bg-info text-white d-flex justify-content-between align-items-center">
                    <h5 class="modal-title mb-0 fw-semibold">MONTH VIEW</h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"
                        aria-label="Close"></button>
                </div>
                <div class="modal-body" id="monthModalBody">
                    <div class="d-flex justify-content-between gap-2 flex-nowrap w-100 px-3" id="monthCardWrapper"
                        style="overflow-x: auto;">
                        <!-- Month cards inserted here -->
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Metric View Modal -->
    <div class="modal fade" id="metricModal" tabindex="-1" aria-labelledby="metricModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content border-0">
                <div class="modal-header bg-info text-white">
                    <h5 class="modal-title fw-bold" id="metricModalLabel">METRIC VIEW</h5>
                    <button type="button" class="btn-close text-white" data-bs-dismiss="modal"
                        aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div id="metricCardWrapper" class="d-flex justify-content-between gap-2 flex-nowrap w-100 px-3">
                        <!-- Metric cards will be appended here -->
                    </div>
                </div>
            </div>
        </div>
    </div>

    {{-- link edit modal --}}
    <div class="modal fade" id="linkEditModal" tabindex="-1" aria-labelledby="linkEditModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered modal-lg shadow-none">
            <div class="modal-content border-0 shadow-lg">
                <div class="modal-header bg-primary text-white border-0">
                    <h5 class="modal-title" id="linkEditModalLabel">
                        <i class="fas fa-link me-2"></i>
                        <span>Edit Link</span>
                    </h5>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"
                        aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <div class="form-group">
                        <label id="linkLabel" class="form-label fw-lg mb-2">Link URL:</label>
                        <div class="input-group">
                            <span class="input-group-text bg-light border-end-0">
                                <i class="fas fa-link text-primary"></i>
                            </span>
                            <input type="url" class="form-control form-control-lg border-start-0 ps-2"
                                id="linkEditInput" placeholder="Enter URL here..." autocomplete="off"
                                spellcheck="false">
                        </div>
                    </div>
                </div>
                <div class="modal-footer border-0 pt-0">
                    <button type="button" class="btn btn-primary" id="saveLinkBtn">
                        <i class="fas fa-check me-1"></i>
                        Save Changes
                    </button>
                </div>
            </div>
        </div>
    </div>

    {{-- Edit Notes Modal --}}
    <div class="modal fade" id="editNotesModal" tabindex="-1" role="dialog" aria-labelledby="editNotesLabel"
        aria-hidden="true">
        <div class="modal-dialog modal-lg shadow-none" role="document">
            <div class="modal-content border-0 shadow-lg">
                <div class="modal-header bg-primary text-white border-0">
                    <h5 class="modal-title" id="editNotesLabel">
                        <i class="fas fa-edit me-2"></i> Edit Notes
                    </h5>
                    <button type="button" class="close text-white custom-close" data-bs-dismiss="modal"
                        aria-label="Close" style="font-size:25px; background-color: transparent; border: none;">
                        <span aria-hidden="true">&times;</span>
                    </button>
                </div>
                <div class="modal-body">
                    <textarea id="notesInput" class="form-control form-control-lg shadow-none mb-4" rows="3"
                        placeholder="Type your note here..." style="resize: vertical;"></textarea>
                    <div class="text-end">
                        <button type="button" class="btn btn-primary" id="saveNotesBtn">
                            <i class="fas fa-save me-2"></i> Save Changes
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="scouthProductsModal" tabindex="-1">
        <div class="modal-dialog modal-xl modal-dialog-scrollable">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">Scout Products View</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal"></button>
                </div>
                <div class="modal-body">
                    <!-- dynamic content here -->
                </div>
            </div>
        </div>
    </div>
@endsection

@section('script')
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <script>
        document.body.style.zoom = "95%";
        const getDilColor = (value) => {
            const percent = parseFloat(value) * 100;
            if (percent < 16.66) return 'red';
            if (percent >= 16.66 && percent < 25) return 'yellow';
            if (percent >= 25 && percent < 50) return 'green';
            return 'pink';
        };

        const getPftColor = (value) => {
            const percent = parseFloat(value) * 100;
            if (percent < 10) return 'red';
            if (percent >= 10 && percent < 15) return 'yellow';
            if (percent >= 15 && percent < 20) return 'blue';
            if (percent >= 20 && percent <= 40) return 'green';
            return 'pink';
        };

        const getRoiColor = (value) => {
            const percent = parseFloat(value) * 100;
            if (percent >= 0 && percent < 50) return 'red';
            if (percent >= 50 && percent < 75) return 'yellow';
            if (percent >= 75 && percent <= 100) return 'green';
            return 'pink';
        };

        //global variables for play btn
        let groupedSkuData = {};

        const table = new Tabulator("#forecast-table", {
            ajaxURL: "/forecast-analysis-data-view",
            ajaxConfig: "GET",
            layout: "fitDataFill",
            pagination: true,
            paginationSize: 200,
            paginationCounter: "rows",
            movableColumns: false,
            resizableColumns: true,
            height: "650px",
            index: "SKU",
            rowFormatter: function(row) {
                const data = row.getData();
                const sku = data["SKU"] || '';

                if (sku.toUpperCase().includes("PARENT")) {
                    row.getElement().classList.add("parent-row");
                }
            },
            columns: [{
                    title: "#",
                    field: "Image",
                    headerSort: false,
                    formatter: function(cell) {
                        const url = cell.getValue();
                        if (!url) return `<span class="text-muted">N/A</span>`;
                        return `<img 
                        src="${url}" 
                        data-full="${url}" 
                        class="hover-thumb" 
                        style="width:30px;height:30px;border-radius:6px;object-fit:contain;box-shadow:0 1px 4px #0001;cursor: pointer;"
                    >`;
                    },
                    cellMouseOver: function(e, cell) {
                        const img = cell.getElement().querySelector('.hover-thumb');
                        if (!img) return;

                        const fullUrl = img.getAttribute('data-full');

                        let preview = document.createElement('div');
                        preview.id = 'image-hover-preview';
                        preview.style.position = 'fixed';
                        preview.style.top = `${e.clientY + 10}px`;
                        preview.style.left = `${e.clientX + 10}px`;
                        preview.style.zIndex = 9999;
                        preview.style.border = '1px solid #ccc';
                        preview.style.background = '#fff';
                        preview.style.padding = '4px';
                        preview.style.boxShadow = '0 2px 8px rgba(0,0,0,0.2)';
                        preview.innerHTML =
                            `<img src="${fullUrl}" style="max-width:350px;max-height:350px;">`;

                        document.body.appendChild(preview);
                    },
                    cellMouseOut: function(e, cell) {
                        const preview = document.getElementById('image-hover-preview');
                        if (preview) preview.remove();
                    },
                },
                {
                    title: "Parent",
                    field: "Parent",
                    minWidth: 130,
                    headerFilter: "input",
                    headerFilterPlaceholder: "Search parent.",
                    headerFilterFunc: "like",
                    accessor: row => row["Parent"]
                },
                {
                    title: "SKU",
                    field: "SKU",
                    minWidth: 130,
                    headerFilter: "input",
                    headerFilterPlaceholder: "Search sku.",
                    headerFilterFunc: "like",
                    accessor: row => row["SKU"]
                },
                {
                    title: "INV",
                    field: "INV",
                    accessor: row => row["INV"],
                    formatter: function(cell) {
                        const value = cell.getValue();
                        return `<span style="display:block; text-align:center;">${value}</span>`;
                    }
                },
                {
                    title: "OV L30",
                    field: "L30",
                    accessor: row => row["L30"],
                    formatter: function(cell) {
                        const value = cell.getValue();
                        return `<span style="display:block; text-align:center;">${value}</span>`;
                    }
                },
                {
                    title: "DIL",
                    field: "ov_dil",
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
                    }
                },

                {
                    title: "MSL",
                    field: "msl",
                    formatter: function(cell) {
                        const value = cell.getValue() || 0;
                        return `
                        <div style="text-align:center; font-weight:bold;">
                            ${value}
                            <button class="btn btn-sm btn-link text-info open-month-modal" style="padding: 0 4px;" title="View Monthly">
                                <i class="bi bi-calendar3"></i>
                            </button>
                        </div>
                    `;
                    },
                    cellClick: function(e, cell) {
                        if (e.target.closest(".open-month-modal")) {
                            const row = cell.getRow().getData();
                            const monthData = {
                                "JAN": row["jan"],
                                "FEB": row["feb"],
                                "MAR": row["Mar"],
                                "APR": row["Apr"],
                                "MAY": row["May"],
                                "JUN": row["Jun"],
                                "JUL": row["Jul"],
                                "AUG": row["Aug"],
                                "SEP": row["Sep"],
                                "OCT": row["Oct"],
                                "NOV": row["Nov"],
                                "DEC": row["Dec"]
                            };
                            openMonthModal(monthData);
                        }
                    }
                },
                {
                    title: "S-MSL",
                    field: "s_msl",
                    headerSort: false,
                    formatter: function(cell) {
                        const value = cell.getValue();
                        const rowData = cell.getRow().getData();

                        const sku = rowData.SKU ?? '';
                        const parent = rowData.Parent ?? '';

                        return `<div 
                        class="editable-qty" 
                        contenteditable="true" 
                        data-field="S-MSL"
                        data-original="${value ?? ''}" 
                        data-sku='${sku}' 
                        data-parent='${parent}' 
                        style="outline:none; min-width:50px; text-align:center;">
                        ${value ?? ''}
                    </div>`;
                    }
                },
                {
                    title: "MIP",
                    field: "order_given",
                    accessor: row => (row ? row["order_given"] : null),
                    sorter: "number",
                    headerSort: true,
                    formatter: function(cell) {
                        const value = cell.getValue();
                        const rowData = cell.getRow().getData();

                        const sku = rowData.SKU ?? '';
                        const parent = rowData.Parent ?? '';

                        return `<div 
                        class="editable-qty" 
                        contenteditable="true" 
                        data-field="order_given" 
                        data-original='${value ?? ''}' 
                        data-sku='${sku}' 
                        data-parent='${parent}' 
                        style="outline:none; min-width:40px; text-align:center; font-weight:bold;">
                        ${value ?? ''}
                    </div>`;
                    }
                },
                {
                    title: "R2S",
                    field: "readyToShipQty",
                    accessor: row => (row ? row["readyToShipQty"] : null),
                    sorter: "number",
                    headerSort: true,
                    formatter: function(cell) {
                        const value = cell.getValue();
                        
                        return value ?? '';
                    }
                },
                {
                    title: "Transit",
                    field: "transit",
                    accessor: row => (row ? row["transit"] : null),
                    sorter: "number",
                    headerSort: true,
                    formatter: function(cell) {
                        const value = cell.getValue();
                        const rowData = cell.getRow().getData();

                        const sku = rowData.SKU ?? '';
                        const parent = rowData.Parent ?? '';

                        return `<div 
                            class="editable-qty" 
                            contenteditable="true" 
                            data-field="Transit" 
                            data-original="${value ?? ''}" 
                            data-sku='${sku}' 
                            data-parent='${parent}' 
                            style="outline:none; min-width:40px; text-align:center; font-weight:bold;">
                            ${value ?? ''}
                        </div>`;
                    }
                },
                {
                    title: "2 ORDER",
                    field: "to_order",
                    formatter: function(cell) {
                        const value = cell.getValue();
                        const isNegative = value < 0;

                        return `<div style="text-align: center;">
                        <span style="
                            background-color: ${isNegative ? '#dc3545' : '#ffc107'};
                            color: ${isNegative ? 'white' : 'black'};
                            padding: 2px 6px;
                            border-radius: 4px;
                            display: inline-block;
                            min-width: 30px;
                            text-align: center;
                            font-weight: bold;
                        ">${value}</span>
                    </div>`;
                    }
                },
                {
                    title: "Appr. QTY",
                    field: "Approved QTY",
                    accessor: row => row?.["Approved QTY"] ?? null,
                    headerSort: false,
                    formatter: function(cell) {
                        const value = cell.getValue();
                        const rowData = cell.getRow().getData();

                        const sku = rowData.SKU ?? '';
                        const parent = rowData.Parent ?? '';

                        return `<div 
                        class="editable-qty" 
                        contenteditable="true" 
                        data-field="Approved QTY" 
                        data-original="${value ?? ''}" 
                        data-sku='${sku}' 
                        data-parent='${parent}' 
                        style="outline:none; min-width:40px; text-align:center; font-weight:bold;">
                        ${value ?? ''}
                    </div>`;
                    }
                },
                {
                    title: "Supplier",
                    field: "Supplier Tag",
                    accessor: row => row["Supplier Tag"]
                },
                {
                    title: "NRP",
                    field: "nr",
                    accessor: row => row ? (row["nr"] ?? null) : null,
                    headerSort: false,
                    formatter: function(cell) {
                        const value = cell.getValue() ?? '';
                        const rowData = cell.getRow().getData();

                        return `
                        <select class="form-select form-select-sm editable-select"
                            data-type="NR"
                            data-sku='${rowData["SKU"]}'
                            data-parent='${rowData["Parent"]}'
                            style="width: auto; min-width: 85px; padding: 4px 24px 4px 8px;
                                font-size: 0.875rem; border-radius: 4px; border: 1px solid #dee2e6;
                                background-color: #fff;">
                            <option value="REQ" ${value === 'REQ' ? 'selected' : ''}>REQ</option>
                            <option value="NR" ${value === 'NR' ? 'selected' : ''}>NR</option>
                            <option value="LATER" ${value === 'LATER' ? 'selected' : ''}>LATER</option>
                        </select>
                    `;
                    }
                },
                {
                    title: "Hide",
                    field: "hide",
                    accessor: row => row?.["hide"] ?? null,
                    headerSort: false,
                    formatter: function(cell) {
                        const value = cell.getValue() ?? '';
                        const rowData = cell.getRow().getData();

                        return `
                        <select class="form-select form-select-sm editable-select"
                            data-type="Hide"
                            data-sku='${rowData["SKU"]}'
                            data-parent='${rowData["Parent"]}'
                            style="width: auto; min-width: 100px; padding: 4px 24px 4px 8px;
                                font-size: 0.875rem; border-radius: 4px; border: 1px solid #dee2e6;
                                background-color: #fff;">
                            <option value="">Select</option>
                            <option value="@Need" ${value === '@Need' ? 'selected' : ''}>@Need</option>
                            <option value="@Taken" ${value === '@Taken' ? 'selected' : ''}>@Taken</option>
                            <option value="@Senior" ${value === '@Senior' ? 'selected' : ''}>@Senior</option>
                        </select>
                    `;
                    }
                },
            ],
            ajaxResponse: function(url, params, response) {
                groupedSkuData = {}; // clear previous

                const groupedMSL = {};
                const groupedS_MSL = {};

                const processed = response.data.map((item, index) => {
                    const sku = item["SKU"] || "";
                    const parentKey = item["Parent"] || "";

                    const total = parseFloat(item["Total"]) || 0;
                    const totalMonth = parseFloat(item["Total month"]) || 0;

                    const inv = parseFloat(item["INV"]) || 0;
                    const transit = parseFloat(item["Transit"] ?? item["transit"]) || 0;
                    const orderGiven = parseFloat(item["order_given"] ?? item["Order Given"]) || 0;
                    const msl = totalMonth > 0 ? (total / totalMonth) * 4 : 0;

                    const toOrder = Math.round(msl - inv - transit - orderGiven);

                    // if (toOrder == 0) {
                    //     return false;
                    // }

                    if (!groupedMSL[parentKey]) groupedMSL[parentKey] = 0;
                    groupedMSL[parentKey] += msl;

                    const s_msl_val = parseFloat(item["s-msl"]) || 0;
                    if (!groupedS_MSL[parentKey]) groupedS_MSL[parentKey] = 0;
                    groupedS_MSL[parentKey] += s_msl_val;

                    const isParent = item.is_parent === true || item.is_parent === "true" || sku.toUpperCase().includes("PARENT");

                    const processedItem = {
                        ...item,
                        sl_no: index + 1,
                        pft_percent: item['pft%'] ?? null,
                        msl: Math.round(msl),
                        to_order: toOrder,
                        parentKey: parentKey,
                        s_msl: s_msl_val,
                        is_parent: isParent,
                        isParent: isParent,
                        raw_data: item || {}
                    };

                    // Group for play button use
                    if (!groupedSkuData[parentKey]) groupedSkuData[parentKey] = [];
                    groupedSkuData[parentKey].push(processedItem);

                    return processedItem;
                });

                // update parent rows
                processed.forEach(row => {
                    if (row.isParent) {
                        const parentKey = row.parentKey;
                        const children = groupedSkuData[parentKey] || [];
                    }
                });

                setTimeout(() => {
                    setCombinedFilters();
                }, 0);
                return processed;
            },

            ajaxError: function(xhr, textStatus, errorThrown) {
                console.error("Error loading data:", textStatus);
            },
        });

        let currentParentFilter = null;
        let currentColorFilter = null;
        let hideNRYes = true;
        let currentRowTypeFilter = 'all';
        let currentRestockFilter = false;

        function setCombinedFilters() {
            const allData = table.getData();
            const groupedChildrenMap = {};
            const visibleParentKeys = new Set();

            // Calculate total restock count
            const restockCount = allData.filter(item => {
                const invValue = item.raw_data ? item.raw_data["INV"] : item["INV"];
                const inv = parseFloat(invValue) || 0;
                return !item.is_parent && inv === 0;
            }).length;
            document.getElementById('total_restock').textContent = restockCount;

            // Group all children by parent
            allData.forEach(item => {
                if (!item.is_parent) {
                    const key = item.Parent;
                    if (!groupedChildrenMap[key]) groupedChildrenMap[key] = [];
                    groupedChildrenMap[key].push(item);
                }
            });

            // Determine which parents should be visible
            Object.keys(groupedChildrenMap).forEach(parentKey => {
                const children = groupedChildrenMap[parentKey];

                const matchingChildren = children.filter(child => {
                    const nrMatch = !hideNRYes || child.nr !== 'NR';
                    let filterMatch = true;
                    if (currentRestockFilter) {
                        const invValue = child.raw_data ? child.raw_data["INV"] : child["INV"];
                        const l30Value = child.raw_data ? child.raw_data["L30"] : child["L30"];
                        const inv = parseFloat(invValue) || 0;
                        const l30 = parseFloat(l30Value) || 0;
                        const dilOver100 = inv === 0 || (inv > 0 && l30 / inv > 1);
                        filterMatch = dilOver100;
                    } 
                    else {
                        filterMatch = currentColorFilter === 'red' ?
                            child.to_order < 0 :
                            currentColorFilter === 'yellow' ?
                            child.to_order >= 0 :
                            true;
                    }
                    return nrMatch && filterMatch;
                });

                if (matchingChildren.length > 0) {
                    visibleParentKeys.add(parentKey);
                }
            });

            table.setFilter(function(row) {
                const data = typeof row.getData === 'function' ? row.getData() : row;

                const isChild = !data.is_parent;
                const isParent = data.is_parent;

                let matchesFilter = true;
                if (currentRestockFilter) {
                    const invValue = data.raw_data ? data.raw_data["INV"] : data["INV"];
                    const l30Value = data.raw_data ? data.raw_data["L30"] : data["L30"];
                    const inv = parseFloat(invValue) || 0;
                    const l30 = parseFloat(l30Value) || 0;
                    const dilOver100 = inv === 0 || (inv > 0 && l30 / inv > 1);
                    matchesFilter = dilOver100;
                } else {
                    matchesFilter = currentColorFilter === 'red' ?
                        data.to_order < 0 :
                        currentColorFilter === 'yellow' ?
                        data.to_order >= 0 :
                        true;
                }

                const matchesNR = hideNRYes ? data.nr !== 'NR' : true;

                // 🎯 Force filter to one parent group if play mode is active
                if (currentParentFilter) {
                    if (isParent) {
                        return data.Parent === currentParentFilter;
                    } else {
                        return data.Parent === currentParentFilter && matchesFilter && matchesNR;
                    }
                }

                if (isChild) {
                    const showChild = matchesFilter && matchesNR;
                    if (currentRowTypeFilter === 'parent') return false;
                    if (currentRowTypeFilter === 'sku') return showChild;
                    return showChild;
                }

                if (isParent) {
                    const showParent =
                        currentRowTypeFilter === 'parent' ?
                        true :
                        visibleParentKeys.has(data.Parent);
                    if (currentRowTypeFilter === 'sku') return false;
                    return showParent;
                }

                return false;
            });

            // update visible count
            setTimeout(() => updateParentTotalsBasedOnVisibleRows(), 50);

            const visibleRows = table.getRows(true).map(r => r.getData());
            const yellowCount = visibleRows.filter(r =>
                r.to_order >= 0 &&
                !r.is_parent &&
                r.nr !== 'NR'
            ).length;

            document.getElementById('yellow-count-box').textContent = `Approval Pending: ${yellowCount}`;
            document.getElementById('toggle-nr-rows').textContent = hideNRYes ? "Show NR" : "Hide NR";
        }

        function updateParentTotalsBasedOnVisibleRows() {
            const visibleRows = table.getRows(true);
            const parentGroups = {};

            visibleRows.forEach(row => {
                const data = row.getData();
                const parent = data.Parent;
                if (!parent) return;

                if (!parentGroups[parent]) {
                    parentGroups[parent] = {
                        approved: 0,
                        inv: 0,
                        l30: 0,
                        orderGiven: 0,
                        transit: 0,
                        toOrder: 0,
                        parentRow: null
                    };
                }

                if (data.is_parent) {
                    // ✅ Skip update if already updated in ajaxResponse
                    parentGroups[parent].parentRow = row;
                } else {
                    const approvedValue = data.raw_data ? data.raw_data["Approved QTY"] : data["Approved QTY"];
                    const invValue = data.raw_data ? data.raw_data["INV"] : data["INV"];
                    const l30Value = data.raw_data ? data.raw_data["L30"] : data["L30"];
                    const orderGivenValue = data.raw_data ? data.raw_data["order_given"] : data["order_given"];
                    const transitValue = data.raw_data ? data.raw_data["transit"] : data["transit"];
                    const toOrderValue = data.raw_data ? data.raw_data["to_order"] : data["to_order"];

                    parentGroups[parent].approved += parseFloat(approvedValue) || 0;
                    parentGroups[parent].inv += parseFloat(invValue) || 0;
                    parentGroups[parent].l30 += parseFloat(l30Value) || 0;
                    parentGroups[parent].orderGiven += parseFloat(orderGivenValue) || 0;
                    parentGroups[parent].transit += parseFloat(transitValue) || 0;
                    parentGroups[parent].toOrder += parseFloat(toOrderValue) || 0;
                }
            });

            Object.values(parentGroups).forEach(group => {
                if (group.parentRow) {
                    const parentData = group.parentRow.getData();

                    // ✅ Only update if current values are null or 0
                    const alreadySet =
                        parentData["to_order"] !== undefined && parentData["to_order"] !== null;

                    if (!alreadySet) {
                        group.parentRow.update({
                            "Approved QTY": group.approved,
                            "INV": group.inv,
                            "L30": group.l30,
                            "order_given": group.orderGiven,
                            "transit": group.transit,
                            "to_order": group.toOrder
                        });
                    }
                }
            });
        }

        //modals
        function openMonthModal(monthData) {
            const wrapper = document.getElementById("monthCardWrapper");
            if (!wrapper) return;

            wrapper.innerHTML = ""; // Clear previous content

            const monthOrder = [
                "JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"
            ];

            // Sort and display in month order
            monthOrder.forEach(month => {
                const value = monthData[month] ?? 0;

                const card = document.createElement("div");
                card.className = "month-card";

                const title = document.createElement("div");
                title.className = "month-title";
                title.innerText = month;

                const count = document.createElement("div");
                count.className = "month-value";
                count.innerText = value;

                card.appendChild(title);
                card.appendChild(count);
                wrapper.appendChild(card);
            });

            const modal = new bootstrap.Modal(document.getElementById("monthModal"));
            modal.show();
        }

        function openMetricModal(row) {
            const metricData = {
                SH: row['SH'],
                CP: row['CP'],
                LP: row['LP'],
                Freight: row['Freight'],
                "GW (KG)": row['GW (KG)'],
                "GW (LB)": row['GW (LB)'],
                "CBM MSL": row['CBM MSL'],
            };

            const wrapper = document.getElementById("metricCardWrapper");
            wrapper.innerHTML = "";

            for (const [key, value] of Object.entries(metricData)) {
                const displayValue = (!isNaN(value) && value !== null) ? parseFloat(value).toFixed(2) : '-';

                const card = document.createElement("div");
                card.className = "month-card";
                card.innerHTML = `
                <div class="month-title">${key}</div>
                <div class="month-value">${displayValue}</div>
            `;
                wrapper.appendChild(card);
            }

            const modal = new bootstrap.Modal(document.getElementById('metricModal'));
            modal.show();
        }

        const COLUMN_VIS_KEY = "tabulator_column_visibility";

        function buildColumnDropdown() {
            const menu = document.getElementById("column-dropdown-menu");
            menu.innerHTML = '';

            const savedVisibility = JSON.parse(localStorage.getItem(COLUMN_VIS_KEY) || '{}');

            const columns = table.getColumns().filter(col => col.getField());

            columns.forEach(col => {
                const field = col.getField();
                const title = col.getDefinition().title;

                // Apply saved visibility on table
                if (savedVisibility[field] === false) {
                    col.hide();
                } else {
                    col.show();
                }

                const li = document.createElement("li");
                const div = document.createElement("div");
                div.className = "form-check d-flex align-items-center gap-2 py-1 px-2 rounded hover-bg-light";

                const input = document.createElement("input");
                input.className = "form-check-input shadow-sm cursor-pointer";
                input.type = "checkbox";
                input.id = `col-${field}`;
                input.value = field;
                input.checked = col.isVisible();
                input.style.cssText = `
                    width: 18px;
                    height: 18px;
                    cursor: pointer;
                    border-color: #dee2e6;
                `;

                const label = document.createElement("label");
                label.className = "form-check-label cursor-pointer mb-0 text-dark";
                label.htmlFor = `col-${field}`;
                label.innerText = title;
                label.style.cssText = `
                    cursor: pointer;
                    font-size: 0.9rem;
                    user-select: none;
                `;

                // Add hover effect
                div.addEventListener('mouseover', () => {
                    div.style.backgroundColor = '#f8f9fa';
                });

                div.addEventListener('mouseout', () => {
                    div.style.backgroundColor = 'transparent';
                });

                // Add ripple effect on click
                div.addEventListener('click', (e) => {
                    if (e.target !== input) {
                        input.click();
                    }
                });

                div.appendChild(input);
                div.appendChild(label);
                li.appendChild(div);
                menu.appendChild(li);
            });
        }

        function saveColumnVisibilityToLocalStorage() {
            const visibility = {};
            table.getColumns().forEach(col => {
                const field = col.getField();
                if (field) {
                    visibility[field] = col.isVisible();
                }
            });
            localStorage.setItem(COLUMN_VIS_KEY, JSON.stringify(visibility));
        }

        document.addEventListener("DOMContentLoaded", () => {
            buildColumnDropdown();

            // Toggle column from dropdown
            document.getElementById("column-dropdown-menu").addEventListener("change", function(e) {
                if (e.target.type === "checkbox") {
                    const field = e.target.value;
                    const col = table.getColumn(field);
                    if (col) {
                        e.target.checked ? col.show() : col.hide();
                        saveColumnVisibilityToLocalStorage();
                    }
                }
            });

            // Show All Columns button
            document.getElementById("show-all-columns-btn").addEventListener("click", function() {
                const checkboxes = document.querySelectorAll(
                    "#column-dropdown-menu input[type='checkbox']");
                checkboxes.forEach(cb => {
                    cb.checked = true;
                    const col = table.getColumn(cb.value);
                    if (col) col.show();
                });
                saveColumnVisibilityToLocalStorage();
            });

            // Handle editable field
            $(document).off('blur', '.editable-qty').on('blur', '.editable-qty', function() {
                const $cell = $(this);
                const newValueRaw = $cell.text().trim();
                const originalValue = ($cell.data('original') ?? '').toString().trim();
                const field = $cell.data('field');
                const sku = $cell.data('sku');
                const parent = $cell.data('parent');

                // Convert raw value to number safely
                const newValue = ['Approved QTY', 'S-MSL', 'order_given'].includes(field) ?
                    Number(newValueRaw) :
                    newValueRaw;

                const original = ['Approved QTY', 'S-MSL', 'order_given'].includes(field) ?
                    Number(originalValue) :
                    originalValue;

                // Avoid unnecessary updates
                if (newValue === original) return;

                // Numeric validation
                if (['Approved QTY', 'S-MSL', 'order_given'].includes(field) && isNaN(newValue)) {
                    alert('Please enter a valid number.');
                    $cell.text(originalValue); // revert
                    return;
                }

                // Optional validation for date fields (YYYY-MM-DD)
                if (['Date of Appr'].includes(field)) {
                    const isValidDate = /^\d{4}-\d{2}-\d{2}$/.test(newValue);
                    if (!isValidDate) {
                        alert('Please enter a valid date in YYYY-MM-DD format.');
                        $cell.text(originalValue);
                        return;
                    }
                }

                updateForecastField({
                    sku,
                    parent,
                    column: field,
                    value: newValue
                }, function() {
                    $cell.data('original', newValue);

                    if (field === 'Approved QTY') {
                        const today = new Date();
                        const currentDate = today.getFullYear() + '-' + String(today.getMonth() + 1)
                            .padStart(2, '0') + '-' + String(today.getDate()).padStart(2, '0');

                        updateForecastField({
                            sku,
                            parent,
                            column: 'Date of Appr',
                            value: currentDate
                        }, function() {
                            const row = table.getRows().find(r => r.getData().SKU === sku &&
                                r.getData().Parent === parent);
                            // if (row) {
                            //     row.delete();
                            // }
                        });
                    }
                    setCombinedFilters();
                }, function() {
                    $cell.text(originalValue);
                });

            });

            // Handle link edit modal save
            $('#saveLinkBtn').on('click', function() {
                const newValue = $('#linkEditInput').val().trim();
                const field = editingField;
                const sku = editingRow['SKU'];
                const parent = editingRow['Parent'];

                editingRow[field] = newValue;

                const iconMap = {
                    'Clink': `<i class="fas fa-link text-primary me-1"></i>`,
                    'Olink': `<i class="fas fa-external-link-alt text-success me-1"></i>`,
                    'rfq_form_link': `<i class="fas fa-file-contract text-success me-1"></i>`,
                    'rfq_report': `<i class="fas fa-file-alt text-info me-1"></i>`
                };


                const iconHtml = newValue ?
                    `<a href="${newValue}" target="_blank" title="${field}">${iconMap[field] || ''}</a>` :
                    '';

                const editIcon = `<a href="#" class="edit-${field.toLowerCase()}" title="Edit ${field}">
                                    <i class="fas fa-edit text-warning"></i>
                                </a>`;

                $(editingLinkCell).html(`
                    <div class="d-flex align-items-center justify-content-center gap-1 ${field.toLowerCase()}-cell">
                        ${iconHtml}${editIcon}
                    </div>
                `);

                $('#linkEditModal').modal('hide');

                updateForecastField({
                        sku,
                        parent,
                        column: field,
                        value: newValue
                    },
                    function() {
                        console.log(`${field} saved successfully.`);
                    },
                    function() {
                        alert(`Failed to save ${field}.`);
                    }
                );
            });

            // Handle editable select field
            $(document).off('change', '.editable-select, .editable-date').on('change',
                '.editable-select, .editable-date',
                function() {
                    const $el = $(this);
                    const isSelect = $el.hasClass('editable-select');
                    const isDate = $el.hasClass('editable-date');

                    const newValue = $el.val().trim();
                    const sku = $el.data('sku');
                    const parent = $el.data('parent');
                    const field = isSelect ? $el.data('type') : $el.data('field');
                    const originalValue = isDate ? $el.data('original') : null;

                    // For date input: skip if no change
                    if (isDate && newValue === originalValue) return;

                    updateForecastField({
                            sku,
                            parent,
                            column: field,
                            value: newValue
                        },
                        function() {
                            if (isDate) {
                                $el.data('original', newValue); // update reference
                            }
                            if (field === 'NR') {
                                const row = table.getRows().find(r =>
                                    r.getData().SKU === sku && r.getData().Parent === parent
                                );
                                if (row) row.update({
                                    nr: newValue
                                });

                                setCombinedFilters();
                            }
                        },
                        function() {
                            if (isDate) {
                                $el.val(originalValue); // revert on fail
                            }
                            alert(`Failed to save ${field}.`);
                        }
                    );
                });

            // Handle notes edit modal save
            $('#saveNotesBtn').on('click', function() {
                const newValue = $('#notesInput').val().trim();
                const field = editingField;
                const sku = editingRow['SKU'];
                const parent = editingRow['Parent'];

                editingRow[field] = newValue;

                // Update DOM cell content
                const display = newValue ? newValue.substring(0, 30) + (newValue.length > 30 ? '...' : '') :
                    '<em class="text-muted">No notes</em>';

                const updatedHTML = `
                    <div class="d-flex align-items-center justify-content-between notes-cell">
                        <span class="text-truncate" title="${newValue}">${display}</span>
                        <a href="#" class="edit-notes ms-2" title="Edit Notes">
                            <i class="fas fa-edit text-warning"></i>
                        </a>
                    </div>
                `;

                $(editingLinkCell).html(updatedHTML);
                $('#editNotesModal').modal('hide');

                updateForecastField({
                        sku,
                        parent,
                        column: 'Notes',
                        value: newValue
                    },
                    () => {
                        $('#editNotesModal').modal('hide');

                        const cell = $(`.edit-notes-btn[data-sku="${sku}"][data-parent="${parent}"]`)
                            .closest('td');

                        if (cell.length === 0) {
                            console.warn('Cell not found for SKU:', sku, 'and Parent:', parent);
                            return;
                        }

                        cell.empty();

                        const viewBtn = $('<i>')
                            .addClass('fas fa-eye text-info ms-2 view-note-btn')
                            .css('cursor', 'pointer')
                            .attr('title', 'View Note')
                            .attr('data-note', newValue);

                        const editBtn = $('<i>')
                            .addClass('fas fa-edit text-primary ms-2 edit-notes-btn')
                            .css('cursor', 'pointer')
                            .attr('title', 'Edit Note')
                            .attr('data-note', newValue)
                            .attr('data-sku', sku)
                            .attr('data-parent', parent);

                        cell.append(viewBtn, editBtn);
                    },
                    () => {
                        alert('Failed to save note.');
                    }
                );

            });

            // Reusable AJAX call
            function updateForecastField(data, onSuccess = () => {}, onFail = () => {}) {
                $.post('/update-forecast-data', {
                    ...data,
                    _token: $('meta[name="csrf-token"]').attr('content')
                }).done(res => {
                    if (res.success) {
                        console.log('Saved:', res.message);
                        onSuccess();
                    } else {
                        console.warn('Not saved:', res.message);
                        onFail();
                    }
                }).fail(err => {
                    console.error('AJAX failed:', err);
                    alert('Error saving data.');
                    onFail();
                });
            }

        });

        //play btn filter
        document.addEventListener('DOMContentLoaded', function() {
            document.documentElement.setAttribute("data-sidenav-size", "condensed");
            const table = Tabulator.findTable("#forecast-table")[0];

            const parentKeys = () => Object.keys(groupedSkuData);
            let currentIndex = 0;
            let isPlaying = false;

            function renderGroup(parentKey) {
                if (!groupedSkuData[parentKey]) return;
                currentParentFilter = parentKey;
                setCombinedFilters();
            }

            document.getElementById('play-auto').addEventListener('click', () => {
                isPlaying = true;
                currentIndex = 0;
                renderGroup(parentKeys()[currentIndex]);
                document.getElementById('play-pause').style.display = 'inline-block';
                document.getElementById('play-auto').style.display = 'none';
            });

            document.getElementById('play-forward').addEventListener('click', () => {
                if (!isPlaying) return;
                currentIndex = (currentIndex + 1) % parentKeys().length;
                renderGroup(parentKeys()[currentIndex]);
            });

            document.getElementById('play-backward').addEventListener('click', () => {
                if (!isPlaying) return;
                currentIndex = (currentIndex - 1 + parentKeys().length) % parentKeys().length;
                renderGroup(parentKeys()[currentIndex]);
            });

            document.getElementById('play-pause').addEventListener('click', () => {
                isPlaying = false;
                currentParentFilter = null; // Show all data
                setCombinedFilters();
                document.getElementById('play-pause').style.display = 'none';
                document.getElementById('play-auto').style.display = 'inline-block';
            });

            const countContainer = document.getElementById('yellow-count-container');
            countContainer.classList.add('d-none');

            // Set yellow filter as default
            currentColorFilter = '';
            document.getElementById('yellow-count-container').classList.remove('d-none');
            document.getElementById('order-color-filter-dropdown').innerHTML =
                '<i class="bi bi-funnel-fill"></i> Yellow Filter';
            setCombinedFilters();

            document.querySelectorAll('#order-color-filter-dropdown + .dropdown-menu [data-filter]').forEach(
                btn => {
                    btn.addEventListener('click', function() {
                        const filter = this.getAttribute('data-filter');
                        currentColorFilter = filter || null;
                        setCombinedFilters();

                        const buttonText = filter ? filter.charAt(0).toUpperCase() + filter.slice(1) +
                            ' Filter' : 'All';
                        document.getElementById('order-color-filter-dropdown').innerHTML =
                            `<i class="bi bi-funnel-fill"></i> ${buttonText}`;

                        const countContainer = document.getElementById('yellow-count-container');
                        if (filter === 'yellow') {
                            countContainer.classList.remove('d-none');
                        } else {
                            countContainer.classList.add('d-none');
                        }
                    });
                });

            document.getElementById('toggle-nr-rows').addEventListener('click', function() {
                hideNRYes = !hideNRYes;
                setCombinedFilters();

                if (currentColorFilter === 'yellow') {
                    const allData = table.getRows().map(r => r.getData());
                    const yellowCount = allData.filter(r => r.to_order >= 0 && !r.is_parent && (hideNRYes ?
                        r.nr !== 'NR' : true)).length;
                    document.getElementById('yellow-count-box').textContent =
                        `Approval Pending: ${yellowCount}`;
                }

                document.getElementById('toggle-nr-rows').textContent = hideNRYes ? "Show NR" : "Hide NR";

            });

            document.getElementById('row-data-type').addEventListener('change', function(e) {
                currentRowTypeFilter = e.target.value;
                setCombinedFilters();
            });


        });

        // Scout products view handler
        document.addEventListener('click', function(e) {
            const trigger = e.target.closest('.scouth-products-view-trigger');
            if (trigger) {
                e.preventDefault();
                e.stopPropagation();

                const encoded = trigger.getAttribute('data-item');
                if (encoded) {
                    try {
                        const rawData = JSON.parse(decodeURIComponent(encoded));
                        openModal(rawData, 'scouth products view');
                    } catch (err) {
                        console.error("Failed to parse rawData", err);
                    }
                }
            }
        });

        window.openModal = function(selectedItem, type) {
            try {
                if (type.toLowerCase() === 'scouth products view') {
                    const modalId = 'scouthProductsModal';
                    return openScouthProductsView(selectedItem, modalId);
                }
            } catch (error) {
                console.error("Error in openModal:", error);
                showNotification('danger', 'Failed to open details view. Please try again.');
            }
        };

        function openScouthProductsView(data, modalId) {
            const modal = document.getElementById(modalId);
            if (!modal) return;

            const title = modal.querySelector('.modal-title');
            const body = modal.querySelector('.modal-body');

            if (!data.scout_data || !data.scout_data.all_data) {
                title.textContent = 'Scout Products View Details';
                body.innerHTML = '<div class="alert alert-warning">No scout data available</div>';
                const instance = new bootstrap.Modal(modal);
                instance.show();
                return;
            }

            // Sort by price
            const sortedProducts = [...data.scout_data.all_data].sort((a, b) => {
                const priceA = parseFloat(a.price) || Infinity;
                const priceB = parseFloat(b.price) || Infinity;
                return priceA - priceB;
            });

            title.textContent = 'Scouth Products View (Sorted by Lowest Price)';

            let html = `
            <div><strong>Parent:</strong> ${data.Parent || 'N/A'} | <strong>SKU:</strong> ${data['(Child) sku'] || 'N/A'}</div>
            <div class="table-responsive mt-3">
                <table class="table table-bordered table-sm align-middle">
                    <thead class="table-light">
                        <tr>
                            <th>ID</th><th>Price</th><th>Category</th><th>Dimensions</th><th>Image</th>
                            <th>Quality Score</th><th>Parent ASIN</th><th>Product Rank</th>
                            <th>Rating</th><th>Reviews</th><th>Weight</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

            sortedProducts.forEach(product => {
                html += `
                <tr>
                    <td>${product.id || 'N/A'}</td>
                    <td>${product.price ? '$' + parseFloat(product.price).toFixed(2) : 'N/A'}</td>
                    <td>${product.category || 'N/A'}</td>
                    <td>${product.dimensions || 'N/A'}</td>
                    <td>
                        ${product.image_url ? `
                                                                <a href="${product.image_url}" target="_blank">
                                                                    <img src="${product.image_url}" width="60" height="60" style="border-radius:50%;">
                                                                </a>` : 'N/A'}
                    </td>
                    <td>${product.listing_quality_score || 'N/A'}</td>
                    <td>${product.parent_asin || 'N/A'}</td>
                    <td>${product.product_rank || 'N/A'}</td>
                    <td>${product.rating || 'N/A'}</td>
                    <td>${product.reviews || 'N/A'}</td>
                    <td>${product.weight || 'N/A'}</td>
                </tr>
            `;
            });

            html += '</tbody></table></div>';
            body.innerHTML = html;

            const instance = new bootstrap.Modal(modal);
            instance.show();
        }

        document.getElementById("total-transit").addEventListener("click", function(e) {
            currentColorFilter = null;

            setCombinedFilters();

            table.setSort([{
                column: "transit",
                dir: "desc",
                sorter: function(a, b) {
                    const aValue = a.getData().transit ? 1 : 0;
                    const bValue = b.getData().transit ? 1 : 0;
                    return bValue - aValue;
                }
            }]);
        });

        document.getElementById("restock_needed").addEventListener("click", function(e) {
            currentRestockFilter = true;
            currentColorFilter = null;

            setCombinedFilters();

            table.setSort([{
                column: "ov_dil",
                dir: "desc",
                sorter: function(a, b) {
                    const aData = a.getData();
                    const aInv = parseFloat(aData.raw_data ? aData.raw_data["INV"] : aData[
                        "INV"]) || 0;
                    const aL30 = parseFloat(aData.raw_data ? aData.raw_data["L30"] : aData[
                        "L30"]) || 0;
                    const aDil = aInv === 0 ? Infinity : (aL30 / aInv) * 100;

                    const bData = b.getData();
                    const bInv = parseFloat(bData.raw_data ? bData.raw_data["INV"] : bData[
                        "INV"]) || 0;
                    const bL30 = parseFloat(bData.raw_data ? bData.raw_data["L30"] : bData[
                        "L30"]) || 0;
                    const bDil = bInv === 0 ? Infinity : (bL30 / bInv) * 100;

                    return bDil - aDil;
                }
            }]);
        });
    </script>
@endsection
