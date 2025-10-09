@extends('layouts.vertical', ['title' => 'Inventory by Sales Value'])

@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
    <link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
    <style>
        #image-hover-preview {
            transition: opacity 0.2s ease;
        }

        /* Table Styling */
        #forecast-table {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
         
        }

        .image-preview-container {
            width: 100px;
            height: 100px;
            border: 2px solid #f0f0f0;
            border-radius: 12px;
            overflow: hidden;
            display: flex;
            justify-content: center;
            align-items: center;
            background: #fff;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }

        .image-preview-container:hover {
            transform: scale(1.05);
            box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        }

        .image-preview-container img {
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
            display: none; /* JS se toggle hoga */
            cursor: pointer;
        }

        .channel-logo {
        cursor: pointer;
        transition: transform 0.2s ease;
        }
        .channel-logo:hover {
        transform: scale(1.1);
        }


        .inventory-cell {
            font-weight: 600;
            background-color: #f8f9fa;
        }

      
        .inventory-cell.low {
            color: #dc3545;
            background-color: rgba(220, 53, 69, 0.1);
        }

        .inventory-cell.medium {
            color: #ffc107;
            background-color: rgba(255, 193, 7, 0.1);
        }

        .inventory-cell.high {
            color: #198754;
            background-color: rgba(25, 135, 84, 0.1);
        }

        /* Tooltip styles */
        .cell-tooltip {
            position: absolute;
            background: #333;
            color: white;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
            z-index: 1000;
            visibility: hidden;
            opacity: 0;
            transition: opacity 0.2s;
            white-space: nowrap;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        }

        .cell-with-tooltip:hover .cell-tooltip {
            visibility: visible;
            opacity: 1;
        }



        .tabulator .tabulator-header {
            background: linear-gradient(135deg, #2c6ed5 0%, #1a56b7 100%) !important;
            border-bottom: 2px solid #1a56b7;
            color: #ffffff;
            font-weight: bold;
            text-transform: uppercase;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
        }

        .tabulator .tabulator-header .tabulator-col {
            background-color: transparent;
            border-right: 1px solid rgba(255, 255, 255, 0.1);
            padding: 15px 18px;
            vertical-align: middle;
            color: #ffffff;
            transition: all 0.2s ease;
        }

        .tabulator .tabulator-header .tabulator-col-content {
            font-weight: 600;
            color: #ffffff;
            padding: 8px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            letter-spacing: 0.5px;
        }

        .tabulator .tabulator-header .tabulator-col-title {
            color: #ffffff;
            font-size: 13px;
            font-weight: 600;
            text-transform: uppercase;
        }

        /* Style for header filter inputs */
        .tabulator .tabulator-header .tabulator-col input {
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

        .tabulator .tabulator-header .tabulator-col input::placeholder {
            color: #8e9ab4;
            font-style: italic;
        }

        .tabulator .tabulator-header .tabulator-col input:focus {
            background-color: white;
            box-shadow: 0 0 0 2px rgba(26, 86, 183, 0.3);
            outline: none;
        }

        .tabulator .tabulator-row {
            border-bottom: 1px solid #eee;
            transition: background-color 0.2s ease;
        }

        .tabulator .tabulator-row:hover {
            background-color: #f8f9fa !important;
        }

        .tabulator .tabulator-row.parent-row {
            background-color: #f8f9fa;
            font-weight: 700;
            border-top: 2px solid #0d6efd;
        }

        .tabulator .tabulator-row.parent-row .tabulator-cell {
            color: #0d6efd;
        }

        .tabulator .tabulator-row .tabulator-cell {
            padding: 12px 8px;
            border-right: 1px solid #eee;
            position: relative;
            overflow: visible;
        }

        /* Hover info tooltip */
        .tabulator-cell[data-numeric='true']:hover::after {
            content: attr(data-full-value);
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            background: #333;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 12px;
            z-index: 1000;
            white-space: nowrap;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
        }

        /* Trending indicators */
        .trend-up::after {
            content: '↑';
            color: #22c55e;
            margin-left: 4px;
        }

        .trend-down::after {
            content: '↓';
            color: #ef4444;
            margin-left: 4px;
        }



        /* Pagination styling */
        .tabulator-footer {
            background-color: #f8f9fa;
            border-top: 2px solid #dee2e6;
            padding: 8px;
        }

        .tabulator-paginator {
            font-weight: 500;
        }

        .tabulator-page {
            margin: 0 2px;
            padding: 6px 12px;
            border-radius: 4px;
            border: 1px solid #dee2e6;
            background-color: white;
            color: #495057;
            transition: all 0.2s ease;
        }

        .tabulator-page:hover {
            background-color: #e9ecef;
            border-color: #adb5bd;
        }

        .tabulator-page.active {
            background-color: #007bff;
            border-color: #007bff;
            color: white;
        }



        /* Additional hover effects */
        .hover-effect {
            transition: transform 0.2s;
        }

        .hover-effect:hover {
            transform: scale(1.02);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }



        /* Modal enhancements */
        .modal-xl .modal-body {
            max-height: 95vh;
            overflow-y: auto;
           
        }

        .analysis-table th {
            background-color: #f8f9fa;
            position: sticky;
            top: 0;
            z-index: 1;
        }

        /* Enhanced Modal Styling */
        .modal-draggable .modal-dialog {
            cursor: move;
            margin: 0;
            pointer-events: all;
            position: fixed;
            /* left: 50%;
            top: 50%; */
            transform: translate(-50%, -50%);
        }

        .modal-content {
            box-shadow: 0 5px 15px rgba(0, 0, 0, .5);
            border: none;
            border-radius: 8px;
        }

        .modal-header.bg-gradient {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            border-radius: 8px 8px 0 0;
            padding: 1rem;
            border: none;
        }

        .market-summary {
            background-color: rgba(0, 0, 0, 0.02);
            border-bottom: 1px solid rgba(0, 0, 0, 0.08);
        }

        .summary-stats .badge {
            font-size: 0.9em;
            padding: 0.5em 1em;
            font-weight: 500;
        }

        .view-controls .btn-group .btn {
            padding: 0.375rem 0.75rem;
            transition: all 0.2s;
        }

        .view-controls .btn-group .btn.active {
            background-color: #2a5298;
            color: white;
            border-color: #2a5298;
        }

        .modal-actions .btn-light-secondary {
            color: white;
            background-color: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.2);
            transition: all 0.2s;
        }

        .modal-actions .btn-light-secondary:hover {
            background-color: rgba(255, 255, 255, 0.2);
        }

        #ovl30Modal .table {
            margin-bottom: 0;
            border-collapse: separate;
            border-spacing: 0 0.5rem;
        }

        #ovl30Modal .table thead th {
            background-color: #f8f9fa;
            border: none;
            padding: 1rem;
            font-weight: 600;
            color: #495057;
        }

        #ovl30Modal .table tbody tr {
            background-color: white;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.02);
            transition: transform 0.2s;
        }

        #ovl30Modal .table tbody tr:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.05);
        }

        #ovl30Modal .table tbody td {
            padding: 1rem;
            border: none;
            vertical-align: middle;
        }

        /* Value indicators */
        .value-indicator {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
        }

        .value-indicator .trend {
            font-size: 0.8em;
        }

        .value-indicator.positive {
            color: #198754;
        }

        .value-indicator.negative {
            color: #dc3545;
        }

        .value-indicator.neutral {
            color: #6c757d;
        }

        .modal-title {
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .modal-body {
            padding: 1.5rem;
            background: #ffffff;
        }

        /* Enhanced Table Styling for Modal */
        .modal-body .table {
            margin-bottom: 0;
        }

        .modal-body .table thead th {
            background: #f8f9fa;
            border-bottom: 2px solid #dee2e6;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.85rem;
            white-space: nowrap;
        }

        .modal-body .table tbody td {
            vertical-align: middle;
            padding: 0.75rem;
            border-color: #e9ecef;
        }

        .modal-body .table tbody tr:hover {
            background-color: #f8f9fa;
        }

        .btn-close {
            color: white;
            text-shadow: none;
            opacity: 0.8;
        }

        .btn-close:hover {
            opacity: 1;
        }

        /* Sorting styles */
        .sortable-table th[data-sort] {
            position: relative;
            cursor: pointer;
            user-select: none;
        }

        .sortable-table th[data-sort] .bi {
            font-size: 0.8em;
            margin-left: 5px;
            opacity: 0.5;
            transition: opacity 0.2s;
        }

        .sortable-table th[data-sort]:hover .bi {
            opacity: 1;
        }

        .sortable-table th.sort-asc .bi,
        .sortable-table th.sort-desc .bi {
            opacity: 1;
            color: #0d6efd;
        }

        .sortable-table tbody tr:nth-child(even) {
            background-color: rgba(0, 0, 0, 0.02);
        }

        .sortable-table tbody tr:hover {
            background-color: rgba(13, 110, 253, 0.05);
        }

        .sortable-table th.default-sort {
            background-color: rgba(13, 110, 253, 0.1);
        }

        .sortable-table th.default-sort .bi {
            color: #0d6efd;
            opacity: 1;
        }
    </style>
@endsection

@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Inventory by Sales Value',
        'sub_title' => 'Inventory by Sales Value',
    ])

    <!-- Image Preview -->
    <div id="image-hover-preview" style="display: none; position: fixed; z-index: 1000; pointer-events: none;">
        <img id="preview-image"
            style="max-width: 200px; max-height: 200px; border: 2px solid #ddd; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
    </div>

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
                        <div class="dropdown">
                            <button class="btn btn-primary dropdown-toggle d-flex align-items-center gap-1"
                                type="button" id="hide-column-dropdown" data-bs-toggle="dropdown">
                                <i class="bi bi-grid-3x3-gap-fill"></i>
                                Manage Columns
                            </button>
                            <ul class="dropdown-menu p-3 shadow-lg border rounded-3" id="column-dropdown-menu"
                                style="max-height: 300px; overflow-y: auto; min-width: 250px;">
                                <li class="fw-semibold text-muted mb-2">Toggle Columns</li>
                                <!-- Columns checkboxes dynamically append karoge -->
                            </ul>
                        </div>

                    <div class="btn-group" id="inv-filter" role="group" aria-label="Inventory Filter">
                        <input type="radio" class="btn-check" name="invFilter" id="filterAll" value="all" checked>
                        <label class="btn btn-outline-secondary" for="filterAll">All</label>
                        <input type="radio" class="btn-check" name="invFilter" id="filterZero" value="zero">
                        <label class="btn btn-outline-danger" for="filterZero">0</label>
                        <input type="radio" class="btn-check" name="invFilter" id="filterOther" value="other">
                        <label class="btn btn-outline-success" for="filterOther">Other</label>
                    </div>

                    <div class="btn-group" id="dil-filter" role="group" aria-label="Dilution Filter">
                        <input type="radio" class="btn-check" name="dilFilter" id="dilFilter10" value="10">
                        <label class="btn btn-outline-danger" for="dilFilter10">Dil ≤ 10%</label>
                        <input type="radio" class="btn-check" name="dilFilter" id="dilFilter50" value="50">
                        <label class="btn btn-outline-info" for="dilFilter50">Dil > 50%</label>
                        <input type="radio" class="btn-check" name="dilFilter" id="dilFilterClear" value="clear" checked>
                        <label class="btn btn-outline-secondary" for="dilFilterClear">Clear</label>
                    </div>

                    <div class="btn-group" id="cvr-filter" role="group" aria-label="CVR Filter">
                        <input type="radio" class="btn-check" name="cvrFilter" id="cvrFilterLow" value="low">
                        <label class="btn btn-outline-warning" for="cvrFilterLow">Low CVR (&lt; 5%)</label>
                        <input type="radio" class="btn-check" name="cvrFilter" id="cvrFilterClear" value="clear" checked>
                        <label class="btn btn-outline-secondary" for="cvrFilterClear">Clear</label>
                    </div>

                    <div class="btn-group" id="margin-filter" role="group" aria-label="Margin Filter">
                        <input type="radio" class="btn-check" name="marginFilter" id="marginFilterHigh" value="high">
                        <label class="btn btn-outline-success" for="marginFilterHigh">High Margin (&gt; 20%)</label>
                        <input type="radio" class="btn-check" name="marginFilter" id="marginFilterClear" value="clear" checked>
                        <label class="btn btn-outline-secondary" for="marginFilterClear">Clear</label>
                    </div>

                    <div class="btn-group" id="view-filter" role="group" aria-label="View Filter">
                        <input type="radio" class="btn-check" name="viewFilter" id="parentFilter" value="parent" checked>
                        <label class="btn btn-primary" for="parentFilter">Parent </label>
                        <input type="radio" class="btn-check" name="viewFilter" id="skuFilter" value="sku">
                        <label class="btn btn-primary" for="skuFilter">SKU </label>
                        <input type="radio" class="btn-check" name="viewFilter" id="bothFilter" value="both">
                        <label class="btn btn-primary" for="bothFilter">Both </label>
                    </div>
                    </div>

                    </div>

                    <div id="forecast-table"></div>

                </div>
            </div>
        </div>
    </div>

    <!-- OVL30 Modal -->
    <div class="modal fade modal-draggable" id="ovl30Modal" tabindex="-1" aria-labelledby="ovl30ModalLabel"
        aria-hidden="true">
        <div class="modal-dialog modal-xl">
            <div class="modal-content">
                <div class="modal-header bg-gradient">
                    <h5 class="modal-title d-flex align-items-center text-dark">
                        <i class="bi bi-bar-chart-line-fill me-2"></i>
                        OVL30 Analysis
                        <span id="ovl30SkuLabel" class="badge  text-danger ms-2 animate__animated animate__fadeIn fw-bold fs-3"></span>
                    </h5>
                    <div class="modal-actions">
                        <button class="btn btn-sm btn-light-secondary me-2">
                            <i class="bi bi-download"></i> Export
                        </button>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"
                            aria-label="Close"></button>
                    </div>
                </div>
                <div class="modal-body p-0">
                    <div class="row g-0">
                        <div class="col-12">
                            <div class="market-summary p-3 bg-light border-bottom position-sticky" style="top: 0; z-index: 1000; background-color: #f8f9fa !important;">
                                <div class="d-flex justify-content-between align-items-center">
                                    <div class="summary-stats">
                                        <div class="d-flex align-items-center gap-3 mb-2">
                                            <div class="input-group" style="width: 200px;">
                                                @csrf
                                                <input type="number" id="topPushPrice" class="form-control form-control-lg" step="any" placeholder="Enter Price" style="height: 40px;">
                                                <button class="btn btn-success d-flex align-items-center" id="topSaveBtn" type="button" style="height: 40px;">
                                                    <i class="fas fa-save"></i>
                                                </button>
                                                <button class="btn btn-primary d-flex align-items-center" id="topPushBtn" type="button" style="height: 40px;">
                                                    <i class="fas fa-upload"></i>
                                                </button>
                                            </div>
                                        </div>
                                        <div>
                                            <span class="badge text-dark fs-4 text-bold me-2 bg-success">INV : <span id="ovl30InvLabel">0%</span></span>
                                            <span class="badge text-dark fs-4 text-bold">OV L30  : <span id="ovl30">0%</span></span>
                                            <span class="badge text-dark fs-4 text-bold ">Dil : <span id="dilPercentage"> </span> %</span>
                                            <span class="badge me-2 text-dark fs-4 text-bold">Avg Price: <span id="formattedAvgPrice">0%</span></span>
                                            <span class="badge text-dark fs-4 text-bold me-2">Profit  : <span id="formattedProfitPercentage">0%</span> %</span>
                                            <span class="badge text-dark fs-4  me-2">ROI : <span id="formattedRoiPercentage">0%</span> %</span>
                                            <span class="badge text-dark fs-4  me-2">Total Views : <span id="total_views">0</span></span>
                                            <span class="badge text-dark fs-4  me-2">Avg CVR : <span id="avgCvr">0%</span></span>
                                        </div>
                                    </div>
                                  <div class="view-controls d-flex justify-content-center align-items-center">
                                        <div class="image-preview-container">
                                            <img id="ovl30Img" src="" alt="Preview">
                                        </div>
                                    </div>

                                </div>
                            </div>
                            <div id="ovl30Content" class="p-3" style="color: #000000; width:100%; max-height: 70vh; overflow-y: auto;">
                                <!-- Marketplace data table will be loaded here -->
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>


    {{-- Marketplace Price Comparison Modal --}}
    <div class="modal fade" id="priceComparisonModal" tabindex="-1" aria-labelledby="priceComparisonModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl">
        <div class="modal-content">
            <div class="modal-header cursor-move">
                <h5 class="modal-title" id="priceComparisonModalLabel">
                    Marketplace Price Comparison for <span id="price-comparison-sku"></span>
                </h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <canvas id="priceComparisonChart"></canvas>
            </div>
        </div>
    </div>
</div>


@endsection
@section('script')
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script>
        document.body.style.zoom = "95%";

        function showPriceComparisonModal(row) {
            const data = row.getData();
            const sku = data.SKU;

            document.getElementById('price-comparison-sku').textContent = sku;

            const marketplaces = [
                { label: "Amazon", prefix: "amz" },
                { label: "eBay", prefix: "ebay" },
                { label: "Doba", prefix: "doba" },
                { label: "Macy", prefix: "macy" },
                { label: "Reverb", prefix: "reverb" },
                { label: "Temu", prefix: "temu" },
                { label: "Walmart", prefix: "walmart" },
                { label: "eBay2", prefix: "ebay2" },
                { label: "eBay3", prefix: "ebay3" },
                { label: "Shopify B2C", prefix: "shopifyb2c" },
                { label: "Shein", prefix: "shein" },
            ];

            const labels = [];
            const prices = [];

            marketplaces.forEach(m => {
                const price = data[`${m.prefix}_price`];
                if (price !== null && price !== undefined && price > 0) {
                    labels.push(m.label);
                    prices.push(price);
                }
            });

            // Chart.js setup and rendering here
             const modalEl = document.getElementById('priceComparisonModal');
             const chartCanvas = document.getElementById('priceComparisonChart');

            // Destroy previous chart instance if it exists
            if (window.priceChart instanceof Chart) {
                window.priceChart.destroy();
            }

            window.priceChart = new Chart(chartCanvas.getContext('2d'), {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Price',
                            data: prices,
                            borderColor: 'rgba(75, 192, 192, 1)',
                            backgroundColor: 'rgba(75, 192, 192, 0.2)',
                            fill: true,
                            tension: 0.1,
                            yAxisID: 'y'
                        },
                        {
                            label: 'L30',
                            data: labels.map(label => data[`${label.toLowerCase()}_l30`] || 0),
                            borderColor: 'rgba(255, 99, 132, 1)',
                            backgroundColor: 'rgba(255, 99, 132, 0.2)',
                            fill: true,
                            tension: 0.1,
                            yAxisID: 'y1'
                        },
                        // {
                        //     label: 'L60', 
                        //     data: labels.map(label => data[`${label.toLowerCase()}_l60`] || 0),
                        //     borderColor: '#2a0032',
                        //     backgroundColor: 'rgba(42, 0, 50, 0.2)',
                        //     fill: true,
                        //     tension: 0.1,
                        //     yAxisID: 'y1'
                        // },
                      
                    ]
                },
                options: {
                    responsive: true,
                    interaction: {
                        mode: 'index',
                        intersect: false,
                    },
                    scales: {
                        y: {
                            type: 'linear',
                            display: true,
                            position: 'left',
                            title: {
                                display: true,
                                text: 'Price ($)'
                            },
                            ticks: {
                                callback: function(value) {
                                    return '$' + value.toFixed(2);
                                }
                            }
                        },
                        y1: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            grid: {
                                drawOnChartArea: false,
                            },
                            title: {
                                display: true,
                                text: 'Sales Volume'
                            }
                        },
                        y2: {
                            type: 'linear',
                            display: true,
                            position: 'right',
                            grid: {
                                drawOnChartArea: false,
                            },
                            title: {
                                display: true,
                                text: 'Percentage (%)'
                            },
                            ticks: {
                                callback: function(value) {
                                    return value.toFixed(0) + '%';
                                }
                            }
                        }
                    },
                    plugins: {
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    let label = context.dataset.label || '';
                                    let value = context.parsed.y;
                                    
                                    if (label === 'Price') {
                                        return `${label}: $${value.toFixed(2)}`;
                                    } else if (label.includes('%')) {
                                        return `${label}: ${value.toFixed(1)}%`;
                                    } else {
                                        return `${label}: ${value}`;
                                    }
                                }
                            }
                        }
                    }
                }
            });
                

            const modal = new bootstrap.Modal(modalEl);
            modal.show();
        }

        // Helper function to calculate ROI
        function calculateROI(data) {
            const LP = parseFloat(data.LP) || 0;
            if (LP === 0) return 0;

            // Calculate total L30
            const totalL30 = (parseFloat(data.amz_l30) || 0) +
                (parseFloat(data.ebay_l30) || 0) +
                (parseFloat(data.shopifyb2c_l30) || 0) +
                (parseFloat(data.macy_l30) || 0) +
                (parseFloat(data.reverb_l30) || 0) +
                (parseFloat(data.doba_l30) || 0) +
                (parseFloat(data.temu_l30) || 0) +
                (parseFloat(data.ebay3_l30) || 0) +
                (parseFloat(data.ebay2_l30) || 0) +
                (parseFloat(data.walmart_l30) || 0) +
                (parseFloat(data.shein_l30) || 0);

            const SHIP = parseFloat(data.SHIP) || 0;
            const temuship = parseFloat(data.temu_ship) || 0;

            // Calculate profits
            const amzProfit = data.amz_price ? ((parseFloat(data.amz_price) * 0.70) - LP - SHIP) * (parseFloat(data
                .amz_l30) || 0) : 0;
            const ebayProfit = data.ebay_price ? ((parseFloat(data.ebay_price) * 0.72) - LP - SHIP) * (parseFloat(data
                .ebay_l30) || 0) : 0;
            const shopifyProfit = data.shopifyb2c_price ? ((parseFloat(data.shopifyb2c_price) * 0.75) - LP - SHIP) * (
                parseFloat(data.shopifyb2c_l30) || 0) : 0;
            const macyProfit = data.macy_price ? ((parseFloat(data.macy_price) * 0.76) - LP - SHIP) * (parseFloat(data
                .macy_l30) || 0) : 0;
            const reverbProfit = data.reverb_price ? ((parseFloat(data.reverb_price) * 0.84) - LP - SHIP) * (parseFloat(data
                .reverb_l30) || 0) : 0;
            const dobaProfit = data.doba_price ? ((parseFloat(data.doba_price) * 0.95) - LP - SHIP) * (parseFloat(data
                .doba_l30) || 0) : 0;
            const temuProfit = data.temu_price ? ((parseFloat(data.temu_price) * 0.87) - LP - temuship) * (parseFloat(data
                .temu_l30) || 0) : 0;
       
            const ebay3Profit = data.ebay3_price ? ((parseFloat(data.ebay3_price) * 0.71) - LP - SHIP) * (parseFloat(data
                .ebay3_l30) || 0) : 0;
            const ebay2Profit = data.ebay2_price ? ((parseFloat(data.ebay2_price) * 0.80) - LP - SHIP) * (parseFloat(data
                .ebay2_l30) || 0) : 0;
            const walmartProfit = data.walmart_price ? ((parseFloat(data.walmart_price) * 0.80) - LP - SHIP) * (parseFloat(
                data.walmart_l30) || 0) : 0;
            const sheinProfit = data.shein_price ? ((parseFloat(data.shein_price) * 0.80) - LP - SHIP) * (parseFloat(
                data.shein_l30) || 0) : 0;



            const totalProfit = amzProfit + ebayProfit + shopifyProfit + macyProfit + reverbProfit +
                dobaProfit + temuProfit  + ebay3Profit + ebay2Profit + walmartProfit + sheinProfit;

            return totalL30 > 0 ? (totalProfit / totalL30) / LP * 100 : 0;
        }

        // Helper function to calculate Average Profit
        function calculateAvgProfit(data) {
            const LP = parseFloat(data.LP) || 0;
            const SHIP = parseFloat(data.SHIP) || 0;
            const temuship = parseFloat(data.temu_ship) || 0;

            // Calculate profits and revenue for each marketplace
            const marketplaces = [
                { name: "amz", price: data.amz_price, l30: data.amz_l30, percent: 0.70 },
                { name: "ebay", price: data.ebay_price, l30: data.ebay_l30, percent: 0.72 },
                { name: "shopifyb2c", price: data.shopifyb2c_price, l30: data.shopifyb2c_l30, percent: 0.75 },
                { name: "macy", price: data.macy_price, l30: data.macy_l30, percent: 0.76 },
                { name: "reverb", price: data.reverb_price, l30: data.reverb_l30, percent: 0.84 },
                { name: "doba", price: data.doba_price, l30: data.doba_l30, percent: 0.95 },
                { name: "temu", price: data.temu_price, l30: data.temu_l30, percent: 0.87 }, // 👈 Temu special case
                { name: "ebay3", price: data.ebay3_price, l30: data.ebay3_l30, percent: 0.72 },
                { name: "ebay2", price: data.ebay2_price, l30: data.ebay2_l30, percent: 0.80 },
                { name: "walmart", price: data.walmart_price, l30: data.walmart_l30, percent: 0.89 },
                { name: "shein", price: data.shein_price, l30: data.shein_l30, percent: 0.89 }
            ];

            let totalProfit = 0;
            let totalRevenue = 0;

            marketplaces.forEach(mp => {
                const price = parseFloat(mp.price) || 0;
                const l30 = parseFloat(mp.l30) || 0;
                if (price && l30) {
                    const shippingCost = mp.name === "temu" ? temuship : SHIP;

                    totalProfit += ((price * mp.percent) - LP - shippingCost) * l30;
                    totalRevenue += price * l30;
                }
            });

            return totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;
        }


        // Image preview functions
        function showImagePreview(img) {
            const preview = document.getElementById('image-hover-preview');
            const previewImg = document.getElementById('preview-image');

            previewImg.src = img.src;
            preview.style.display = 'block';

            document.addEventListener('mousemove', moveImagePreview);
        }

        function hideImagePreview() {
            const preview = document.getElementById('image-hover-preview');
            preview.style.display = 'none';
            document.removeEventListener('mousemove', moveImagePreview);
        }

        function moveImagePreview(e) {
            const preview = document.getElementById('image-hover-preview');
            const rect = preview.getBoundingClientRect();

            // Calculate position, keeping the preview within viewport
            let x = e.pageX + 20;
            let y = e.pageY + 20;

            // Adjust if preview would go off screen
            if (x + rect.width > window.innerWidth) {
                x = e.pageX - rect.width - 20;
            }
            if (y + rect.height > window.innerHeight) {
                y = e.pageY - rect.height - 20;
            }

            preview.style.left = x + 'px';
            preview.style.top = y + 'px';
        }


        //global variables for play btn
     function renderGroup(parentKey) {
    if (!groupedSkuData[parentKey]) return;

    // Update current filter
    currentParentFilter = parentKey;
    setCombinedFilters();

    // Apply Tabulator filter for the selected group
    table.setFilter(function(data) {
        return data.Parent === parentKey;
    });
}

// Inventory Filter Listener
document.querySelectorAll("input[name='invFilter']").forEach(input => {
    input.addEventListener("change", function() {
        setCombinedFilters();
    });
});

// Dilution Filter Listener
document.querySelectorAll("input[name='dilFilter']").forEach(input => {
    input.addEventListener("change", function() {
        setCombinedFilters();
    });
});

// CVR Filter Listener
document.querySelectorAll("input[name='cvrFilter']").forEach(input => {
    input.addEventListener("change", function() {
        setCombinedFilters();
    });
});

// Margin Filter Listener
document.querySelectorAll("input[name='marginFilter']").forEach(input => {
    input.addEventListener("change", function() {
        setCombinedFilters();
    });
});

// View Filter Listener
document.querySelectorAll("input[name='viewFilter']").forEach(input => {
    input.addEventListener("change", function() {
        currentViewFilter = this.value;
        currentParentFilter = null; // Reset parent filter when view changes
        setCombinedFilters();
    });
});

const table = new Tabulator("#forecast-table", {
    ajaxURL: "/pricing-master-data-views",
    fixedHeader: true,
    width: "100%",
    height: "700px",
    pagination: true, 
    paginationSize: 50,
    initialFilter: [
        {field: "INV", type: ">", value: 0}
    ],
    initialSort: [{
        column: "inv_value",
        dir: "desc"
    }],

    rowFormatter: function(row) {
        const data = row.getData();
        const sku = data["SKU"] || '';

        if (sku.toUpperCase().includes("PARENT")) {
            row.getElement().classList.add("parent-row");
        }
    },

    dataLoaded: function(data) {
        updateHeaderTotals();
    },

    dataFiltering: function(filters, rows) {
        setTimeout(updateHeaderTotals, 0);
    },

    dataSorting: function(sorters, rows) {
        setTimeout(updateHeaderTotals, 0);
    },

    columns: [{
            title: "Image",
            field: "shopifyb2c_image",
            formatter: function(cell) {
                const value = cell.getValue();
                if (!value) return "";
                return `<img src="${value}" width="40" height="40" class="product-thumb" onmouseover="showImagePreview(this)" onmouseout="hideImagePreview()" style="cursor: pointer">`;
            },
            headerSort: false,
            width: 70,
            hozAlign: "center"
        },
       
        {
            title: "Parent",
            field: "Parent",
            headerFilter: "input",
            headerFilterPlaceholder: "Search Parent...",
            cssClass: "text-muted",
            tooltip: true,
            frozen: true
        },
        {
            title: "SKU",
            field: "SKU",
            headerFilter: "input",
            headerFilterPlaceholder: "Search SKU...",
            cssClass: "font-weight-bold",
            tooltip: true,
            frozen: true,
            formatter: function(cell) {
                let value = cell.getValue();
                return `
                    <span class="sku-text">${value}</span>
                    <i class="bi bi-clipboard ms-2 copy-icon" 
                    style="cursor:pointer;color:#007bff;" 
                    title="Copy SKU"></i>
                    <span class="copied-msg" style="display:none;color:green;font-size:12px;margin-left:5px;">Copied!</span>
                `;
            },
            cellClick: function(e, cell) {
                if (e.target.classList.contains("copy-icon")) {
                    let sku = cell.getValue();
                    navigator.clipboard.writeText(sku).then(() => {
                        let copiedMsg = cell.getElement().querySelector(".copied-msg");
                        copiedMsg.style.display = "inline";
                        setTimeout(() => {
                            copiedMsg.style.display = "none";
                        }, 500);
                    }).catch(err => {
                        console.error("Failed to copy: ", err);
                    });
                } else {
                    showPriceComparisonModal(cell.getRow());
                }
            },
        },
        {
            title: "INV",
            field: "INV",
            hozAlign: "right",
            formatter: function(cell) {
                const value = cell.getValue();
                return `<strong>${Math.round(value)}</strong>`;
            }
        },
        {
            title: "OVL30",
            field: "ovl30",
            hozAlign: "center",
            headerSort: false,
            formatter: function(cell) {
                const data = cell.getRow().getData();
                const l30 = data.shopifyb2c_l30 || 0;
                return `<button class="btn px-3" style="cursor:default !important;">
                    <i class="bi bi-eye me-1"></i>${l30}
                </button>`;
            },
            cellClick: function(e, cell) {
                showOVL30Modal(cell.getRow());
            }
        },
        {
            title: "DIL%",
            field: "Dil%",
            hozAlign: "right",
            formatter: function(cell) {
                const data = cell.getRow().getData();
                let value = 0;
                const l30 = parseFloat(data.ovl30) || 0;
                const inv = parseFloat(data.INV) || 0;
                if (inv !== 0) {
                    value = (l30 / inv) * 100;
                }
                const element = document.createElement("div");
                const rounded = parseFloat(Math.round(value));
                element.textContent = rounded + "%";
                if (rounded >= 0 && rounded <= 10) {
                    element.style.color = "red";
                } else if (rounded > 10 && rounded <= 15) {
                    element.style.backgroundColor = "yellow";
                    element.style.color = "black";
                    element.style.padding = "2px 4px";
                    element.style.borderRadius = "4px";
                } else if (rounded > 15 && rounded <= 20) {
                    element.style.color = "blue";
                } else if (rounded > 20 && rounded <= 40) {
                    element.style.color = "green";
                } else if (rounded > 40) {
                    element.style.color = "purple";
                }
                data.dilPercentage = rounded;
                return element;
            }
        },
        {
            title: "Inv Value<br><span id='invValueHeader' style='font-size:12px; color:#000; background:#fff; padding:1px 4px; border-radius:3px;'></span>",
            field: "inv_value",
            hozAlign: "center",
            headerSort: true,
            formatter: function(cell) {
                const value = parseFloat(cell.getValue()) || 0;
                return `<span class="text-success">$${Math.round(value)}</span>`;
            },
            sorter: function(a, b, aRow, bRow) {
                const valA = parseFloat(a) || 0;
                const valB = parseFloat(b) || 0;
                if (valA === valB) {
                    const skuA = aRow.getData().SKU || "";
                    const skuB = bRow.getData().SKU || "";
                    return skuA.localeCompare(skuB);
                }
                return valA - valB;
            }
        },
        {
            title: "COGS<br><span id='cogsHeader' style='font-size:12px; color:#000; background:#fff; padding:1px 4px; border-radius:3px;'></span>",
            field: "COGS",
            hozAlign: "center",
            headerSort: true,
            formatter: function(cell) {
                const value = parseFloat(cell.getValue()) || 0;
                return `<span class="text-success">$${Math.round(value)}</span>`;
            },
            sorter: function(a, b, aRow, bRow) {
                const valA = parseFloat(a) || 0;
                const valB = parseFloat(b) || 0;
                if (valA === valB) {
                    const skuA = aRow.getData().SKU || "";
                    const skuB = bRow.getData().SKU || "";
                    return skuA.localeCompare(skuB);
                }
                return valA - valB;
            }
        },
        {
            title: "Total Views",
            field: "total_views",
            hozAlign: "center",
            headerSort: false,
            formatter: function(cell) {
                const value = cell.getValue() || 0;
                return `<span class="text-danger">${Math.round(value)} </span>`;
            }
        },
        {
            title: "Total Req Views",
            field: "total_req_view",
            hozAlign: "center",
            headerSort: false,
            formatter: function(cell) {
                const value = cell.getValue() || 0;
                return `<span class="text-dark">${Math.round(value)} </span>`;
            }
        },
        {
            title: "Avg CVR",
            field: "avgCvr",
            hozAlign: "center",
            headerSort: false,
            formatter: function(cell) {
                let value = cell.getValue() || 0;
                if (typeof value === "string" && value.includes("%")) {
                    value = value.replace("%", "");
                }
                value = parseFloat(value);
                if (isNaN(value)) value = 0;

                let textColor;
                if (!isNaN(value)) {
                    if (value >= 0 && value < 3) {
                        textColor = '#dc3545';
                    } else if (value >= 3 && value < 6) {
                        textColor = '#ffc107';
                    } else if (value >= 6 && value < 9) {
                        textColor = '#0d6efd';
                    } else if (value >= 9 && value <= 13) {
                        textColor = '#198754';
                    } else {
                        textColor = '#6c757d';
                    }
                } else {
                    textColor = '#6c757d';
                }

                return `<span style="color: ${textColor}">${value.toFixed(2)}</span>`;
            }
        },
        {
            title: "AVG PRC",
            field: "avgPrice",
            hozAlign: "right",
            formatter: function(cell) {
                const data = cell.getRow().getData();

                const calculateAvgPrice = () => {
                    const marketplaces = [{
                            price: data.amz_price,
                            l30: data.amz_l30
                        },
                        {
                            price: data.ebay_price,
                            l30: data.ebay_l30
                        },
                        {
                            price: data.macy_price,
                            l30: data.macy_l30
                        },
                        {
                            price: data.reverb_price,
                            l30: data.reverb_l30
                        },
                        {
                            price: data.doba_price,
                            l30: data.doba_l30
                        },
                        {
                            price: data.temu_price,
                            l30: data.temu_l30
                        },
                        {
                            price: data.ebay3_price,
                            l30: data.ebay3_l30
                        },
                        {
                            price: data.ebay2_price,
                            l30: data.ebay2_l30
                        },
                        {
                            price: data.walmart_price,
                            l30: data.walmart_l30
                        },
                        {
                            price: data.shopifyb2c_price,
                            l30: data.shopifyb2c_l30
                        },
                        {
                            price: data.shein_price,
                            l30: data.shein_l30
                        },
                    ];

                    let totalWeightedPrice = 0;
                    let totalL30 = 0;

                    marketplaces.forEach(mp => {
                        const price = parseFloat(mp.price) || 0;
                        const l30 = parseFloat(mp.l30) || 0;
                        totalWeightedPrice += (price * l30);
                        totalL30 += l30;
                    });

                    return totalL30 > 0 ? (totalWeightedPrice / totalL30).toFixed(2) : '---';
                };

                const avgPrice = calculateAvgPrice();
                const avgPriceValue = parseFloat(avgPrice);

                let textColor, bgColor;
                if (!isNaN(avgPriceValue)) {
                    if (avgPriceValue < 10) {
                        textColor = '#dc3545';
                    } else if (avgPriceValue >= 10 && avgPriceValue < 15) {
                        textColor = '#fd7e14';
                    } else if (avgPriceValue >= 15 && avgPriceValue < 20) {
                        textColor = '#0d6efd';
                    } else if (avgPriceValue >= 20) {
                        textColor = '#198754';
                    }
                } else {
                    textColor = '#6c757d';
                }

                const element = document.createElement('div');
                element.innerHTML = avgPrice === '---' ? avgPrice : `$${avgPrice}`;
                element.style.color = textColor;
                element.style.fontWeight = '700';
                element.style.backgroundColor = bgColor;
                element.style.padding = '4px 8px';
                element.style.borderRadius = '4px';
                element.style.textAlign = 'center';

                data.formattedAvgPrice = avgPrice;
                return element;
            }
        },
        {
            title: "AVG PFT%<br><span id='avgPftHeader' style='font-size:12px; color:#fff; '></span>",
            field: "avgPftPercent",
            hozAlign: "right",
            headerSort: true,
            sortable: true,
            sorterParams: {
                alignEmptyValues: "bottom"
            },
            sorter: function(a, b) {
                let rowA = this.getRow(a);
                let rowB = this.getRow(b);
                return calculateAvgProfit(rowA.getData()) - calculateAvgProfit(rowB.getData());
            },
            headerSort: true,
            sorter: function(a, b) {
                const valA = a || 0;
                const valB = b || 0;
                return valA - valB;
            },
            formatter: function(cell) {
                const data = cell.getRow().getData();

                const LP = parseFloat(data.LP) || 0;
                const SHIP = parseFloat(data.SHIP) || 0;
                const ovl30 = parseFloat(data.shopifyb2c_l30) || 0;
                const temuship = parseFloat(data.temu_ship) || 0;
                const avgPrice = parseFloat(data.formattedAvgPrice) || 0;

                const amzPrice = parseFloat(data.amz_price) || 0;
                const ebayPrice = parseFloat(data.ebay_price) || 0;
                const shopifyPrice = parseFloat(data.shopifyb2c_price) || 0;
                const macyPrice = parseFloat(data.macy_price) || 0;
                const reverbPrice = parseFloat(data.reverb_price) || 0;
                const dobaPrice = parseFloat(data.doba_price) || 0;
                const temuPrice = parseFloat(data.temu_price) || 0;
                const ebay3Price = parseFloat(data.ebay3_price) || 0;
                const ebay2Price = parseFloat(data.ebay2_price) || 0;
                const walmartPrice = parseFloat(data.walmart_price) || 0;
                const sheinPrice = parseFloat(data.shein_price) || 0;

                const amzL30 = parseFloat(data.amz_l30) || 0;
                const ebayL30 = parseFloat(data.ebay_l30) || 0;
                const shopifyL30 = parseFloat(data.shopifyb2c_l30) || 0;
                const macyL30 = parseFloat(data.macy_l30) || 0;
                const reverbL30 = parseFloat(data.reverb_l30) || 0;
                const dobaL30 = parseFloat(data.doba_l30) || 0;
                const temuL30 = parseFloat(data.temu_l30) || 0;
                const ebay3L30 = parseFloat(data.ebay3_l30) || 0;
                const ebay2L30 = parseFloat(data.ebay2_l30) || 0;
                const walmartL30 = parseFloat(data.walmart_l30) || 0;
                const sheinL30 = parseFloat(data.shein_l30) || 0;

                const amzProfit = ((amzPrice * 0.70) - LP - SHIP);
                const ebayProfit = ((ebayPrice * 0.72) - LP - SHIP);
                const shopifyProfit = ((shopifyPrice * 0.75) - LP - SHIP);
                const macyProfit = ((macyPrice * 0.76) - LP - SHIP);
                const reverbProfit = ((reverbPrice * 0.84) - LP - SHIP);
                const dobaProfit = ((dobaPrice * 0.95) - LP - SHIP);
                const temuProfit = ((temuPrice * 0.87) - LP - temuship);
                const ebay3Profit = ((ebay3Price * 0.71) - LP - SHIP);
                const ebay2Profit = ((ebay2Price * 0.80) - LP - SHIP);
                const walmartProfit = ((walmartPrice * 0.80) - LP - SHIP);
                const sheinProfit = ((sheinPrice * 0.89) - LP - SHIP);

                const totalProfit = amzProfit * amzL30 + ebayProfit * ebayL30 + shopifyProfit * shopifyL30 + macyProfit * macyL30 +
                    reverbProfit * reverbL30 + dobaProfit * dobaL30 + temuProfit * temuL30 +
                    ebay3Profit * ebay3L30 + ebay2Profit * ebay2L30 + walmartProfit * walmartL30 +
                    sheinProfit * sheinL30;

                const totalRevenue =
                    (amzPrice * amzL30) +
                    (ebayPrice * ebayL30) +
                    (shopifyPrice * shopifyL30) +
                    (macyPrice * macyL30) +
                    (reverbPrice * reverbL30) +
                    (dobaPrice * dobaL30) +
                    (temuPrice * temuL30) +
                    (ebay3Price * ebay3L30) +
                    (ebay2Price * ebay2L30) +
                    (walmartPrice * walmartL30) +
                    (sheinPrice * sheinL30);

                let avgPftPercent = totalRevenue > 0 ? (totalProfit / totalRevenue) * 100 : 0;

                if (!isFinite(avgPftPercent) || isNaN(avgPftPercent)) {
                    avgPftPercent = 0;
                }

                avgPftPercent = Math.round(avgPftPercent);

                let bgColor, textColor;
                if (avgPftPercent < 11) {
                    textColor = '#ff0000';
                } else if (avgPftPercent >= 10 && avgPftPercent < 15) {
                    bgColor = 'yellow';
                    textColor = '#000000';
                } else if (avgPftPercent >= 15 && avgPftPercent < 20) {
                    textColor = '#0d6efd';
                } else if (avgPftPercent >= 21 && avgPftPercent < 50) {
                    textColor = '#198754';
                }
                else{
                    textColor = '#800080';
                }

                const element = document.createElement('div');
                element.textContent = avgPftPercent + '%';
                element.style.backgroundColor = bgColor;
                element.style.color = textColor;
                element.style.padding = '4px 8px';
                element.style.borderRadius = '4px';
                element.style.fontWeight = '600';
                element.style.textAlign = 'center';

                data.avgPftPercent = avgPftPercent;

                TotalAvgpft= (avgPrice * ovl30 ) * avgPftPercent / 100;
                TotalAvgSales = avgPrice * ovl30;
                TotalAvgpftForTop = (TotalAvgpft / TotalAvgSales) * 100;
                totalCogs = LP * ovl30;
                TotalAvgRoiPer = (TotalAvgpft / totalCogs) * 100;

                data.TotalAvgpft = Math.round(TotalAvgpft);
                data.TotalAvgSales = Math.round(TotalAvgSales);
                data.TotalAvgpftForTop = Math.round(TotalAvgpftForTop);
                data.totalCogs = Math.round(totalCogs);
                data.TotalAvgRoiPer = Math.round(TotalAvgRoiPer);

                return element;
            }
        },
        {
            title: "AVG ROI%<br><span id='avgRoiHeader' style='font-size:12px; color:#fff; '></span>",
            field: "avgRoi",
            hozAlign: "right",
            headerSort: true,
            sorter: function(a, b) {
                const valA = parseFloat(a) || 0;
                const valB = parseFloat(b) || 0;
                return valA - valB;
            },
            formatter: function(cell) {
                const data = cell.getRow().getData();
                const LP = parseFloat(data.LP) || 0;
                const ovl30 = parseFloat(data.shopifyb2c_l30) || 0;
                const temuship = parseFloat(data.temu_ship) || 0;
                const avgPrice = parseFloat(data.formattedAvgPrice) || 0;
                const SHIP = parseFloat(data.SHIP) || 0;
                
                if (LP === 0) return "N/A";

                const amzL30     = parseFloat(data.amz_l30) || 0;
                const ebayL30    = parseFloat(data.ebay_l30) || 0;
                const shopifyL30 = parseFloat(data.shopifyb2c_l30) || 0;
                const macyL30    = parseFloat(data.macy_l30) || 0;
                const reverbL30  = parseFloat(data.reverb_l30) || 0;
                const dobaL30    = parseFloat(data.doba_l30) || 0;
                const temuL30    = parseFloat(data.temu_l30) || 0;
                const ebay3L30   = parseFloat(data.ebay3_l30) || 0;
                const ebay2L30   = parseFloat(data.ebay2_l30) || 0;
                const walmartL30 = parseFloat(data.walmart_l30) || 0;
                const sheinL30   = parseFloat(data.shein_l30) || 0;

                const totalL30 = amzL30 + ebayL30 + shopifyL30 + macyL30 + reverbL30 + dobaL30 + temuL30  + ebay3L30 + ebay2L30 + walmartL30 + sheinL30;

                const amzProfit     = data.amz_price        ? ((parseFloat(data.amz_price) * 0.70) - LP - SHIP) * amzL30 : 0;
                const ebayProfit    = data.ebay_price       ? ((parseFloat(data.ebay_price) * 0.72) - LP - SHIP) * ebayL30 : 0;
                const shopifyProfit = data.shopifyb2c_price ? ((parseFloat(data.shopifyb2c_price) * 0.75) - LP - SHIP) * shopifyL30 : 0;
                const macyProfit    = data.macy_price       ? ((parseFloat(data.macy_price) * 0.76) - LP - SHIP) * macyL30 : 0;
                const reverbProfit  = data.reverb_price     ? ((parseFloat(data.reverb_price) * 0.84) - LP - SHIP) * reverbL30 : 0;
                const dobaProfit    = data.doba_price       ? ((parseFloat(data.doba_price) * 0.95) - LP - SHIP) * dobaL30 : 0;
                const temuProfit    = data.temu_price       ? ((parseFloat(data.temu_price) * 0.87) - LP - temuship) * temuL30 : 0;
                const ebay3Profit   = data.ebay3_price      ? ((parseFloat(data.ebay3_price) * 0.71) - LP - SHIP) * ebay3L30 : 0;
                const ebay2Profit   = data.ebay2_price      ? ((parseFloat(data.ebay2_price) * 0.80) - LP - SHIP) * ebay2L30 : 0;
                const walmartProfit = data.walmart_price    ? ((parseFloat(data.walmart_price) * 0.80) - LP - SHIP) * walmartL30 : 0;
                const sheinProfit   = data.shein_price      ? ((parseFloat(data.shein_price) * 0.89) - LP - SHIP) * sheinL30 : 0;

                const totalProfit = amzProfit + ebayProfit + shopifyProfit + macyProfit +
                                    reverbProfit + dobaProfit + temuProfit  +
                                    ebay3Profit + ebay2Profit + walmartProfit + sheinProfit;

                const roi = totalL30 > 0 ? (totalProfit / totalL30) / LP * 100 : 0;

                data.TotalAvgRoiPer = Math.round(roi);

                let bgColor, textColor;
                if (roi < 11) {
                    textColor = '#ff0000';
                } else if (roi >= 10 && roi < 15) {
                    bgColor = 'yellow';
                    textColor = '#000000';
                } else if (roi >= 15 && roi < 20) {
                    textColor = '#0d6efd';
                } else if (roi >= 21 && roi < 50) {
                    textColor = '#198754';
                } else {
                    textColor = '#800080';
                }

                const element = document.createElement('div');
                element.textContent = Math.round(roi) + '%';
                element.style.backgroundColor = bgColor;
                element.style.color = textColor;
                element.style.padding = '4px 8px';
                element.style.borderRadius = '4px';
                element.style.fontWeight = '600';
                element.style.textAlign = 'center';

                data.avgRoi = Math.round(roi);
                return element;
            }
        },
        {
            title: "MSRP",
            field: "MSRP",
            hozAlign: "right",
            formatter: "money",
            formatterParams: {
                precision: 2
            }
        },
        {
            title: "MAP",
            field: "MAP",
            hozAlign: "right",
            formatter: "money",
            formatterParams: {
                precision: 2
            }
        },
        {
            title: "LP",
            field: "LP",
            hozAlign: "right",
            formatter: "money",
            formatterParams: {
                precision: 2
            }
        },
        {
            title: "SHIP",
            field: "SHIP",
            hozAlign: "right",
            formatter: "money",
            formatterParams: {
                precision: 2
            }
        },
    ],

    ajaxResponse: function(url, params, response) {
        groupedSkuData = {};
        
        response.data = response.data.map((item, index) => {
            const sku = item.SKU || "";
            const isParent = item.is_parent || sku.toUpperCase().includes("PARENT");
            item.ovl30 = parseFloat(item.shopifyb2c_l30) || 0;
            
            const inv = parseFloat(item.INV) || 0;
            item["Dil%"] = inv !== 0 ? ((item.ovl30 / inv) * 100).toFixed(2) : "0";
            
            if (!item.inv_value) {
                const inv = item.INV || 0;
                const shopifyPrice = parseFloat(item.shopifyb2c_price) || 0;
                item.inv_value = (inv * shopifyPrice).toFixed(2);
            }
            
            if (!item.COGS) {
                const lp = parseFloat(item.LP) || 0;
                const inv = item.INV || 0;
                item.COGS = (lp * inv).toFixed(2);
            }
            
            return {
                ...item,
                calculatedRoi: calculateROI(item),
                calculatedProfit: calculateAvgProfit(item),
                sl_no: index + 1,
                is_parent: isParent ? 1 : 0,
                isParent: isParent,
                raw_data: item || {}
            };
        });

        let grouped = {};
        response.data.forEach(item => {
            const parentKey = item.Parent || "";
            if (!grouped[parentKey]) grouped[parentKey] = [];
            grouped[parentKey].push(item);
            if (!groupedSkuData[parentKey]) {
                groupedSkuData[parentKey] = [];
            }
            groupedSkuData[parentKey].push(item);
        });

        Object.keys(grouped).forEach(parentKey => {
            const rows = grouped[parentKey];
            const children = rows.filter(item => !item.is_parent);
            const parent = rows.find(item => item.is_parent);
            if (!parent || children.length === 0) return;

            const additiveFields = ['INV', 'total_views', 'total_req_view', 'inv_value', 'COGS', 'ovl30'];
            additiveFields.forEach(field => {
                parent[field] = children.reduce((sum, c) => sum + (parseFloat(c[field]) || 0), 0).toFixed(2);
            });

            const parentL30 = parseFloat(parent.ovl30) || 0;
            const parentInv = parseFloat(parent.INV) || 0;
            parent["Dil%"] = parentInv !== 0 ? ((parentL30 / parentInv) * 100).toFixed(2) : "0";

            const rateFields = ['avgCvr', 'MSRP', 'MAP', 'LP', 'SHIP', 'temu_ship', 'avgPftPercent'];
            rateFields.forEach(field => {
                const values = children.map(c => parseFloat(c[field]) || 0);
                const valid = values.filter(v => !isNaN(v) && v !== 0);
                parent[field] = valid.length > 0 ? (valid.reduce((sum, v) => sum + v, 0) / valid.length).toFixed(2) :
                    (values.length > 0 ? (values.reduce((sum, v) => sum + v, 0) / values.length).toFixed(2) : 0);
            });

            const mps = ['amz', 'ebay', 'macy', 'reverb', 'doba', 'temu', 'ebay3', 'ebay2', 'walmart', 'shein', 'shopifyb2c'];
            mps.forEach(mp => {
                const l30Field = (mp === 'shopifyb2c' ? 'shopifyb2c_l30' : `${mp}_l30`);
                const priceField = (mp === 'shopifyb2c' ? 'shopifyb2c_price' : `${mp}_price`);
                parent[l30Field] = children.reduce((sum, c) => sum + (parseFloat(c[l30Field]) || 0), 0);
                let totalWeighted = 0;
                let totalWeight = 0;
                children.forEach(c => {
                    const price = parseFloat(c[priceField]) || 0;
                    const weight = parseFloat(c[l30Field]) || 0;
                    totalWeighted += price * weight;
                    totalWeight += weight;
                });
                parent[priceField] = totalWeight > 0 ? (totalWeighted / totalWeight).toFixed(2) :
                    (children.reduce((sum, c) => sum + (parseFloat(c[priceField]) || 0), 0) / children.length).toFixed(2);
            });

            const inv = parseFloat(parent.INV) || 0;
            const shopifyPrice = parseFloat(parent.shopifyb2c_price) || 0;
            parent.inv_value = (inv * shopifyPrice).toFixed(2);

            const lp = parseFloat(parent.LP) || 0;
            parent.COGS = (lp * inv).toFixed(2);

            const marketplaces = [
                { price: parent.amz_price, l30: parent.amz_l30, factor: 0.70 },
                { price: parent.ebay_price, l30: parent.ebay_l30, factor: 0.72 },
                { price: parent.shopifyb2c_price, l30: parent.shopifyb2c_l30, factor: 0.75 },
                { price: parent.macy_price, l30: parent.macy_l30, factor: 0.76 },
                { price: parent.reverb_price, l30: parent.reverb_l30, factor: 0.84 },
                { price: parent.doba_price, l30: parent.doba_l30, factor: 0.95 },
                { price: parent.temu_price, l30: parent.temu_l30, factor: 0.87, ship: parent.temu_ship },
                { price: parent.ebay3_price, l30: parent.ebay3_l30, factor: 0.71 },
                { price: parent.ebay2_price, l30: parent.ebay2_l30, factor: 0.80 },
                { price: parent.walmart_price, l30: parent.walmart_l30, factor: 0.80 },
                { price: parent.shein_price, l30: parent.shein_l30, factor: 0.89 }
            ];
            let totalProfit = 0;
            let totalRevenue = 0;
            marketplaces.forEach(mp => {
                const price = parseFloat(mp.price) || 0;
                const l30 = parseFloat(mp.l30) || 0;
                const ship = parseFloat(mp.ship || parent.SHIP) || 0;
                const profit = ((price * mp.factor) - lp - ship) * l30;
                totalProfit += profit;
                totalRevenue += price * l30;
            });
            parent.avgPftPercent = totalRevenue > 0 ? ((totalProfit / totalRevenue) * 100).toFixed(2) : 0;

            if (!parent.shopifyb2c_image && children[0]) {
                parent.shopifyb2c_image = children[0].shopifyb2c_image;
            }

            parent.ovl30 = parseFloat(parent.shopifyb2c_l30) || 0;
        });

        let finalData = [];
        Object.values(grouped).forEach(rows => {
            rows.sort((a, b) => {
                if (a.is_parent !== b.is_parent) {
                    return a.is_parent - b.is_parent;
                }
                return (a.SKU || "").localeCompare(b.SKU || "");
            });
            finalData = finalData.concat(rows);
        });

        setTimeout(() => {
            setCombinedFilters();
            updateHeaderTotals();
        }, 0);
        
        console.log("Processed Response:", finalData);
        return finalData;
    },

    });

    // Function to update header totals (exclude parent rows)
    function updateHeaderTotals() {
        const tableData = table.getData();

        let totalInvValue = 0;
        let totalCogs = 0;

        tableData.forEach(row => {
            // Only add SKU rows (not parent)
            if (!row.is_parent) {
                totalInvValue += parseFloat(row.inv_value) || 0;
                totalCogs += parseFloat(row.COGS) || 0;
            }
        });

        const formattedInvValue = Math.round(totalInvValue).toLocaleString();
        const formattedCogs = Math.round(totalCogs).toLocaleString();

        const invValueHeader = document.getElementById('invValueHeader');
        const cogsHeader = document.getElementById('cogsHeader');

        if (invValueHeader) {
            invValueHeader.textContent = `${formattedInvValue}`;
        }

        if (cogsHeader) {
            cogsHeader.textContent = `${formattedCogs}`;
        }
    }

// Initialize totals when page loads
setTimeout(updateHeaderTotals, 1000);

//   On to Percentaeg color
    table.on("dataProcessed", function(){
        let data = table.getData();

        // --- AVG PFT% calculation ---
        let totalPft = 0, countPft = 0;
        data.forEach(row => {
            if (row.TotalAvgpftForTop) {
                totalPft += row.TotalAvgpftForTop;
                countPft++;
            }
        });
        let avgPft = countPft > 0 ? (totalPft / countPft) : 0;

        let pftHeader = document.getElementById("avgPftHeader");
        pftHeader.innerText = avgPft.toFixed(1) + "%";


        // Style for AVG PFT%
        let bgColorPft, textColorPft;
        if (avgPft < 11) {
            bgColorPft = "#ffe5e5"; textColorPft = "#ff0000";
        } else if (avgPft >= 10 && avgPft < 15) {
            bgColorPft = "yellow"; textColorPft = "#000000";
        } else if (avgPft >= 15 && avgPft < 20) {
            bgColorPft = "#e6f0ff"; textColorPft = "#0d6efd";
        } else if (avgPft >= 21 && avgPft < 50) {
            bgColorPft = "#e6ffe6"; textColorPft = "#198754";
        } else {
            bgColorPft = "#f5e6f5"; textColorPft = "#800080";
        }
        pftHeader.style.color = textColorPft;
        pftHeader.style.backgroundColor = bgColorPft;
        pftHeader.style.padding = "3px 6px";
        pftHeader.style.borderRadius = "4px";
        pftHeader.style.fontWeight = "600";

        // --- AVG ROI% calculation ---
        let totalRoi = 0, countRoi = 0;
        data.forEach(row => {
            if (row.TotalAvgRoiPer) {
                totalRoi += row.TotalAvgRoiPer;
                countRoi++;
            }
        });
        let avgRoi = countRoi > 0 ? (totalRoi / countRoi) : 0;

        let roiHeader = document.getElementById("avgRoiHeader");
        roiHeader.innerText = Math.round(avgRoi) + "%";


        // Style for AVG ROI%
        let bgColorRoi, textColorRoi;
        if (avgRoi < 11) {
            bgColorRoi = "#ffe5e5"; textColorRoi = "#ff0000";
        } else if (avgRoi >= 10 && avgRoi < 15) {
            bgColorRoi = "yellow"; textColorRoi = "#000000";
        } else if (avgRoi >= 15 && avgRoi < 20) {
            bgColorRoi = "#e6f0ff"; textColorRoi = "#0d6efd";
        } else if (avgRoi >= 21 && avgRoi < 50) {
            bgColorRoi = "#e6ffe6"; textColorRoi = "#198754";
        } else {
            bgColorRoi = "#f5e6f5"; textColorRoi = "#800080";
        }
        roiHeader.style.color = textColorRoi;
        roiHeader.style.backgroundColor = bgColorRoi;
        roiHeader.style.padding = "3px 6px";
        roiHeader.style.borderRadius = "4px";
        roiHeader.style.fontWeight = "600";
    });




let currentViewFilter = null;
let currentParentFilter = null;

// Function to apply combined filters
function setCombinedFilters() {
    const filters = [];

    // Apply inventory filter
    const invFilter = document.querySelector("input[name='invFilter']:checked")?.value;
    if (invFilter === "zero") {
        filters.push({ field: "INV", type: "=", value: 0 }); // Show only rows with INV = 0
    } else if (invFilter === "other") {
        filters.push({ field: "INV", type: ">", value: 0 }); // Show only rows with INV > 0
    } else {
        // Default: Exclude rows with INV = 0
        filters.push({ field: "INV", type: ">", value: 0 });
    }

    // Apply dilution filter if active
    const dilFilter = document.querySelector("input[name='dilFilter']:checked")?.value;
    if (dilFilter === "10") {
        filters.push([
            { field: "Dil%", type: "<=", value: 10 },
            function(row) {
                // Exclude -1 Dil%
                return parseFloat(row.getData()["Dil%"]) !== -1;
            }
        ]);
    } else if (dilFilter === "50") {
        filters.push({ field: "Dil%", type: ">", value: 50 }); 
    }

    // Apply CVR filter if active
    const cvrFilter = document.querySelector("input[name='cvrFilter']:checked")?.value;
    if (cvrFilter === "low") {
        filters.push({ field: "avgCvr", type: "<", value: 5 });
    }

    // Apply margin filter if active
    const marginFilter = document.querySelector("input[name='marginFilter']:checked")?.value;
    if (marginFilter === "high") {
        filters.push({ field: "avgPftPercent", type: ">", value: 20 });
    }

    // Apply parent or view filter
    if (currentViewFilter === "parent") {
        filters.push({ field: "is_parent", type: "=", value: 1 });
        table.getColumn("Parent").show();
        table.setSort([{ column: "inv_value", dir: "desc" }]);
    } else if (currentViewFilter === "sku") {
        filters.push({ field: "is_parent", type: "=", value: 0 });
        table.getColumn("SKU").show();
        table.setSort([{ column: "SKU", dir: "asc" }]);
    } else if (currentViewFilter === "both") {
        // Show both Parent and SKU columns, no extra filter
        table.getColumn("Parent").show();
        table.getColumn("SKU").show();
        table.clearFilter();
        table.setSort([{ column: "inv_value", dir: "desc" }]);
    }

    // Apply all filters
    if (filters.length > 0) {
        table.setFilter(filters);
    } else {
        table.clearFilter();
    }
}

    // Function to add trend indicators
    function addTrendIndicators(row) {
        const data = row.getData();
        const cells = row.getCells();

        cells.forEach(cell => {
            const field = cell.getColumn().getField();
            if (field.includes('l30') || field.includes('l60')) {
                const value = cell.getValue();
                const prevValue = data[field.replace('l30', 'l60')] || 0;

                if (value > prevValue) {
                    cell.getElement().classList.add('trend-up');
                } else if (value < prevValue) {
                    cell.getElement().classList.add('trend-down');
                }
            }
        });
    }

        // On TOp Caalculation

    
// Variable to prevent infinite scroll loop
let isSyncingScroll = false;

document.addEventListener('DOMContentLoaded', function() {
    // Get the scrollable containers
    const mainTableHolder = document.querySelector('#forecast-table .tabulator-tableholder');
    const modalTableHolder = document.querySelector('#ovl30Modal .table-responsive');

    // Function to sync scroll from main table to modal table
    function syncMainToModal() {
        if (!isSyncingScroll && mainTableHolder && modalTableHolder) {
            isSyncingScroll = true;
            modalTableHolder.scrollTop = mainTableHolder.scrollTop;
            isSyncingScroll = false;
        }
    }

    // Function to sync scroll from modal table to main table
    function syncModalToMain() {
        if (!isSyncingScroll && mainTableHolder && modalTableHolder) {
            isSyncingScroll = true;
            mainTableHolder.scrollTop = modalTableHolder.scrollTop;
            isSyncingScroll = false;
        }
    }

    // Add scroll event listeners
    if (mainTableHolder) {
        mainTableHolder.addEventListener('scroll', syncMainToModal);
    }
    if (modalTableHolder) {
        modalTableHolder.addEventListener('scroll', syncModalToMain);
    }

    // Clean up event listeners when modal is closed
    const ovl30Modal = document.getElementById('ovl30Modal');
    ovl30Modal.addEventListener('hidden.bs.modal', function() {
        if (mainTableHolder) {
            mainTableHolder.removeEventListener('scroll', syncMainToModal);
        }
        if (modalTableHolder) {
            modalTableHolder.removeEventListener('scroll', syncModalToMain);
        }
    });

    // Re-attach event listeners when modal is shown
    ovl30Modal.addEventListener('shown.bs.modal', function() {
        const newMainTableHolder = document.querySelector('#forecast-table .tabulator-tableholder');
        const newModalTableHolder = document.querySelector('#ovl30Modal .table-responsive');
        if (newMainTableHolder) {
            newMainTableHolder.addEventListener('scroll', syncMainToModal);
        }
        if (newModalTableHolder) {
            newModalTableHolder.addEventListener('scroll', syncModalToMain);
        }
    });
});

    </script>

    <script>
        // Helper: percent formatting
        function fmtPct(v) {
            if (v === null || v === undefined || v === "") return "-";
            const num = parseFloat(v);
            if (isNaN(num)) return "-";

            
            return Math.round(num * 100) + "%";
        }


        // Helper: money formatting
        function fmtMoney(v) {
            if (v === null || v === undefined || v === "") return "-";
            const num = parseFloat(v);
            if (isNaN(num)) return "-";
            return "$" + num.toFixed(2);
        }

        // Marketplace table generator
        function buildOVL30Table(data) {
          const rows = [
                { label: "Amazon", prefix: "amz", logo: "{{ asset('uploads/amazon.png') }}" },
                { label: "eBay", prefix: "ebay", logo:  "{{ asset('uploads/1.png') }}" },
                { label: "Doba", prefix: "doba", logo: "{{ asset('uploads/doba.png') }}" },
                { label: "Macy", prefix: "macy", logo: "{{ asset('uploads/macy.png') }}" },
                { label: "Reverb", prefix: "reverb", logo: "{{ asset('uploads/reverb.png') }}" },
                { label: "Temu", prefix: "temu", logo: "{{ asset('uploads/temu.jpeg') }}" },
                { label: "Walmart", prefix: "walmart", logo: "{{ asset('uploads/walmart.png') }}" },
                { label: "eBay2", prefix: "ebay2", logo: "{{ asset('uploads/2.png') }}" },
                { label: "eBay3", prefix: "ebay3", logo: "{{ asset('uploads/3.png') }}" },
                { label: "Shopify B2C", prefix: "shopifyb2c", logo: "{{ asset('uploads/shopify.png') }}" },
                { label: "Shein", prefix: "shein", logo: "{{ asset('uploads/Shein.jpg') }}" }
            ];



            let html = `
            <div class="table-responsive">
            <div class="mb-2 text-muted small">
                <i class="bi bi-info-circle"></i> Default sorting: L30 (Highest to Lowest)
            </div>
            <div class="table-responsive" style="max-height: 600px; overflow-y: auto; position: relative;">
            <table class="table table-sm table-bordered align-middle sortable-table">
                <thead class="table-light position-sticky" style="top: 0; z-index: 1000;">
                <tr>
                    <th data-sort="string">Channel <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number" class="default-sort">L30 <i class="bi bi-arrow-down"></i></th>
                    <th data-sort="number">PRC <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">PFT % <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">ROI % <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">Views L30 <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">CVR <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">Req Views <i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">LMP <i class="bi bi-arrow-down-up"></i></th>
                    <th>S Price</th>
                    <th data-sort="number">S PFT<i class="bi bi-arrow-down-up"></i></th>
                    <th data-sort="number">S ROI<i class="bi bi-arrow-down-up"></i></th>
                </tr>
                </thead>
                <tbody>
            `;
            

            rows.forEach(r => {
                const price = data[`${r.prefix}_price`];
                const l30 = r.prefix === 'shopifyb2c' ? data['shopify_l30'] : data[`${r.prefix}_l30`];
                const l60 = data[`${r.prefix}_l60`];
                const pft = data[`${r.prefix}_pft`];
                const roi = data[`${r.prefix}_roi`];
                const cvr = data[`${r.prefix}_cvr`];
                const reqCvr = data[`${r.prefix}_req_view`];

                const hasAny = price != null || l30 != null || l60 != null || pft != null || roi != null;
                if (!hasAny) return;

               const getColor = (value) => {
                    // Convert to number and handle percentage values
                    const val = typeof value === 'string' ? parseFloat(value.replace('%', '')) : Number(value);
                    // Make sure value is a finite number
                    if (!isFinite(val) || isNaN(val)) return '#000000'; // default black
                    
                    // Handle percentage ranges
                    if (val >= 0 && val <= 10) return '#ff0000';        // red
                    if (val > 10 && val <= 14) return '#fd7e14';       // orange
                    if (val > 14 && val <= 19) return '#0d6efd';       // blue
                    if (val > 19 && val <= 40) return '#198754';       // green
                    if (val > 40) return '#800080';                    // purple
                    
                    return '#000000';                                  // default black
                };

                const pftClass = pft > 20 ? 'positive' : pft < 10 ? 'negative' : 'neutral';
                const roiClass = roi > 30 ? 'positive' : roi < 15 ? 'negative' : 'neutral';
                

                html += `
                 <tr>
                    <td>
                    <div class="d-flex flex-column align-items-center text-center">
                        <div class="position-relative">
                            <img src="${r.logo}" alt="${r.label}" 
                                class="channel-logo mb-1" 
                                style="width:30px; height:30px; object-fit:contain; cursor: pointer;"
                                onmouseenter="showTooltip(this)"
                                onmouseleave="hideTooltip(this)">
                            
                            <!-- Tooltip for links -->
                            <div class="position-absolute bg-dark text-white p-2 rounded shadow-sm link-tooltip" 
                                style="bottom: 0px; left: 70px; 
                                       opacity: 0; visibility: hidden; transition: all 0.3s; 
                                       white-space: nowrap; z-index: 1000; font-size: 11px;"
                                onmouseenter="showTooltip(this.previousElementSibling)"
                                onmouseleave="hideTooltip(this.previousElementSibling)">
                                ${r.prefix === 'amz' ? `
                                    ${data.amz_seller_link ? `<div><strong>SL:</strong> <a href="${data.amz_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.amz_buyer_link ? `<div><strong>BL:</strong> <a href="${data.amz_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'ebay' ? `
                                    ${data.ebay_seller_link ? `<div><strong>SL:</strong> <a href="${data.ebay_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.ebay_buyer_link ? `<div><strong>BL:</strong> <a href="${data.ebay_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'ebay2' ? `
                                    ${data.ebay2_seller_link ? `<div><strong>SL:</strong> <a href="${data.ebay2_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.ebay2_buyer_link ? `<div><strong>BL:</strong> <a href="${data.ebay2_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'ebay3' ? `
                                    ${data.ebay3_seller_link ? `<div><strong>SL:</strong> <a href="${data.ebay3_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.ebay3_buyer_link ? `<div><strong>BL:</strong> <a href="${data.ebay3_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'macy' ? `
                                    ${data.macy_seller_link ? `<div><strong>SL:</strong> <a href="${data.macy_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.macy_buyer_link ? `<div><strong>BL:</strong> <a href="${data.macy_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'reverb' ? `
                                    ${data.reverb_seller_link ? `<div><strong>SL:</strong> <a href="${data.reverb_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.reverb_buyer_link ? `<div><strong>BL:</strong> <a href="${data.reverb_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'walmart' ? `
                                    ${data.walmart_seller_link ? `<div><strong>SL:</strong> <a href="${data.walmart_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.walmart_buyer_link ? `<div><strong>BL:</strong> <a href="${data.walmart_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'doba' ? `
                                    ${data.doba_seller_link ? `<div><strong>SL:</strong> <a href="${data.doba_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.doba_buyer_link ? `<div><strong>BL:</strong> <a href="${data.doba_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'temu' ? `
                                    ${data.temu_seller_link ? `<div><strong>SL:</strong> <a href="${data.temu_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.temu_buyer_link ? `<div><strong>BL:</strong> <a href="${data.temu_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : r.prefix === 'shopifyb2c' ? `
                                    ${data.shopifyb2c_seller_link ? `<div><strong>SL:</strong> <a href="${data.shopifyb2c_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.shopifyb2c_buyer_link ? `<div><strong>BL:</strong> <a href="${data.shopifyb2c_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                 ` : r.prefix === 'shein' ? `
                                    ${data.shein_seller_link ? `<div><strong>SL:</strong> <a href="${data.shein_seller_link}" target="_blank" class="text-info">Seller Link</a></div>` : ''}
                                    ${data.shein_buyer_link ? `<div><strong>BL:</strong> <a href="${data.shein_buyer_link}" target="_blank" class="text-success">Buyer Link</a></div>` : ''}
                                ` : ''}

                            </div>
                        </div>
                        <span class="small fw-bold">${r.label}</span>
                    </div>
                    </td>


                    <td>
                        <div class="value-indicator">
                            ${l30 ?? "-"}
                        </div>
                    </td>
                
                    <td>
                        <div class="value-indicator">
                            ${fmtMoney(price)}
                        </div>
                    </td>
                    <td>
                        <div class="value-indicator ${pftClass}" style="color: ${getColor(pft * 100)};">
                            ${fmtPct(pft)}
                        </div>
                    </td>
                    <td>
                        <div class="value-indicator ${roiClass}" style="color: ${getColor(roi * 100)};">
                            ${fmtPct(roi)}
                        </div>
                    </td>
                
                
                    <td>
                        <div class="value-indicator">
                            ${r.prefix === 'amz' ? (data.sessions_l30 ?? "-") 
                                : r.prefix === 'ebay' ? (data.ebay_views ?? "-") 
                                : r.prefix === 'ebay2' ? (data.ebay2_views ?? "-") 
                                : r.prefix === 'ebay3' ? (data.ebay3_views ?? "-") 
                                : r.prefix === 'shein' ? (data.views_clicks ?? "-")
                                : "-" }
                        </div>
                    </td>
                    <td>
                        <div class="value-indicator">
                            ${(() => {
                                if (r.prefix === 'amz' && cvr) {
                                    return `<span style="color: ${cvr.color}">${Math.round(cvr.value)}%</span>`;
                                } else if (r.prefix === 'ebay' && cvr) {
                                    return `<span style="color: ${cvr.color}">${Math.round(cvr.value)}%</span>`;
                                } else if (r.prefix === 'ebay3' && cvr) {
                                    return `<span style="color: ${cvr.color}">${Math.round(cvr.value)}%</span>`;
                                }
                                else if (r.prefix === 'shein' && cvr) {
                                    return `<span style="color: ${cvr.color}">${Math.round(cvr.value)}%</span>`;
                                }

                                return "N/A";
                            })()} 
                        </div>
                    </td>
                     <td>
                        <div class="value-indicator">
                        ${r.prefix === 'amz' ? Math.round(data.amz_req_view) ?? "-" : 
                            r.prefix === 'ebay' ? Math.round(data.ebay_req_view) ?? "-" :
                            r.prefix === 'ebay2' ? Math.round(data.ebay2_req_view) ?? "-" :
                            r.prefix === 'ebay3' ? Math.round(data.ebay3_req_view) ?? "-" :
                            r.prefix === 'shein' ? Math.round(data.shein_req_view) ?? "-" : "-"}
                        </div>
                    </td>


                    <td>
                        <div class="value-indicator">
                            ${r.prefix === 'amz' ? fmtMoney(data.price_lmpa) 
                                : r.prefix === 'ebay' ? fmtMoney(data.ebay_price_lmpa) 
                                : r.prefix === 'shein' ? fmtMoney(data.lmp) 
                                : '-'}
                        </div>
                    </td>


                

              <td>
                    <div class="d-flex align-items-center gap-2">
                        <input type="text" 
                            class="form-control form-control-sm s-price" 
                            value="${
                                r.prefix === 'amz' ? (data.amz_sprice || '') 
                                : r.prefix === 'ebay' ? (data.ebay_sprice || '') 
                                : r.prefix === 'shopifyb2c' ? (data.shopifyb2c_sprice || '') 
                                : r.prefix === 'ebay2' ? (data.ebay2_sprice || '') 
                                : r.prefix === 'ebay3' ? (data.ebay3_sprice || '')
                                : r.prefix === 'doba' ? (data.doba_sprice || '')
                                : r.prefix === 'temu' ? (data.temu_sprice || '')
                                : r.prefix === 'macy' ? (data.macy_sprice || '')
                                : r.prefix === 'reverb' ? (data.reverb_sprice || '')
                                : r.prefix === 'walmart' ? (data.walmart_sprice || '')
                                : r.prefix === 'shein' ? (data.shein_sprice || '')
                            
                                : ''
                            }"
                            style="width: 85px;" 
                            step="any"
                            data-sku="${data.SKU}" 
                            data-lp="${data.LP}" 
                            data-ship="${
                                r.prefix === 'temu' ? (data.temu_ship || '') : (data.SHIP || '')
                            }"
                            data-type="${r.prefix}">

                        <!-- Push to Marketplace -->
                        <button class="btn btn-success btn-sm d-flex align-items-center pushPriceBtn" 
                            type="button"
                            data-sku="${data.SKU}" 
                            data-type="${r.prefix}">
                            <i class="bi bi-cloud-arrow-up"></i>
                        </button>
                    </div>
                </td>


                    <td class="spft-field">
                        ${(() => {
                            let value, textColor, bgColor;
                            
                            if (r.prefix === 'amz' && data.amz_spft) {
                                value = Math.round(data.amz_spfst);
                            } else if (r.prefix === 'ebay' && data.ebay_spft) {
                                value = Math.round(data.ebay_spft);
                            } else if (r.prefix === 'shopifyb2c' && data.shopifyb2c_spft) {
                                value = Math.round(data.shopifyb2c_spft);
                            } else if (r.prefix === 'ebay2' && data.ebay2_spft) {
                                value = Math.round(data.ebay2_spft);
                            } else if (r.prefix === 'ebay3' && data.ebay3_spft) {
                                value = Math.round(data.ebay3_spft);
                            } else if (r.prefix === 'doba' && data.doba_spft) {
                                value = Math.round(data.doba_spft);
                            } else if (r.prefix === 'temu' && data.temu_spft) {
                                value = Math.round(data.temu_spft);
                            } else if (r.prefix === 'macy' && data.macy_spft) {
                                value = Math.round(data.macy_spft);
                            } else if (r.prefix === 'reverb' && data.reverb_spft) {
                                value = Math.round(data.reverb_spft);
                            } else if (r.prefix === 'walmart' && data.walmart_spft) {
                                value = Math.round(data.walmart_spft);
                            }
                            else if (r.prefix === 'shein' && data.shein_spft) {
                                value = Math.round(data.shein_spft);
                            }

                            if (value !== undefined) {
                                if (value < 11) {
                                    textColor = '#ff0000';
                                } else if (value >= 10 && value < 15) {
                                    bgColor = 'yellow';
                                    textColor = '#000000';
                                } else if (value >= 15 && value < 20) {
                                    textColor = '#0d6efd';
                                } else if (value >= 21 && value < 50) {
                                    textColor = '#198754';
                                } else {
                                    textColor = '#800080';
                                }
                                
                                return `<span style="color: ${textColor}; ${bgColor ? `background-color: ${bgColor};` : ''}">${value}%</span>`;
                            }
                            
                            return '-';
                        })()}
                    </td>

                    <td class="sroi-field">
                        ${(() => {
                            let value, textColor, bgColor;
                            
                            if (r.prefix === 'amz' && data.amz_sroi) {
                                value = Math.round(data.amz_sroi);
                            } else if (r.prefix === 'ebay' && data.ebay_sroi) {
                                value = Math.round(data.ebay_sroi);
                            } else if (r.prefix === 'shopifyb2c' && data.shopifyb2c_sroi) {
                                value = Math.round(data.shopifyb2c_sroi);
                            } else if (r.prefix === 'ebay2' && data.ebay2_sroi) {
                                value = Math.round(data.ebay2_sroi);
                            } else if (r.prefix === 'ebay3' && data.ebay3_sroi) {
                                value = Math.round(data.ebay3_sroi);
                            }else if (r.prefix === 'doba' && data.doba_sroi) {
                                value = Math.round(data.doba_sroi);
                            } else if (r.prefix === 'temu' && data.temu_sroi) {
                                value = Math.round(data.temu_sroi);
                            } else if (r.prefix === 'macy' && data.macy_sroi) {
                                value = Math.round(data.macy_sroi);
                            } else if (r.prefix === 'reverb' && data.reverb_sroi) {
                                value = Math.round(data.reverb_sroi);
                            } else if (r.prefix === 'walmart' && data.walmart_sroi) {
                                value = Math.round(data.walmart_sroi);
                            } else if (r.prefix === 'shein' && data.shein_sroi) {
                                value = Math.round(data.shein_sroi);
                            }

                            if (value !== undefined) {
                                if (value < 11) {
                                    textColor = '#ff0000';
                                } else if (value >= 10 && value < 15) {
                                    bgColor = 'yellow';
                                    textColor = '#000000';
                                } else if (value >= 15 && value < 20) {
                                    textColor = '#0d6efd';
                                } else if (value >= 21 && value < 50) {
                                    textColor = '#198754';
                                } else {
                                    textColor = '#800080';
                                }
                                
                                return `<span style="color: ${textColor}; ${bgColor ? `background-color: ${bgColor};` : ''}">${value}%</span>`;
                            }
                            
                            return '-';
                        })()}
                    </td>
                   

                </tr>
                `;
            });

            html += "</tbody></table></div>";
          

            return html;
        }

        // Modal open function
        function showOVL30Modal(row) {
            const data = row.getData();
            
            // Initialize top push button
            const topPushPrice = document.getElementById('topPushPrice');
            const topPushBtn = document.getElementById('topPushBtn');
            const topSaveBtn = document.getElementById('topSaveBtn');
            
            topPushBtn.dataset.sku = data.SKU;
            topSaveBtn.dataset.sku = data.SKU;
            topSaveBtn.dataset.lp = data.LP || 0;
            topSaveBtn.dataset.ship = data.SHIP || 0;
            topSaveBtn.dataset.temuShip = data.temu_ship || 0;
            topPushPrice.value = data.shopifyb2c_price || data.ebay_price || data.amz_price || '';
            document.getElementById('ovl30SkuLabel').textContent = data.SKU ? `${data.SKU}` : "0";     
            document.getElementById('ovl30InvLabel').textContent = data.INV ? `${data.INV}` : "0"; 
            document.getElementById('ovl30').textContent = data.shopifyb2c_l30 ? `${data.shopifyb2c_l30}` : "0";    
            document.getElementById('total_views').textContent = data.total_views ? `${data.total_views}` : "0";  
            document.getElementById('avgCvr').textContent = data.avgCvr ? `${data.avgCvr}` : "0";        
            const imgEl = document.getElementById('ovl30Img');

            if (imgEl) {
                if (data.shopifyb2c_image) {
                    imgEl.src = data.shopifyb2c_image;
                    imgEl.style.display = "block";   // show image
                } else {
                    imgEl.style.display = "none";    // hide if missing
                }
            }


            document.getElementById('dilPercentage').textContent = data.dilPercentage ? `${data.dilPercentage}` : "0";
            if (data.dilPercentage) {
                const dilElement = document.getElementById('dilPercentage');
                const rounded = data.dilPercentage;
                
                if (rounded >= 0 && rounded <= 10) {
                    dilElement.style.color = "red";
                } else if (rounded >= 11 && rounded <= 15) {
                    dilElement.style.backgroundColor = "yellow";
                    dilElement.style.color = "black"; 
                    dilElement.style.padding = "2px 4px";
                    dilElement.style.borderRadius = "4px";
                } else if (rounded >= 16 && rounded <= 20) {
                    dilElement.style.color = "blue";
                } else if (rounded >= 21 && rounded <= 40) {
                    dilElement.style.color = "green";
                } else if (rounded >= 41) {
                    dilElement.style.color = "purple";
                }
            }
            document.getElementById('formattedAvgPrice').textContent = data.formattedAvgPrice ? `${data.formattedAvgPrice}` : " 0";
            if (data.formattedAvgPrice) {
                const avgPriceValue = parseFloat(data.formattedAvgPrice.replace(/[^0-9.-]+/g, ''));
                let textColor;
                if (!isNaN(avgPriceValue)) {
                    if (avgPriceValue < 10) {
                        textColor = '#dc3545'; // red
                    } else if (avgPriceValue >= 10 && avgPriceValue < 15) {
                        textColor = '#fd7e14'; // orange
                    } else if (avgPriceValue >= 15 && avgPriceValue < 20) {
                        textColor = '#0d6efd'; // blue
                    } else if (avgPriceValue >= 20) {
                        textColor = '#198754'; // green
                    }
                } else {
                    textColor = '#6c757d'; // gray
                }
                document.getElementById('formattedAvgPrice').style.color = textColor;
            }
            document.getElementById('formattedProfitPercentage').textContent = data.avgPftPercent ? `${data.avgPftPercent}` : "0";
            if (data.avgPftPercent) {
                let bgColor, textColor;
                const avgPftPercent = data.avgPftPercent;
                
                if (avgPftPercent < 11) {
                    textColor = '#ff0000';
                } else if (avgPftPercent >= 10 && avgPftPercent < 15) {
                    bgColor = 'yellow';
                    textColor = '#000000';
                } else if (avgPftPercent >= 15 && avgPftPercent < 20) {
                    textColor = '#0d6efd';
                } else if (avgPftPercent >= 21 && avgPftPercent < 50) {
                    textColor = '#198754';
                } else {
                    textColor = '#800080';
                }
                
                const element = document.getElementById('formattedProfitPercentage');
                element.style.color = textColor;
                if (bgColor) {
                    element.style.backgroundColor = bgColor;
                }
            }
            document.getElementById('formattedRoiPercentage').textContent = data.avgRoi ? `${data.avgRoi}` : "0";
            if (data.avgRoi) {
                let bgColor, textColor;
                const avgRoi = data.avgRoi;
                
                if (avgRoi < 11) {
                    textColor = '#ff0000';
                } else if (avgRoi >= 10 && avgRoi < 15) {
                    bgColor = 'yellow';
                    textColor = '#000000'; 
                } else if (avgRoi >= 15 && avgRoi < 20) {
                    textColor = '#0d6efd';
                } else if (avgRoi >= 21 && avgRoi < 50) {
                    textColor = '#198754';
                } else {
                    textColor = '#800080';
                }
                
                const element = document.getElementById('formattedRoiPercentage');
                element.style.color = textColor;
                if (bgColor) {
                    element.style.backgroundColor = bgColor;
                }
            }



            document.getElementById('ovl30Content').innerHTML = buildOVL30Table(data);

            const modalEl = document.getElementById('ovl30Modal');
            const modal = new bootstrap.Modal(modalEl);

            // Automatically sort by L30 (highest to lowest) when modal opens
            setTimeout(() => {
                const table = modalEl.querySelector('.sortable-table');
                const l30Header = Array.from(table.querySelectorAll('th')).find(th => th.textContent.includes('L30'));
                if (l30Header) {
                    // Trigger two clicks if needed to get descending order (highest to lowest)
                    if (!l30Header.classList.contains('sort-desc')) {
                        l30Header.click();
                        if (!l30Header.classList.contains('sort-desc')) {
                            l30Header.click();
                        }
                    }
                }
            }, 100);

            // Make modal draggable
            const dialogEl = modalEl.querySelector('.modal-dialog');
            let isDragging = false;
            let currentX;
            let currentY;
            let initialX;
            let initialY;
            let xOffset = 0;
            let yOffset = 0;

            dialogEl.addEventListener('mousedown', dragStart);
            document.addEventListener('mousemove', drag);
            document.addEventListener('mouseup', dragEnd);

            function dragStart(e) {
                if (e.target.closest('.modal-header')) {
                    isDragging = true;
                    initialX = e.clientX - xOffset;
                    initialY = e.clientY - yOffset;
                }
            }

            function drag(e) {
                if (isDragging) {
                    e.preventDefault();
                    currentX = e.clientX - initialX;
                    currentY = e.clientY - initialY;
                    xOffset = currentX;
                    yOffset = currentY;
                    dialogEl.style.transform = `translate(${currentX}px, ${currentY}px)`;
                }
            }

            function dragEnd() {
                isDragging = false;
            }

            // Reset position when modal is hidden
            modalEl.addEventListener('hidden.bs.modal', function() {
                dialogEl.style.transform = 'none';
                xOffset = 0;
                yOffset = 0;
            });

            // Initialize table sorting
            initTableSorting(modalEl.querySelector('.sortable-table'));
            modal.show();
        }
    
    // Calculate dilPercentage using the formula: L30 / INV
    let dilPercentage = 0;
    const l30 = ovl30Value; // Use the same ovl30Value for consistency
    const inv = parseFloat(data.INV) || 0;
    if (inv !== 0) {
        dilPercentage = (l30 / inv) * 100;
    }
    const dilElement = document.getElementById('dilPercentage');
    dilElement.textContent = dilPercentage ? `${Math.round(dilPercentage)}` : "0%";
    // Apply color based on dilPercentage ranges
    if (dilPercentage) {
        const rounded = dilPercentage;
        if (rounded >= 0 && rounded <= 10) {
            dilElement.style.color = "red";
            dilElement.style.backgroundColor = "";
            dilElement.style.padding = "";
            dilElement.style.borderRadius = "";
        } else if (rounded > 10 && rounded <= 15) {
            dilElement.style.color = "black";
            dilElement.style.backgroundColor = "yellow";
            dilElement.style.padding = "2px 4px";
            dilElement.style.borderRadius = "4px";
        } else if (rounded > 15 && rounded <= 20) {
            dilElement.style.color = "blue";
            dilElement.style.backgroundColor = "";
            dilElement.style.padding = "";
            dilElement.style.borderRadius = "";
        } else if (rounded > 20 && rounded <= 40) {
            dilElement.style.color = "green";
            dilElement.style.backgroundColor = "";
            dilElement.style.padding = "";
            dilElement.style.borderRadius = "";
        } else if (rounded > 40) {
            dilElement.style.color = "purple";
            dilElement.style.backgroundColor = "";
            dilElement.style.padding = "";
            dilElement.style.borderRadius = "";
        }
    } else {
        dilElement.style.color = "#6c757d";
        dilElement.style.backgroundColor = "";
        dilElement.style.padding = "";
        dilElement.style.borderRadius = "";
    }
    // Rest of the modal fields
    document.getElementById('formattedAvgPrice').textContent = data.formattedAvgPrice ? `${data.formattedAvgPrice}` : "0";
    if (data.formattedAvgPrice) {
        const avgPriceValue = parseFloat(data.formattedAvgPrice.replace(/[^0-9.-]+/g, ''));
        let textColor;
        if (!isNaN(avgPriceValue)) {
            if (avgPriceValue < 10) {
                textColor = '#dc3545';
            } else if (avgPriceValue >= 10 && avgPriceValue < 15) {
                textColor = '#fd7e14';
            } else if (avgPriceValue >= 15 && avgPriceValue < 20) {
                textColor = '#0d6efd';
            } else if (avgPriceValue >= 20) {
                textColor = '#198754';
            }
        } else {
            textColor = '#6c757d';
        }
        document.getElementById('formattedAvgPrice').style.color = textColor;
    }
    document.getElementById('formattedProfitPercentage').textContent = data.avgPftPercent ? `${data.avgPftPercent}` : "0";
    if (data.avgPftPercent) {
        let bgColor, textColor;
        const avgPftPercent = data.avgPftPercent;
        if (avgPftPercent < 11) {
            textColor = '#ff0000';
        } else if (avgPftPercent >= 10 && avgPftPercent < 15) {
            bgColor = 'yellow';
            textColor = '#000000';
        } else if (avgPftPercent >= 15 && avgPftPercent < 20) {
            textColor = '#0d6efd';
        } else if (avgPftPercent >= 21 && avgPftPercent < 50) {
            textColor = '#198754';
        } else {
            textColor = '#800080';
        }
        const element = document.getElementById('formattedProfitPercentage');
        element.style.color = textColor;
        if (bgColor) {
            element.style.backgroundColor = bgColor;
        }
    }
    document.getElementById('formattedRoiPercentage').textContent = data.avgRoi ? `${data.avgRoi}` : "0";
    if (data.avgRoi) {
        let bgColor, textColor;
        const avgRoi = data.avgRoi;
        if (avgRoi < 11) {
            textColor = '#ff0000';
        } else if (avgRoi >= 10 && avgRoi < 15) {
            bgColor = 'yellow';
            textColor = '#000000';
        } else if (avgRoi >= 15 && avgRoi < 20) {
            textColor = '#0d6efd';
        } else if (avgRoi >= 21 && avgRoi < 50) {
            textColor = '#198754';
        } else {
            textColor = '#800080';
        }
        const element = document.getElementById('formattedRoiPercentage');
        element.style.color = textColor;
        if (bgColor) {
            element.style.backgroundColor = bgColor;
        }
    }
    document.getElementById('ovl30Content').innerHTML = buildOVL30Table(data);
    const modalEl = document.getElementById('ovl30Modal');
    const modal = new bootstrap.Modal(modalEl);
    setTimeout(() => {
        const table = modalEl.querySelector('.sortable-table');
        const l30Header = Array.from(table.querySelectorAll('th')).find(th => th.textContent.includes('L30'));
        if (l30Header) {
            if (!l30Header.classList.contains('sort-desc')) {
                l30Header.click();
                if (!l30Header.classList.contains('sort-desc')) {
                    l30Header.click();
                }
            }
        }
    }, 100);
    const dialogEl = modalEl.querySelector('.modal-dialog');
    let isDragging = false;
    let currentX;
    let currentY;
    let initialX;
    let initialY;
    let xOffset = 0;
    let yOffset = 0;
    dialogEl.addEventListener('mousedown', dragStart);
    document.addEventListener('mousemove', drag);
    document.addEventListener('mouseup', dragEnd);
    function dragStart(e) {
        if (e.target.closest('.modal-header')) {
            isDragging = true;
            initialX = e.clientX - xOffset;
            initialY = e.clientY - yOffset;
        }
    }
    function drag(e) {
        if (isDragging) {
            e.preventDefault();
            currentX = e.clientX - initialX;
            currentY = e.clientY - initialY;
            xOffset = currentX;
            yOffset = currentY;
            dialogEl.style.transform = `translate(${currentX}px, ${currentY}px)`;
        }
    }
    function dragEnd() {
        isDragging = false;
    }
    modalEl.addEventListener('hidden.bs.modal', function() {
        dialogEl.style.transform = 'none';
        xOffset = 0;
        yOffset = 0;
    });
    initTableSorting(modalEl.querySelector('.sortable-table'));
    modal.show();

        // Table sorting functionality
        function initTableSorting(table) {
            const headers = table.querySelectorAll('th[data-sort]');
            headers.forEach(header => {
                header.style.cursor = 'pointer';
                header.addEventListener('click', () => {
                    const sortType = header.getAttribute('data-sort');
                    const columnIndex = Array.from(header.parentElement.children).indexOf(header);
                    const rows = Array.from(table.querySelector('tbody').rows);
                    const isAscending = header.classList.contains('sort-asc');

                    // Remove sorting classes from all headers
                    headers.forEach(h => {
                        h.classList.remove('sort-asc', 'sort-desc');
                        h.querySelector('.bi').className = 'bi bi-arrow-down-up';
                    });

                    // Sort the rows
                    rows.sort((a, b) => {
                        let aVal = a.cells[columnIndex].textContent.trim();
                        let bVal = b.cells[columnIndex].textContent.trim();

                        if (sortType === 'number') {
                            // Extract numbers from strings and convert to float
                            aVal = parseFloat(aVal.replace(/[^0-9.-]+/g, '')) || 0;
                            bVal = parseFloat(bVal.replace(/[^0-9.-]+/g, '')) || 0;
                        }

                        if (aVal === bVal) return 0;
                        if (isAscending) {
                            return sortType === 'string' ? 
                                bVal.localeCompare(aVal) : 
                                bVal - aVal;
                        } else {
                            return sortType === 'string' ? 
                                aVal.localeCompare(bVal) : 
                                aVal - bVal;
                        }
                    });

                    // Update sorting indicators
                    header.classList.add(isAscending ? 'sort-desc' : 'sort-asc');
                    header.querySelector('.bi').className = `bi bi-arrow-${isAscending ? 'down' : 'up'}`;

                    // Reorder the rows in the table
                    const tbody = table.querySelector('tbody');
                    rows.forEach(row => tbody.appendChild(row));
                });
            });
        }

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

                // Filter table by Parent
                table.setFilter("Parent", "=", parentKey);
                console.log("Showing group:", parentKey);
            }

            // ▶️ Play (activate filter mode, start at first group)
            document.getElementById('play-auto').addEventListener('click', () => {
                isPlaying = true;
                currentIndex = 0;
                renderGroup(parentKeys()[currentIndex]);

                document.getElementById('play-pause').style.display = 'inline-block';
                document.getElementById('play-auto').style.display = 'none';
            });

            // ⏸ Pause (show all data again)
            document.getElementById('play-pause').addEventListener('click', () => {
                isPlaying = false;
                currentParentFilter = null; 
                setCombinedFilters();
                table.clearFilter();

                document.getElementById('play-pause').style.display = 'none';
                document.getElementById('play-auto').style.display = 'inline-block';
            });

            // ⏩ Forward (manual step)
            document.getElementById('play-forward').addEventListener('click', () => {
                if (!isPlaying) return;
                currentIndex = (currentIndex + 1) % parentKeys().length;
                renderGroup(parentKeys()[currentIndex]);
            });

            // ⏪ Backward (manual step)
            document.getElementById('play-backward').addEventListener('click', () => {
                if (!isPlaying) return;
                currentIndex = (currentIndex - 1 + parentKeys().length) % parentKeys().length;
                renderGroup(parentKeys()[currentIndex]);
            });
        });

        // Draggable Modal for Chart 
        document.addEventListener("DOMContentLoaded", function () {
            const modal = document.querySelector("#priceComparisonModal .modal-dialog");
            const header = document.querySelector("#priceComparisonModal .modal-header");

            let isDragging = false;
            let offsetX, offsetY;

            header.style.cursor = "move";

            header.addEventListener("mousedown", (e) => {
                isDragging = true;
                const rect = modal.getBoundingClientRect();
                offsetX = e.clientX - rect.left;
                offsetY = e.clientY - rect.top;
                modal.style.position = "absolute";
                modal.style.margin = "0";
            });

            document.addEventListener("mousemove", (e) => {
                if (isDragging) {
                    modal.style.left = e.clientX - offsetX + "px";
                    modal.style.top = e.clientY - offsetY + "px";
                }
            });


            document.addEventListener("mouseup", () => {
                isDragging = false;
            });

            // Reset position when modal is closed
            document.getElementById("priceComparisonModal").addEventListener("hidden.bs.modal", function () {
                modal.style.position = "";
                modal.style.left = "";
                modal.style.top = "";
                modal.style.margin = "";
            });
        });


        // Push Price
        $(document).on('blur', '.s-price', function() {
            const $input = $(this);
            const sprice = parseFloat($input.val()) || 0;
            const sku = $input.data('sku');
            const type = $input.data('type');
            const LP = parseFloat($input.data('lp')) || 0;
            const SHIP = parseFloat($input.data('ship')) || 0;
            const temu_ship = parseFloat($input.data('temu_ship')) || 0;

            if (!sku || !type) return;

            $.ajax({
                url: '/pricing-master/save-sprice',
                type: 'POST',
                data: {
                    _token: $('meta[name="csrf-token"]').attr('content'),
                    sku: sku,
                    type: type,
                    sprice: sprice,
                    LP: LP,
                    SHIP: SHIP,
                    temu_ship: temu_ship
                },
                 beforeSend: function() {
                        $('#savePricingBtn').html(
                            '<i class="fa fa-spinner fa-spin"></i> Saving...');
                    },
                success: function(res) {
                    if(res.status === 200) {
                        const $row = $input.closest('tr');
                        const spft = Math.round(res.data.SPFT);
                        const sroi = Math.round(res.data.SROI);

                        function getColoredSpan(value) {
                            let textColor, bgColor;

                            if (value < 11) {
                                textColor = '#ff0000';
                            } else if (value >= 10 && value < 15) {
                                bgColor = 'yellow';
                                textColor = '#000000';
                            } else if (value >= 15 && value < 20) {
                                textColor = '#0d6efd';
                            } else if (value >= 21 && value < 50) {
                                textColor = '#198754';
                            } else {
                                textColor = '#800080';
                            }

                            return `<span style="color:${textColor};${bgColor ? `background-color:${bgColor};` : ''}">${value}%</span>`;
                        }

                        // Update with styled spans
                        $row.find('.spft-field').html(getColoredSpan(spft));
                        $row.find('.sroi-field').html(getColoredSpan(sroi));
                    } else {
                        console.error('Error saving S Price:', res.message);
                    }
                },

                error: function(err) {
                    console.error('Error saving S Price:', err);
                }
            });
        });


        // Save top price button handler
        $(document).on('click', '#topSaveBtn', function() {
            const sku = $(this).data('sku');
            const price = parseFloat($('#topPushPrice').val()) || 0;
            const LP = parseFloat($(this).data('lp')) || 0;
            const SHIP = parseFloat($(this).data('ship')) || 0;
            const temu_ship = parseFloat($(this).data('temuShip')) || 0;

            console.log('Saving for SKU:', sku, 'Price:', price);

            if (!sku || price <= 0) {
                alert('Please enter a valid price');
                return;
            }

            $.ajax({
                url: '/pricing-master/save-sprice',
                type: 'POST',
                data: {
                    _token: $('meta[name="csrf-token"]').attr('content'),
                    sku: sku,
                    type: 'top',
                    sprice: price,
                    LP: LP,
                    SHIP: SHIP,
                    temu_ship: temu_ship
                },
                success: function(res) {
                    if (res.status === 200) {
                        alert('Price saved to all marketplaces successfully');
                    } else {
                        alert('Failed to save price');
                    }
                },
                error: function(err) {
                    console.error('Error saving price:', err);
                    alert('Error saving price');
                }
            });
        });

            $(document).on('click', '.pushPriceBtn, #topPushBtn', function() {
            const $btn = $(this);
            const sku = $btn.data('sku') || $('#topPushBtn').data('sku');
            let price;
            
            if($btn.attr('id') === 'topPushBtn') {
                price = parseFloat($('#topPushPrice').val()) || 0;
                if(price <= 0) {
                    alert('Please enter a valid price');
                    return;
                }
                // Push to all marketplaces
                $.ajax({
                    url: '/push-shopify-price',
                    type: 'POST',
                    data: { 
                        sku: sku, 
                        price: price,
                        _token: '{{ csrf_token() }}'
                    },
                    success: function(response) {
                        console.log('Shopify price updated');
                    },
                    error: function(xhr, status, error) {
                        console.error('Shopify update failed:', error);
                    }
                });
                
                $.ajax({
                    url: '/push-ebay-price',
                    type: 'POST',
                    data: { 
                        sku: sku, 
                        price: price,
                        _token: '{{ csrf_token() }}'
                    },
                    success: function(response) {
                        console.log('eBay price updated');
                    },
                    error: function(xhr, status, error) {
                        console.error('eBay update failed:', error);
                    }
                });
                
                $.ajax({
                    url: '/update-amazon-price',
                    type: 'POST',
                    data: { 
                        sku: sku, 
                        price: price,
                        _token: '{{ csrf_token() }}'
                    },
                    success: function(response) {
                        console.log('Amazon price updated');
                    },
                    error: function(xhr, status, error) {
                        console.error('Amazon update failed:', error);
                    }
                });


                 // ✅ Walmart
                $.ajax({
                    url: '/push-walmart-price',
                    type: 'POST',
                    data: { sku: sku, price: price, _token: '{{ csrf_token() }}' },
                    success: function(response) {
                        console.log('Walmart price update requested');
                    },
                    error: function(err) {
                        console.error('Walmart update failed:', err);
                    }
                });

                // ✅ Doba
                $.ajax({
                    url: '/update-doba-price',
                    type: 'POST',
                    data: { sku: sku, price: price, _token: '{{ csrf_token() }}' },
                    success: function(response) {
                        console.log('Doba price update requested');
                    },
                    error: function(err) {
                        console.error('Doba update failed:', err);
                    }
                });
                
                alert('Price is being updated across all marketplaces');
                return;
            }
            
            const type = $btn.data('type');

            if(!sku || !type) return;

            // Fetch the latest SPrice from your input field in the modal/table
            const $input = $btn.closest('tr').find('.s-price');
            const sprice = parseFloat($input.val()) || 0;

            if(sprice <= 0) {
                alert('Please enter a valid SPrice before pushing.');
                return;
            }

            $btn.html('<i class="fa fa-spinner fa-spin"></i> Pushing...');

            // Example: Marketplace AJAX
            if(type === 'amz') {
                $.ajax({
                    url: '/update-amazon-price',
                    type: 'POST',
                    data: {
                        _token: $('meta[name="csrf-token"]').attr('content'),
                        sku: sku,
                        price: sprice
                    },
                    success: function(res) {
                       alert('Amazon price updated successfully!');
                    },
                    error: function(err) {
                        alert('Error updating Amazon price: ' + err);
                    },
                    complete: function() {
                        $btn.html('Push to Marketplace'); // reset button text
                    }
                });
            }
             else if(type === 'ebay') {
                $.ajax({
                    url: '/push-ebay-price',
                    type: 'POST',
                    data: {
                        _token: $('meta[name="csrf-token"]').attr('content'),
                        sku: sku,
                        price: sprice
                    },
                    success: function(res) {
                        if (res.success) {
                            alert('eBay price updated successfully!');
                        } else {
                            alert('Error: ' + (res.error || 'Unknown error'));
                        }
                    },
                    error: function(err) {
                        const errorMsg = err.responseJSON && err.responseJSON.error ? 
                            err.responseJSON.error : 'Error updating eBay price';
                        alert(errorMsg);
                    },
                    complete: function() {
                        $btn.html('<i class="bi bi-cloud-arrow-up"></i>');
                    }
                });
            }

              else if(type === 'ebay2') {
                $.ajax({
                    url: '/push-ebay2-price',
                    type: 'POST',
                    data: {
                        _token: $('meta[name="csrf-token"]').attr('content'),
                        sku: sku,
                        price: sprice
                    },
                    success: function(res) {
                        if (res.success) {
                            alert('eBay2 price updated successfully!');
                        } else {
                            alert('Error: ' + (res.error || 'Unknown error'));
                        }
                    },
                    error: function(err) {
                        const errorMsg = err.responseJSON && err.responseJSON.error ? 
                            err.responseJSON.error : 'Error updating eBay2 price';
                        alert(errorMsg);
                    },
                    complete: function() {
                        $btn.html('<i class="bi bi-cloud-arrow-up"></i>');
                    }
                });
            }
              else if(type === 'ebay3') {
                $.ajax({
                    url: '/push-ebay3-price',
                    type: 'POST',
                    data: {
                        _token: $('meta[name="csrf-token"]').attr('content'),
                        sku: sku,
                        price: sprice
                    },
                    success: function(res) {
                        if (res.success) {
                            alert('eBay3 price updated successfully!');
                        } else {
                            alert('Error: ' + (res.error || 'Unknown error'));
                        }
                    },
                    error: function(err) {
                        const errorMsg = err.responseJSON && err.responseJSON.error ? 
                            err.responseJSON.error : 'Error updating eBay3 price';
                        alert(errorMsg);
                    },
                    complete: function() {
                        $btn.html('<i class="bi bi-cloud-arrow-up"></i>');
                    }
                });
            }

            else if(type === 'shopifyb2c') {
                $.ajax({
                    url: '/push-shopify-price',
                    type: 'POST',
                    data: {
                        _token: $('meta[name="csrf-token"]').attr('content'),
                        sku: sku,
                        price: sprice
                    },
                    success: function(res) {
                        if (res.status === "success") {
                            alert(res.message || 'Shopify price updated successfully!');
                        } else {
                            alert("Error: " + (res.message || "Something went wrong"));
                        }
                        console.log(res);
                    },
                    error: function(err) {
                        let errorMsg = "Error updating Shopify price!";
                        if (err.responseJSON && err.responseJSON.message) {
                            errorMsg = err.responseJSON.message;
                        }
                        alert(errorMsg);
                        console.error('Error updating Shopify price:', err);
                    },
                    complete: function() {
                        $btn.html('Push to Marketplace');
                    }
                });
            }  else if(type === 'temu') {
                $.ajax({
                    url: '/update-temu-price',
                    type: 'POST',
                    data: {
                        _token: $('meta[name="csrf-token"]').attr('content'),
                        sku: sku,
                        price: sprice
                    },
                    success: function(res) {
                     alert('Temu price updated successfully!');
                    },
                    error: function(err) {
                        alert('Error updating Temu price: ' + err);
                    },
                    complete: function() {
                        $btn.html('Push to Marketplace');
                    }
                });
            } else if(type === 'macy') {
                $.ajax({
                    url: '/update-macy-price',
                    type: 'POST',
                    data: {
                            _token: $('meta[name="csrf-token"]').attr('content'),
                            sku: sku,
                            price: sprice
                        },
                        success: function(res) {
                        alert('Macy price updated successfully!');
                        },
                        error: function(err) {
                            alert('Error updating Macy price: ' + err);
                        },
                        complete: function() {
                            $btn.html('Push to Marketplace');
                        }
                    });
                }
                else if(type === 'reverb') {
                    $.ajax({
                        url: '/update-reverb-price',
                        type: 'POST',
                        data: {
                            _token: $('meta[name="csrf-token"]').attr('content'),
                            sku: sku,
                            price: sprice
                        },
                        success: function(res) {
                        alert('Reverb price updated successfully!');
                        },
                        error: function(err) {
                            alert('Error updating Reverb price: ' + err);
                        },
                        complete: function() {
                            $btn.html('Push to Marketplace');
                        }
                    });
                }
                else if(type === 'walmart') {
                    $.ajax({
                        url: '/push-walmart-price',
                        type: 'POST',
                        data: {
                            _token: $('meta[name="csrf-token"]').attr('content'),
                            sku: sku,
                            price: sprice
                        },
                        success: function(res) {
                            alert('Walmart Price Change Requested, Will Be Completed after 10  Minutes!');
                        },
                        error: function(err) {
                            alert('Error updating Walmart price: ' + err.responseText);
                        },
                        complete: function() {
                            $btn.html('Push to Marketplace');
                        }
                    });
                }

                else if(type === 'doba') {
                    $.ajax({
                        url: '/update-doba-price',
                        type: 'POST',
                        data: {
                            _token: $('meta[name="csrf-token"]').attr('content'),
                            sku: sku,
                            price: sprice
                        },
                        success: function(res) {
                            alert('Doba Price Change Requested, Will Be Completed after 10  Minutes!');
                        },
                        error: function(err) {
                            alert('Error updating Doba price: ' + err.responseText);
                        },
                        complete: function() {
                            $btn.html('Push to Marketplace');
                        }
                    });
                }
          });

       
    </script>

    <script>
function showTooltip(img) {
    const tooltip = img.nextElementSibling;
    if (tooltip && tooltip.classList.contains('link-tooltip')) {
        tooltip.style.opacity = '1';
        tooltip.style.visibility = 'visible';
    }
}

function hideTooltip(img) {
    const tooltip = img.nextElementSibling;
    if (tooltip && tooltip.classList.contains('link-tooltip')) {
        tooltip.style.opacity = '0';
        tooltip.style.visibility = 'hidden';
    }
}
</script>
@endsection
