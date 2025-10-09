@extends('layouts.vertical', ['title' => 'Account Health Master Dashboard', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
<meta name="csrf-token" content="{{ csrf_token() }}">

@section('css')
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
<link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
<style>
    /* Custom styles for the Tabulator table */
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

    .custom-select-wrapper {
        width: 100%;
        cursor: pointer;
        position: relative;
    }

    .custom-select-display {
        background-color: #fff;
        border: 1px solid #ced4da;
        padding: 0.375rem 0.75rem;
        border-radius: 0.375rem;
    }

    .custom-select-options {
        position: absolute;
        z-index: 999;
        top: 100%;
        left: 0;
        right: 0;
        max-height: 200px;
        overflow-y: auto;
        border: 1px solid #ced4da;
        border-top: none;
        background-color: #fff;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
    }

    .custom-select-search {
        width: 100%;
        padding: 0.5rem;
        border: none;
        border-bottom: 1px solid #eee;
        outline: none;
    }

    .custom-select-option {
        padding: 0.5rem 0.75rem;
        cursor: pointer;
    }

    .custom-select-option:hover {
        background-color: #f1f1f1;
    }

</style>
@endsection

@section('content')
    @include('layouts.shared.page-title', [
        'page_title' => 'Account Health Master Dashboard',
        'sub_title' => 'Manage your channels and monitor their performance metrics',
    ])
    <div class="container-fluid">
        <div class="col-md-12 mt-0 pt-0 mb-1 pb-1">
            <div class="row justify-content-center align-items-center g-3">
                <div class="col d-flex align-items-center gap-2 flex-wrap">
                    <div class="btn-group time-navigation-group" role="group" aria-label="Parent navigation">
                        <button id="play-backward" class="btn btn-light rounded-circle" title="Previous parent">
                            <i class="fas fa-step-backward"></i>
                        </button>
                        <button id="play-pause" class="btn btn-light rounded-circle" title="Pause playback"
                            style="display: none;">
                            <i class="fas fa-pause"></i>
                        </button>
                        <button id="play-auto" class="btn btn-light rounded-circle" title="Start auto-play">
                            <i class="fas fa-play"></i>
                        </button>
                        <button id="play-forward" class="btn btn-light rounded-circle" title="Next parent">
                            <i class="fas fa-step-forward"></i>
                        </button>
                    </div>
                    <button id="addChannelBtn" class="btn btn-primary" data-bs-toggle="modal"
                        data-bs-target="#addChannelModal"
                        style="background: linear-gradient(135deg, #4361ee, #3f37c9); border: none;">
                        <i class="fas fa-plus-circle me-2"></i> Add Channel
                    </button>
                    <div class="d-flex flex-wrap gap-2 align-items-center">
                        <a href="{{ route('account-health-master.export') }}" class="btn btn-success">
                            <i class="fas fa-file-export me-1"></i> Export Health Data
                        </a>
                        <button type="button" class="btn btn-primary" data-bs-toggle="modal"
                            data-bs-target="#accountHealthImportModal">
                            <i class="fas fa-file-import me-1"></i> Import Health Data
                        </button>
                    </div>
                </div>
                <div class="col-auto">
                    <div class="dropdown-search-container" style="position: relative;">
                        <input type="text" class="form-control form-control-sm channel-search"
                            placeholder="Search Channel" id="channelSearchInput">
                        <div class="dropdown-search-results" id="channelSearchDropdown"
                            style="position: absolute; top: 100%; left: 0; right: 0; z-index: 9999; background: #fff; border: 1px solid #ccc; display: none; max-height: 200px; overflow-y: auto;">
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Add Channel Modal -->
        <div class="modal fade right-to-left" id="addChannelModal" tabindex="-1" aria-hidden="true">
            <div class="modal-dialog modal-side">
                <div class="modal-content">
                    <div class="modal-header bg-primary text-white">
                        <h5 class="modal-title"><i class="fas fa-plus-circle me-2"></i> Add New Channel</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"
                            aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <form id="channelForm">
                            <div class="row mb-3">
                                <div class="col-md-12">
                                    <label for="channelName" class="form-label">Channel Name</label>
                                    <input type="text" class="form-control" id="channelName" required>
                                </div>
                            </div>
                            <div class="mb-3">
                                <label for="channelUrl" class="form-label">Sheet Link</label>
                                <input type="url" class="form-control" id="channelUrl">
                            </div>
                            <div class="mb-3">
                                <label for="type" class="form-label">Type</label>
                                <input type="text" class="form-control" id="type">
                            </div>
                        </form>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="button" class="btn btn-primary" id="saveChannelBtn">Save Channel</button>
                    </div>
                </div>
            </div>
        </div>

        <!-- Edit Channel Modal -->
        <div class="modal fade" id="editChannelModal" tabindex="-1" aria-labelledby="editChannelModalLabel"
            aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <form id="editChannelForm">
                        <div class="modal-header">
                            <h5 class="modal-title" id="editChannelModalLabel">Edit Channel</h5>
                            <button type="button" class="btn-close" data-bs-dismiss="modal"
                                aria-label="Close"></button>
                        </div>
                        <div class="modal-body">
                            <input type="hidden" id="originalChannel" name="original_channel">
                            <div class="mb-3">
                                <label for="editChannelName" class="form-label">Channel Name</label>
                                <input type="text" class="form-control" id="editChannelName" name="channel_name"
                                    readonly>
                            </div>
                            <div class="mb-3">
                                <label for="editChannelUrl" class="form-label">Sheet URL</label>
                                <input type="text" class="form-control" id="editChannelUrl" name="sheet_url"
                                    required>
                            </div>
                            <div class="mb-3">
                                <label for="editType" class="form-label">Type</label>
                                <input type="text" class="form-control" id="editType" name="type" required>
                            </div>
                        </div>
                        <div class="modal-footer">
                            <button type="submit" class="btn btn-primary">Update Channel</button>
                        </div>
                    </form>
                </div>
            </div>
        </div>

        <!-- Import Modal -->
        <div class="modal fade" id="accountHealthImportModal" tabindex="-1"
            aria-labelledby="accountHealthImportModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <form action="{{ route('account-health-master.import') }}" method="POST" enctype="multipart/form-data"
                    class="modal-content" id="accountHealthImportForm">
                    @csrf
                    <div class="modal-header">
                        <h5 class="modal-title" id="accountHealthImportModalLabel">Import Account Health Data</h5>
                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <div class="mb-3">
                            <label for="accountHealthExcelFile" class="form-label">Select Excel File</label>
                            <input type="file" class="form-control" id="accountHealthExcelFile" name="excel_file"
                                accept=".xlsx,.xls,.csv" required>
                        </div>
                        <div class="mb-3">
                            <label for="importType" class="form-label">Import Type</label>
                            <select class="form-control" id="importType" name="import_type" required>
                                <option value="">Select Import Type</option>
                                <option value="channel_data">Channel Performance Data</option>
                                <option value="health_rates">Health Rates (ODR, Fulfillment, etc.)</option>
                                <option value="both">Both Channel Data & Health Rates</option>
                            </select>
                        </div>
                        <div class="mb-3">
                            <label for="updateMode" class="form-label">Update Mode</label>
                            <select class="form-control" id="updateMode" name="update_mode" required>
                                <option value="update">Update Existing Records</option>
                                <option value="create">Create New Records Only</option>
                                <option value="replace">Replace All Data</option>
                            </select>
                        </div>
                        <div class="alert alert-info">
                            <small>
                                <i class="fas fa-info-circle me-1"></i>
                                Download sample files:
                                <br>
                                • <a href="{{ route('account-health-master.sample', 'channel') }}"
                                    class="alert-link">Channel Data Sample</a>
                                <br>
                                • <a href="{{ route('account-health-master.sample', 'rates') }}" class="alert-link">Health
                                    Rates Sample</a>
                                <br>
                                • <a href="{{ route('account-health-master.sample', 'combined') }}"
                                    class="alert-link">Combined Sample</a>
                            </small>
                        </div>
                        <div class="progress mb-3" id="importProgress" style="display: none;">
                            <div class="progress-bar progress-bar-striped progress-bar-animated" role="progressbar"
                                style="width: 0%">0%</div>
                        </div>
                        <div id="importResults" class="alert alert-success" style="display: none;">
                            <h6>Import Results:</h6>
                            <ul id="importResultsList"></ul>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="submit" class="btn btn-primary" id="importSubmitBtn">
                            <i class="fas fa-file-import me-1"></i> Import
                        </button>
                    </div>
                </form>
            </div>
        </div>

        <div id="customLoader" style="display: flex; justify-content: center; align-items: center; height: 300px;">
            <div class="spinner-border text-info" role="status">
                <span class="sr-only">Loading...</span>
            </div>
            <span class="ms-2">Loading table, please wait...</span>
        </div>

        <div class="mb-4">
            <div id="channelSalesChart" style="width: 100%; height: 400px;"></div>
        </div>

        <div class="table-container" id="channelTableWrapper" style="display: none;">
            <div id="channelTable"></div>
        </div>
    </div>
@endsection

@section('script')
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://www.gstatic.com/charts/loader.js"></script>

    <script>
        var jq = jQuery.noConflict(true);
        let originalChannelData = [];
        let table;
        let isPlaying = false;
        let currentChannelIndex = 0;
        let uniqueChannels = [];
        let uniqueChannelRows = [];

        // Mapping of fields to their update endpoints
        const fieldUpdateEndpoints = {
            'odr_rate': '/odr-rate/update',
            'odr_rate_allowed': '/odr-rate/update',
            'fulfillment_rate': '/fullfillment-rate/update',
            'fulfillment_rate_allowed': '/fullfillment-rate/update',
            'valid_tracking_rate': '/validtracking-rate/update',
            'valid_tracking_rate_allowed': '/validtracking-rate/update',
            'on_time_delivery': '/onTimeDelivery-rate/update',
            'on_time_delivery_allowed': '/onTimeDelivery-rate/update',
            'atoz_claims_rate': '/atozclaims-rate/update',
            'atoz_claims_rate_allowed': '/atozclaims-rate/update',
            'violation_rate': '/violation-rate/update',
            'violation_rate_allowed': '/violation-rate/update',
            'late_shipment_rate': '/lateshipment-rate/update',
            'late_shipment_rate_allowed': '/lateshipment-rate/update',
            'negative_seller_rate': '/negativeseller-rate/update',
            'negative_seller_rate_allowed': '/negativeseller-rate/update',
            'refund_rate': '/refund-rate/update',
            'refund_rate_allowed': '/refund-rate/update'
        };

        function parseNumber(value) {
            if (value === null || value === undefined || value === '' || value === '#DIV/0!' || value === 'N/A') return 0;
            if (typeof value === 'number') return value;
            const cleaned = String(value).replace(/[^0-9.-]/g, '');
            return parseFloat(cleaned) || 0;
        }

        function initializeTable() {
            const toNum = (v, def = 0) => {
                const n = parseFloat(String(v).replace(/,/g, ''));
                return Number.isFinite(n) ? n : def;
            };
            const pick = (obj, keys, def = '') => {
                for (const k of keys) {
                    const v = obj[k];
                    if (v !== undefined && v !== null && v !== '') return v;
                }
                return def;
            };

            table = new Tabulator("#channelTable", {
                ajaxURL: "/account-health-master/dashboard-data",
                ajaxConfig: "GET",
                layout: "fitDataFill",
                pagination: true,
                paginationSize: 50,
                paginationMode: "local",
                movableColumns: false,
                resizableColumns: true,
                height: "550px",
                ajaxResponse: function(url, params, response) {
                    console.log('AJAX response:', response);
                    jq('#customLoader').hide();
                    jq('#channelTableWrapper').show();
                    if (!response || !response.data) {
                        console.warn('No data in response:', response);
                        alert('No data received from server.');
                        return [];
                    }
                    originalChannelData = response.data;
                    drawChannelChart(response.data);
                    updateL30OrdersTotal(response.data);
                    return response.data.map(item => ({
                        channel: pick(item, ['channel', 'Channel', 'Channel '], ''),
                        l30_orders: toNum(pick(item, ['l30_orders', 'L30 Orders'], 0), 0),
                        nr: toNum(pick(item, ['nr', 'NR'], 0), 0),
                        odr_rate: pick(item, ['odr_rate', 'ODR Rate', 'ODR'], 'N/A'),
                        odr_rate_allowed: pick(item, ['odr_rate_allowed'], ''),
                        fulfillment_rate: pick(item, ['fulfillment_rate', 'Fulfillment Rate'],
                            'N/A'),
                        fulfillment_rate_allowed: pick(item, ['fulfillment_rate_allowed'], ''),
                        valid_tracking_rate: pick(item, ['valid_tracking_rate',
                            'Valid Tracking Rate'
                        ], 'N/A'),
                        valid_tracking_rate_allowed: pick(item, ['valid_tracking_rate_allowed'],
                            ''),
                        on_time_delivery: pick(item, ['on_time_delivery_rate', 'On Time Delivery Rate'],
                            'N/A'),
                        on_time_delivery_allowed: pick(item, ['on_time_delivery_rate_allowed'], ''),
                        atoz_claims_rate: pick(item, ['atoz_claims_rate', 'AtoZ Claims Rate'],
                            'N/A'),
                        atoz_claims_rate_allowed: pick(item, ['atoz_claims_rate_allowed'], ''),
                        violation_rate: pick(item, ['violation_rate', 'Violation Rate'], 'N/A'),
                        violation_rate_allowed: pick(item, ['violation_rate_allowed'], ''),
                        late_shipment_rate: pick(item, ['late_shipment_rate', 'Late Shipment Rate'],
                            'N/A'),
                        late_shipment_rate_allowed: pick(item, ['late_shipment_rate_allowed'], ''),
                        negative_seller_rate: pick(item, ['negative_seller_rate',
                            'Negative Seller Rate'
                        ], 'N/A'),
                        negative_seller_rate_allowed: pick(item, ['negative_seller_rate_allowed'],
                            ''),
                        refund_rate: pick(item, ['refund_rate', 'Refund Rate'], 'N/A'),
                        refund_rate_allowed: pick(item, ['refund_rate_allowed'], ''),
                        sheet_link: pick(item, ['sheet_link'], ''),
                        type: pick(item, ['type'], '')
                    }));
                },
                ajaxError: function(xhr, error, thrown) {
                    console.error("AJAX error:", error, thrown, xhr.responseText);
                    jq('#customLoader').hide();
                    jq('#channelTableWrapper').show();
                    alert('Failed to load data: ' + (xhr.responseJSON?.message ||
                        'Unknown error. Check console for details.'));
                },
                ajaxRequesting: function() {
                    jq('#customLoader').show();
                    jq('#channelTableWrapper').hide();
                },
                columns: [{
                        title: "Channel",
                        field: "channel",
                        headerSort: true,
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (!value) return "";
                            const channelName = value.trim().toLowerCase();
                            const routeMap = {
                                'amazon': '/overall-amazon',
                                'amazon fba': '/overall-amazon-fba',
                                'ebay': '/ebay',
                                'ebaytwo': '/ebayTwoAnalysis',
                                'ebaythree': '/ebayThreeAnalysis',
                                'temu': '/temu',
                                'macys': '/macys',
                                'wayfair': '/Wayfair',
                                'reverb': '/reverb',
                                'shopify b2c': '/shopifyB2C',
                                'doba': '/doba',
                                'walmart': '/walmartAnalysis',
                                'bestbuy usa': '/bestbuyusa-analytics',
                            };
                            const routeUrl = routeMap[channelName];
                            return routeUrl ?
                                `<a href="${routeUrl}" target="_blank" style="color: #007bff; text-decoration: underline;">${value}</a>` :
                                `<span style="background:#f5f5f5;color:#000;font-weight:600;padding:4px 14px;border-radius:12px;">${value}</span>`;
                        },
                        hozAlign: "left"
                    },
                    {
                        title: `<small id="l30OrdersCountBadgeHeader" class="badge bg-dark text-white mb-1" style="font-size: 13px;">0</small><br>L30 Orders`,
                        field: "l30_orders",
                        formatter: function(cell) {
                            const value = parseFloat(cell.getValue()) || 0;
                            return `<span class="metric-value">${value.toLocaleString('en-US')}</span>`;
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "NR",
                        field: "nr",
                        formatter: function(cell) {
                            const value = toNum(cell.getValue());
                            const checked = value === 1 ? 'checked' : '';
                            return `<input type="checkbox" class="checkbox-nr" data-channel="${cell.getData().channel}" ${checked}>`;
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "ODR Rate",
                        field: "odr_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed > 5 ? "#dc3545" : parsed > 2 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['odr_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "ODR Rate Allowed",
                        field: "odr_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['odr_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Fulfillment Rate",
                        field: "fulfillment_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed < 95 ? "#dc3545" : parsed < 98 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['fulfillment_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Fulfillment Rate Allowed",
                        field: "fulfillment_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['fulfillment_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Valid Tracking Rate",
                        field: "valid_tracking_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed < 90 ? "#dc3545" : parsed < 95 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['valid_tracking_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Valid Tracking Rate Allowed",
                        field: "valid_tracking_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['valid_tracking_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "On Time Delivery",
                        field: "on_time_delivery",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed < 90 ? "#dc3545" : parsed < 95 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['on_time_delivery']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "On Time Delivery Allowed",
                        field: "on_time_delivery_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['on_time_delivery_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "A-Z Claims",
                        field: "atoz_claims_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed > 5 ? "#dc3545" : parsed > 2 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['atoz_claims_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "A-Z Claims Allowed",
                        field: "atoz_claims_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['atoz_claims_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Violation/Compliance",
                        field: "violation_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed > 5 ? "#dc3545" : parsed > 2 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['violation_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Violation/Compliance Allowed",
                        field: "violation_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['violation_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Late Shipment Rate",
                        field: "late_shipment_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed > 10 ? "#dc3545" : parsed > 5 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['late_shipment_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Late Shipment Rate Allowed",
                        field: "late_shipment_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['late_shipment_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Negative Seller Rate",
                        field: "negative_seller_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed > 3 ? "#dc3545" : parsed > 1 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['negative_seller_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Negative Seller Rate Allowed",
                        field: "negative_seller_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['negative_seller_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Refund Rate",
                        field: "refund_rate",
                        editor: function(cell, onRendered, success, cancel) {
                            const input = document.createElement("input");
                            input.type = "number";
                            input.min = 0;
                            input.max = 100;
                            input.step = 0.1;

                            const currentValue = cell.getValue();
                            let numericValue = currentValue;
                            if (typeof currentValue === 'string' && currentValue.includes('%')) {
                                numericValue = parseFloat(currentValue.replace('%', ''));
                            }
                            input.value = isNaN(numericValue) ? '' : numericValue;

                            onRendered(function() {
                                input.focus();
                                input.style.height = "100%";
                                input.style.width = "100%";
                                input.style.border = "none";
                                input.style.padding = "4px";
                                input.style.boxSizing = "border-box";
                            });

                            function onSubmit() {
                                let value = parseFloat(input.value);
                                if (isNaN(value) || value < 0 || value > 100) {
                                    cancel();
                                    return;
                                }
                                success(value);
                            }

                            input.addEventListener("blur", onSubmit);
                            input.addEventListener("keydown", function(e) {
                                if (e.key === "Enter") {
                                    onSubmit();
                                }
                                if (e.key === "Escape") {
                                    cancel();
                                }
                            });

                            return input;
                        },
                        formatter: function(cell) {
                            const value = cell.getValue();
                            if (value === 'N/A' || value === '' || isNaN(parseFloat(value))) {
                                return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                            <span>N/A</span>
                                        </div>`;
                            }
                            const parsed = parseFloat(value);
                            let color = parsed > 2 ? "#dc3545" : parsed > 1 ? "#ffc107" : "#28a745";
                            return `<div class="text-center editable-field" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                        <span>${parsed.toFixed(1)}%</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            const newValue = parseFloat(cell.getValue());
                            if (isNaN(newValue) || newValue < 0 || newValue > 100) {
                                alert('Please enter a value between 0 and 100');
                                cell.restoreOldValue();
                                return;
                            }
                            updateField(cell, fieldUpdateEndpoints['refund_rate']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Refund Rate Allowed",
                        field: "refund_rate_allowed",
                        editor: "input",
                        editorParams: {
                            minLength: 0,
                            maxLength: 50
                        },
                        formatter: function(cell) {
                            const value = cell.getValue() || '';
                            return `<div class="text-center editable-field" style="padding:4px; border-radius:4px;">
                                        <span>${value}</span>
                                    </div>`;
                        },
                        cellEdited: function(cell) {
                            updateField(cell, fieldUpdateEndpoints['refund_rate_allowed']);
                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Actions",
                        field: null,
                        formatter: function(cell) {
                            return `
                                <div class="d-flex justify-content-center">
                                    <button class="btn btn-sm btn-outline-primary edit-btn me-1" title="Edit" data-channel="${cell.getData().channel}">
                                        <i class="fas fa-edit"></i>
                                    </button>
                                    <button class="btn btn-sm btn-outline-danger delete-btn" title="Archive" data-channel="${cell.getData().channel}">
                                        <i class="fa fa-archive"></i>
                                    </button>
                                </div>`;
                        },
                        hozAlign: "center"
                    }
                ]
            });

            table.on("cellEdited", function(cell) {
                // Handled by column-specific cellEdited
            });

            return table;
        }

        function updateField(cell, updateUrl) {
            const rowData = cell.getData();
            const fieldName = cell.getField();
            const fieldValue = cell.getValue();

            // Create payload with specific field data
            const payload = {
                channel: rowData.channel,
                [fieldName]: fieldValue
            };

            console.log('Sending update payload:', payload);

            fetch(updateUrl, {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "X-CSRF-TOKEN": document.querySelector('meta[name="csrf-token"]').getAttribute("content")
                    },
                    body: JSON.stringify(payload)
                })
                .then(res => res.json())
                .then(response => {
                    if (response.success) {
                        cell.getElement().innerHTML += '<span class="save-indicator text-success">Saved</span>';
                        setTimeout(() => {
                            const indicator = cell.getElement().querySelector('.save-indicator');
                            if (indicator) indicator.remove();
                        }, 2000);
                        table.redraw();
                        console.log(`Updated ${fieldName}:`, response);
                    } else {
                        alert(`Error: ${response.message || 'Failed to save field'}`);
                        cell.restoreOldValue();
                    }
                })
                .catch(error => {
                    console.error(`Update failed for ${fieldName}:`, error);
                    alert('Server error while saving field');
                    cell.restoreOldValue();
                });
        }

        function updateL30OrdersTotal(data) {
            let l30OrdersTotal = 0;
            data.forEach(row => {
                const l30Orders = parseNumber(row['l30_orders'] || row['L30 Orders'] || 0);
                l30OrdersTotal += l30Orders;
            });
            const badgeElement = document.getElementById('l30OrdersCountBadgeHeader');
            if (badgeElement) {
                badgeElement.textContent = Math.round(l30OrdersTotal).toLocaleString('en-US');
            }
        }

        function drawChannelChart(data) {
            console.log('Drawing chart with data:', data);
            if (!data || data.length === 0) {
                console.warn('No data for chart');
                return;
            }
            google.charts.load('current', {
                packages: ['corechart']
            });
            google.charts.setOnLoadCallback(function() {
                renderChart(data);
            });
        }

        function renderChart(data) {
            console.log('Rendering chart with data:', data);
            let chartData = [
                ['Channel', 'L30 Orders', 'ODR Rate', 'ODR Rate Allowed', 'Fulfillment Rate',
                    'Fulfillment Rate Allowed', 'Valid Tracking Rate', 'Valid Tracking Rate Allowed',
                    'On Time Delivery', 'On Time Delivery Allowed', 'A-to-Z Claims', 'A-to-Z Claims Allowed',
                    'Violation/Compliance', 'Violation/Compliance Allowed', 'Late Shipment Rate',
                    'Late Shipment Rate Allowed', 'Negative Seller Rate', 'Negative Seller Rate Allowed',
                    'Refund Rate', 'Refund Rate Allowed', {
                        role: 'annotation'
                    }
                ]
            ];

            data.forEach(row => {
                let channel = row['channel'] || row['Channel'] || row['Channel '] || '';
                let l30Orders = parseFloat(row['l30_orders'] || row['L30 Orders'] || 0);
                let odrRate = parseNumber(row['odr_rate'] || row['ODR'] || row['ODR Rate'] || 'N/A');
                let odrRateAllowed = parseNumber(row['odr_rate_allowed'] || 0);
                let fulfillmentRate = parseNumber(row['fulfillment_rate'] || row['Fulfillment Rate'] || 'N/A');
                let fulfillmentRateAllowed = parseNumber(row['fulfillment_rate_allowed'] || 0);
                let validTrackingRate = parseNumber(row['valid_tracking_rate'] || row['Valid Tracking Rate'] ||
                    'N/A');
                let validTrackingRateAllowed = parseNumber(row['valid_tracking_rate_allowed'] || 0);
                let onTimeDelivery = parseNumber(row['on_time_delivery'] || row['On Time Delivery Rate'] || 'N/A');
                let onTimeDeliveryAllowed = parseNumber(row['on_time_delivery_allowed'] || 0);
                let atozClaimsRate = parseNumber(row['atoz_claims_rate'] || row['AtoZ Claims Rate'] || 'N/A');
                let atozClaimsRateAllowed = parseNumber(row['atoz_claims_rate_allowed'] || 0);
                let violationRate = parseNumber(row['violation_rate'] || row['Violation Rate'] || 'N/A');
                let violationRateAllowed = parseNumber(row['violation_rate_allowed'] || 0);
                let lateShipmentRate = parseNumber(row['late_shipment_rate'] || row['Late Shipment Rate'] || 'N/A');
                let lateShipmentRateAllowed = parseNumber(row['late_shipment_rate_allowed'] || 0);
                let negativeSellerRate = parseNumber(row['negative_seller_rate'] || row['Negative Seller Rate'] ||
                    'N/A');
                let negativeSellerRateAllowed = parseNumber(row['negative_seller_rate_allowed'] || 0);
                let refundRate = parseNumber(row['refund_rate'] || row['Refund Rate'] || 'N/A');
                let refundRateAllowed = parseNumber(row['refund_rate_allowed'] || 0);

                chartData.push([
                    channel,
                    l30Orders,
                    odrRate,
                    odrRateAllowed,
                    fulfillmentRate,
                    fulfillmentRateAllowed,
                    validTrackingRate,
                    validTrackingRateAllowed,
                    onTimeDelivery,
                    onTimeDeliveryAllowed,
                    atozClaimsRate,
                    atozClaimsRateAllowed,
                    violationRate,
                    violationRateAllowed,
                    lateShipmentRate,
                    lateShipmentRateAllowed,
                    negativeSellerRate,
                    negativeSellerRateAllowed,
                    refundRate,
                    refundRateAllowed,
                    channel
                ]);
            });

            let dataTable = google.visualization.arrayToDataTable(chartData);
            let options = {
                title: 'Channel Performance Metrics',
                curveType: 'function',
                legend: {
                    position: 'bottom'
                },
                hAxis: {
                    textPosition: 'none'
                },
                vAxis: {
                    title: 'Value',
                    minValue: 0
                },
                pointSize: 5,
                annotations: {
                    alwaysOutside: true,
                    textStyle: {
                        fontSize: 11,
                        bold: true,
                        color: '#000'
                    }
                },
                series: {
                    0: {
                        color: '#1E88E5',
                        lineWidth: 4
                    }, // L30 Orders
                    1: {
                        color: '#E53935',
                        lineWidth: 2
                    }, // ODR Rate
                    2: {
                        color: '#E57373',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // ODR Rate Allowed
                    3: {
                        color: '#43A047',
                        lineWidth: 2
                    }, // Fulfillment Rate
                    4: {
                        color: '#81C784',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // Fulfillment Rate Allowed
                    5: {
                        color: '#F06292',
                        lineWidth: 2
                    }, // Valid Tracking Rate
                    6: {
                        color: '#F8BBD0',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // Valid Tracking Rate Allowed
                    7: {
                        color: '#FFB300',
                        lineWidth: 2
                    }, // On Time Delivery
                    8: {
                        color: '#FFCA28',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // On Time Delivery Allowed
                    9: {
                        color: '#8E24AA',
                        lineWidth: 2
                    }, // A-to-Z Claims
                    10: {
                        color: '#BA68C8',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // A-to-Z Claims Allowed
                    11: {
                        color: '#26A69A',
                        lineWidth: 2
                    }, // Violation/Compliance
                    12: {
                        color: '#4DB6AC',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // Violation/Compliance Allowed
                    13: {
                        color: '#D81B60',
                        lineWidth: 2
                    }, // Late Shipment Rate
                    14: {
                        color: '#F06292',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // Late Shipment Rate Allowed
                    15: {
                        color: '#0288D1',
                        lineWidth: 2
                    }, // Negative Seller Rate
                    16: {
                        color: '#4FC3F7',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    }, // Negative Seller Rate Allowed
                    17: {
                        color: '#7CB342',
                        lineWidth: 2
                    }, // Refund Rate
                    18: {
                        color: '#AED581',
                        lineWidth: 2,
                        lineDashStyle: [4, 4]
                    } // Refund Rate Allowed
                }
            };
            let chart = new google.visualization.LineChart(document.getElementById('channelSalesChart'));
            chart.draw(dataTable, options);
        }

        function startPlayback() {
            if (!originalChannelData.length) return;
            uniqueChannels = [...new Set(originalChannelData.map(item => item['channel']?.trim() || item['Channel ']
            ?.trim() || item['Channel']?.trim()))].filter(Boolean);
            uniqueChannelRows = uniqueChannels.map(channel => originalChannelData.find(item => (item['channel']?.trim() ||
                item['Channel ']?.trim() || item['Channel']?.trim()) === channel));
            if (!uniqueChannelRows.length) return;

            currentChannelIndex = 0;
            isPlaying = true;
            table.setPageSize(1);
            table.clearFilter();
            showCurrentChannel();

            jq('#play-auto').hide();
            jq('#play-pause').show();
            setTimeout(() => updatePlayButtonColor(), 500);
        }

        function stopPlayback() {
            isPlaying = false;
            table.setData(originalChannelData);
            table.setPageSize(50);
            jq('#play-auto').show();
            jq('#play-pause').hide();
        }

        function showCurrentChannel() {
            if (!isPlaying || !uniqueChannelRows.length) return;
            const currentRow = uniqueChannelRows[currentChannelIndex];
            if (currentRow) {
                table.setData([currentRow]);
                jq('#channelSearchInput').val(currentRow['channel']?.trim() || currentRow['Channel ']?.trim() || currentRow[
                    'Channel']?.trim() || '');
                const tableContainer = jq('#channelTable').parent();
                if (tableContainer.length) tableContainer.scrollTop(0);
                setTimeout(() => updatePlayButtonColor(), 500);
            }
        }

        function nextChannel() {
            if (!isPlaying) return;
            if (currentChannelIndex < uniqueChannelRows.length - 1) {
                currentChannelIndex++;
                showCurrentChannel();
            } else {
                stopPlayback();
            }
        }

        function previousChannel() {
            if (!isPlaying) return;
            if (currentChannelIndex > 0) {
                currentChannelIndex--;
                showCurrentChannel();
            }
        }

        function updatePlayButtonColor() {
            const rows = table.getRows();
            if (!rows.length) return;
            const row = rows[0];
            const checkbox = jq(row.getElement()).find('.checkbox-nr');
            if (checkbox.length) {
                const isChecked = checkbox.prop('checked');
                jq('#play-pause').removeClass('btn-light btn-success btn-danger')
                    .addClass(isChecked ? 'btn-success' : 'btn-danger')
                    .css('color', 'white');
            }
        }

        function populateChannelDropdown(searchTerm = '') {
            const channelData = originalChannelData.map(row => row['channel'] || row['Channel '] || row['Channel']);
            const uniqueChannels = [...new Set(channelData)].filter(ch => ch && ch.trim() !== '').sort();
            const lowerSearch = searchTerm.toLowerCase();
            const sortedChannels = uniqueChannels.sort((a, b) => {
                const aMatch = a.toLowerCase().includes(lowerSearch);
                const bMatch = b.toLowerCase().includes(lowerSearch);
                if (aMatch && !bMatch) return -1;
                if (!aMatch && bMatch) return 1;
                return a.localeCompare(b);
            });

            const dropdown = jq('#channelSearchDropdown');
            dropdown.empty();
            sortedChannels.forEach(channel => {
                dropdown.append(`<div class="dropdown-search-item" data-value="${channel}">${channel}</div>`);
            });
            dropdown.show();
        }

        jq(document).ready(function() {
            jq('#customLoader').show();
            jq('#channelTableWrapper').hide();

            table = initializeTable();
            if (!table) return;

            jq.ajax({
                url: '/account-health-master/dashboard-data',
                type: "GET",
                success: function(json) {
                    console.log('Initial AJAX response:', json);
                    if (json && json.data) {
                        originalChannelData = json.data;
                        setupEventHandlers();
                    } else {
                        console.warn('No data in initial AJAX response:', json);
                        alert('No data received from server.');
                    }
                    jq('#customLoader').hide();
                    jq('#channelTableWrapper').show();
                },
                error: function(xhr, error, thrown) {
                    console.error('Error loading initial data:', error, thrown, xhr.responseText);
                    jq('#customLoader').hide();
                    jq('#channelTableWrapper').show();
                    alert('Failed to load initial data: ' + (xhr.responseJSON?.message ||
                        'Unknown error. Check console for details.'));
                }
            });
        });

        function setupEventHandlers() {
            jq('#play-auto').on('click', startPlayback);
            jq('#play-pause').on('click', stopPlayback);
            jq('#play-forward').on('click', nextChannel);
            jq('#play-backward').on('click', previousChannel);

            const channelSearchInput = jq('#channelSearchInput');
            const channelSearchDropdown = jq('#channelSearchDropdown');

            channelSearchInput.on('focus', () => populateChannelDropdown());
            channelSearchInput.on('input', function() {
                const val = this.value.trim();
                if (val === '') {
                    table.clearFilter();
                } else {
                    table.setFilter("channel", "=", val);
                }
                populateChannelDropdown(val);
            });

            channelSearchDropdown.on('click', '.dropdown-search-item', function() {
                const selectedChannel = jq(this).data('value');
                channelSearchInput.val(selectedChannel);
                channelSearchDropdown.hide();
                table.setFilter("channel", "=", selectedChannel);
            });

            jq(document).on('click', function(e) {
                if (!e.target.closest('.dropdown-search-container')) {
                    channelSearchDropdown.hide();
                }
            });

            table.on("dataLoaded", function(data) {
                updateL30OrdersTotal(data);
            });

            setupModalHandlers();
            setupCheckboxHandlers();

            table.on("cellClick", function(e, cell) {
                if (jq(e.target).closest('.edit-btn').length) {
                    const rowData = cell.getData();
                    const channel = rowData['channel']?.trim() || '';
                    const sheetUrl = rowData['sheet_link'] || '';
                    const type = rowData['type']?.trim() || '';

                    jq('#editChannelName').val(channel);
                    jq('#editChannelUrl').val(sheetUrl);
                    jq('#editType').val(type);
                    jq('#originalChannel').val(channel);
                    jq('#editChannelModal').modal('show');
                }

                if (jq(e.target).closest('.delete-btn').length) {
                    const channel = cell.getData().channel;
                    if (confirm(`Are you sure you want to archive channel "${channel}"?`)) {
                        jq.ajax({
                            url: '/channel_master/archive',
                            method: 'POST',
                            data: {
                                channel: channel,
                                _token: jq('meta[name="csrf-token"]').attr('content')
                            },
                            success: function(res) {
                                if (res.success) {
                                    table.deleteRow(cell.getRow());
                                    alert('Channel archived successfully.');
                                } else {
                                    alert('Error: ' + (res.message || 'Failed to archive channel.'));
                                }
                            },
                            error: function() {
                                alert('Server error while archiving channel.');
                            }
                        });
                    }
                }
            });
        }

        function setupModalHandlers() {
            jq('#saveChannelBtn').on('click', function() {
                const channelName = jq('#channelName').val().trim();
                const channelUrl = jq('#channelUrl').val().trim();
                const type = jq('#type').val().trim();

                if (!channelName || !channelUrl || !type) {
                    alert("Channel Name, Sheet Link, and Type are required.");
                    return;
                }

                jq.ajax({
                    url: '/channel_master/store',
                    method: 'POST',
                    data: {
                        channel: channelName,
                        sheet_link: channelUrl,
                        type: type,
                        _token: jq('meta[name="csrf-token"]').attr('content')
                    },
                    success: function(res) {
                        if (res.success) {
                            jq('#addChannelModal').modal('hide');
                            jq('#channelForm')[0].reset();
                            table.setData('/account-health-master/dashboard-data');
                            alert('Channel added successfully.');
                        } else {
                            alert("Error: " + (res.message || 'Failed to add channel.'));
                        }
                    },
                    error: function() {
                        alert("Server error while adding channel.");
                    }
                });
            });

            jq('#editChannelForm').on('submit', function(e) {
                e.preventDefault();
                const channel = jq('#editChannelName').val().trim();
                const sheetUrl = jq('#editChannelUrl').val().trim();
                const type = jq('#editType').val().trim();
                const originalChannel = jq('#originalChannel').val().trim();

                if (!channel || !sheetUrl || !type) {
                    alert("Channel Name, Sheet URL, and Type are required.");
                    return;
                }

                jq.ajax({
                    url: '/channel_master/update',
                    method: 'POST',
                    data: {
                        channel: channel,
                        sheet_url: sheetUrl,
                        type: type,
                        original_channel: originalChannel,
                        _token: jq('meta[name="csrf-token"]').attr('content')
                    },
                    success: function(res) {
                        if (res.success) {
                            jq('#editChannelModal').modal('hide');
                            jq('#editChannelForm')[0].reset();
                            table.setData('/account-health-master/dashboard-data');
                            alert('Channel updated successfully.');
                        } else {
                            alert("Error: " + (res.message || 'Failed to update channel.'));
                        }
                    },
                    error: function() {
                        alert("Server error while updating channel.");
                    }
                });
            });

            const importForm = jq('#accountHealthImportForm');
            const progressBar = jq('#importProgress');
            const resultsDiv = jq('#importResults');
            const submitBtn = jq('#importSubmitBtn');

            importForm.on('submit', function(e) {
                e.preventDefault();
                const formData = new FormData(this);
                progressBar.show();
                resultsDiv.hide();
                submitBtn.prop('disabled', true).html('<i class="fas fa-spinner fa-spin me-1"></i> Importing...');

                let progress = 0;
                const progressInterval = setInterval(() => {
                    progress += 10;
                    progressBar.find('.progress-bar').css('width', progress + '%').text(progress + '%');
                    if (progress >= 90) clearInterval(progressInterval);
                }, 200);

                fetch(this.action, {
                        method: 'POST',
                        body: formData,
                        headers: {
                            'X-CSRF-TOKEN': jq('meta[name="csrf-token"]').attr('content')
                        }
                    })
                    .then(response => response.json())
                    .then(data => {
                        clearInterval(progressInterval);
                        progressBar.find('.progress-bar').css('width', '100%').text('100%');
                        setTimeout(() => {
                            progressBar.hide();
                            if (data.success) {
                                resultsDiv.removeClass('alert-danger').addClass('alert-success');
                                const resultsList = jq('#importResultsList').empty();
                                if (data.results) {
                                    Object.keys(data.results).forEach(key => {
                                        resultsList.append(
                                            `<li>${key}: ${data.results[key]}</li>`);
                                    });
                                }
                                resultsDiv.show();
                                table.setData('/account-health-master/dashboard-data');
                                setTimeout(() => {
                                    jq('#accountHealthImportModal').modal('hide');
                                    location.reload();
                                }, 3000);
                            } else {
                                resultsDiv.removeClass('alert-success').addClass('alert-danger');
                                resultsDiv.html(
                                    `<strong>Error:</strong> ${data.message || 'Import failed'}`);
                                resultsDiv.show();
                            }
                        }, 500);
                    })
                    .catch(error => {
                        clearInterval(progressInterval);
                        progressBar.hide();
                        resultsDiv.removeClass('alert-success').addClass('alert-danger');
                        resultsDiv.html('<strong>Error:</strong> Network error occurred during import');
                        resultsDiv.show();
                        console.error('Import error:', error);
                    })
                    .finally(() => {
                        submitBtn.prop('disabled', false).html(
                        '<i class="fas fa-file-import me-1"></i> Import');
                    });
            });

            jq('#accountHealthImportModal').on('hidden.bs.modal', function() {
                importForm[0].reset();
                progressBar.hide();
                resultsDiv.hide();
                submitBtn.prop('disabled', false).html('<i class="fas fa-file-import me-1"></i> Import');
            });
        }

        function setupCheckboxHandlers() {
            jq('#channelTable').on('change', '.checkbox-nr', function() {
                const channel = jq(this).data('channel');
                const value = jq(this).is(':checked') ? 1 : 0;

                jq.ajax({
                    url: '/channels-master/toggle-flag',
                    method: 'POST',
                    data: {
                        channel: channel,
                        field: 'nr',
                        value: value,
                        _token: jq('meta[name="csrf-token"]').attr('content')
                    },
                    success: function(res) {
                        if (!res.success) {
                            alert('Failed to update NR flag: ' + (res.message || 'Unknown error'));
                            jq(this).prop('checked', !value);
                        }
                    },
                    error: function() {
                        alert('Server error while updating NR flag.');
                        jq(this).prop('checked', !value);
                    }
                });
            });
        }
    </script>
@endsection
