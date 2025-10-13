@extends('layouts.vertical', ['title' => 'A-Z Claims'])

@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
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

        .custom-select-wrapper.disabled {
            cursor: not-allowed;
        }

        .custom-select-wrapper.disabled .custom-select-display {
            background-color: #f8f9fa;
            color: #6c757d;
            cursor: not-allowed;
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
        'page_title' => 'A-Z Claims',
        'sub_title' => 'Account Health Master',
    ])
    @if (Session::has('flash_message'))
        <div class="alert alert-primary bg-primary text-white alert-dismissible fade show" role="alert"
            style="background-color: #169e28 !important; color: #fff !important;">
            {{ Session::get('flash_message') }}
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="alert" aria-label="Close"></button>
        </div>
    @endif

    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body">
                    <div class="d-flex flex-wrap justify-content-between align-items-center mb-4 gap-1">
                        <div class="d-flex align-items-center gap-2">
                            <select id="channelFilter" class="form-select form-control-lg me-2"
                                style="width: 200px; min-width: 160px;">
                                <option value="">All Channels</option>
                                @foreach ($channels as $channel)
                                    <option value="{{ $channel->channel }}">{{ $channel->channel }}</option>
                                @endforeach
                            </select>
                        </div>
                        <!-- <div class="d-flex align-items-center gap-2">
                            <div class="form-check form-switch">
                                <input class="form-check-input" type="checkbox" id="toggleAction" checked>
                                <label class="form-check-label" for="toggleAction">Show Actions</label>
                            </div>
                            <button type="button" class="btn btn-success d-flex align-items-center d-none" data-bs-toggle="modal"
                                data-bs-target="#accountHealthModal">
                                <i class="fas fa-plus-circle me-1"></i> Add Report
                            </button>
                        </div> -->
                    </div>

                    <div id="odr-rate"></div>
                </div>
            </div>
        </div>
    </div>

    {{-- odr rate modal form --}}
    <div class="modal fade" id="accountHealthModal" tabindex="-1" aria-hidden="true">
        <div class="modal-dialog modal-lg modal-dialog-centered shadow-none" role="document">
            <div class="modal-content border border-primary rounded-2">
                <div class="modal-header bg-white border-bottom rounded-top-2">
                    <h5 class="modal-title fw-semibold text-primary" id="modal-title">
                        <i class="fas fa-heartbeat me-2"></i> ODR Rate Report
                    </h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>

                <div class="modal-body bg-light">
                    <form id="edit-form">
                        @csrf
                        <input type="hidden" name="id" id="edit_id">
                        <div class="row g-4">

                            {{-- Marketplace Channel --}}
                            <div class="col-md-4">
                                <label class="form-label fw-semibold">Marketplace Channel <span
                                        class="text-danger">*</span></label>
                                <div class="custom-select-wrapper position-relative">
                                    <div class="custom-select-display form-select">Choose a channel...</div>
                                    <div class="custom-select-options position-absolute bg-white border rounded shadow-sm mt-1 w-100"
                                        style="display: none; z-index: 1000;">
                                        <input type="text" class="custom-select-search form-control mb-2"
                                            placeholder="Search..." />
                                        @foreach ($channels as $channel)
                                            <div class="custom-select-option px-3 py-2 border-bottom cursor-pointer"
                                                data-value="{{ $channel->id }}">
                                                {{ $channel->channel }}
                                            </div>
                                        @endforeach
                                    </div>
                                    <input type="hidden" name="channel_id" id="channel_id_hidden" required>
                                </div>
                                @error('channel')
                                    <div class="text-danger mt-1 small">{{ $message }}</div>
                                @enderror
                            </div>

                            {{-- Report Date --}}
                            <div class="col-md-4">
                                <label class="form-label fw-semibold">Report Date <span class="text-danger">*</span></label>
                                <input type="date" class="form-control" name="report_date" id="report_date_input" value="{{ date('Y-m-d') }}"
                                    required>
                            </div>

                            {{-- Account Health Links --}}
                            <div class="col-md-4">
                                <label class="form-label fw-semibold">Account Health Links</label>
                                <input type="text" class="form-control" name="account_health_links" id="account_health_links_input"
                                    placeholder="Enter Account Health link">
                            </div>
                        </div>

                        <div class="row g-4 mt-3">
                            <div class="col-md-6">
                                <label class="form-label fw-semibold">What</label>
                                <textarea class="form-control" name="what" id="what_input" rows="2"></textarea>
                            </div>
                            <div class="col-md-6">
                                <label class="form-label fw-semibold">Why</label>
                                <textarea class="form-control" name="why" id="why_input" rows="2"></textarea>
                            </div>
                            <div class="col-md-6">
                                <label class="form-label fw-semibold">Action</label>
                                <input type="text" class="form-control" name="action" id="action_input">
                            </div>
                            <div class="col-md-6">
                                <label class="form-label fw-semibold">C+ Action</label>
                                <input type="text" class="form-control" name="c_action" id="c_action_input">
                            </div>
                        </div>

                        <div class="mt-4 d-flex justify-content-end">
                            <button type="submit" class="btn btn-primary px-4" id="submit-btn">
                                <i class="fas fa-save me-1"></i> Save Report
                            </button>
                        </div>
                    </form>
                </div>
            </div>
        </div>
    </div>
@endsection
@section('script')
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <script>
        document.addEventListener('DOMContentLoaded', function() {

            let tableData = []; // To store the claims data
            let table; // To reference the table
            const editModal = new bootstrap.Modal(document.getElementById('accountHealthModal'));
            let isEditMode = false; // Flag to track if in edit mode

            // Fetch data manually to handle both table and eBay VTR
            fetch("/fetchAtoZClaimsRates", {
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                    "X-Requested-With": "XMLHttpRequest"
                }
            })
            .then(response => response.json())
            .then(data => {
                tableData = data.atoz_claims;

                // Update eBay row with VTR data
                if (data.ebay_vtr && data.ebay_vtr.success) {
                    const ebayRow = tableData.find(row => row.channel === 'eBay');
                    if (ebayRow) {
                        ebayRow.allowed = data.ebay_vtr.thresholdLower;
                        ebayRow.current = data.ebay_vtr.vtr;
                    }
                }

                initializeTable(tableData);
            })
            .catch(error => {
                console.error("Error fetching data:", error);
            });

            function initializeTable(data) {
                table = new Tabulator("#odr-rate", {
                    data: data, // Use the fetched data
                    layout: "fitDataFill",
                    pagination: true,
                    paginationSize: 50,
                    paginationMode: "local",
                    movableColumns: false,
                    resizableColumns: true,
                    height: "550px",
                    columns: [{
                            title: "Channels",
                            field: "channel",
                            headerSort: true,
                            formatter: function(cell) {
                                const value = cell.getValue();
                                if (!value) return "";
                                let color = "#f5f5f5";
                                return `<span style="display:inline-flex;align-items:center;gap:7px;">
                    
                            <span style="background:${color};color:#000;font-weight:600;padding:4px 14px;border-radius:12px;font-size:1em;letter-spacing:0.01em;">
                                ${value}
                            </span>
                            </span>`;
                            }
                        },
                        {
                            title: "Allowed",
                            field: "allowed",
                            formatter: function(cell) {
                                const value = cell.getValue();
                                let percent = parseFloat(value);
                                if (isNaN(percent)) return '';
                                return `<div class="text-center" style="background-color:transparent; padding:4px; border-radius:4px;">
                                    <span>${percent}%</span>
                                </div>`;
                            },
                            hozAlign: "center"
                        },
                        {
                            title: "Current",
                            field: "current",
                            formatter: function(cell) {
                                const value = cell.getValue();
                                let percent = parseFloat(value);
                                if (isNaN(percent)) return '';
                                let color = "transparent";
                                if (percent <= 95) {
                                    color = "#dc3545"; // red
                                } else if (percent >= 96 && percent <= 97) {
                                    color = "#ffc107"; // orange
                                } else if (percent >= 98) {
                                    color = "#28a745"; // green
                                }
                                return `<div class="text-center" style="background-color:${color}; color:#ffffff; padding:4px; border-radius:4px;">
                                    <span>${percent}%</span>
                                </div>`;
                            },
                            hozAlign: "center"
                        },
                        {
                            title: "Action",
                            field: "action",
                            editor: false,
                            formatter: function(cell){
                                const data = cell.getRow().getData();
                                if (data.channel === 'eBay') {
                                    return `<span style="cursor: pointer; color: #0d6efd; text-decoration: underline;" title="Open eBay Dashboard">
                                        View Dashboard
                                    </span>`;
                                }
                                const value = cell.getValue() || '';
                                return `-`;
                            },
                            hozAlign: "center",
                            cellClick: function(e, cell){
                                const rowData = cell.getRow().getData();
                                if (rowData.channel === 'eBay') {
                                    window.open('https://sellerstandards.ebay.com/dashboard?region=US', '_blank');
                                    return;
                                }
                                openEditModal(cell.getRow());
                            }
                        },
                    ],
                });

                // Action column toggle
                const toggleAction = document.getElementById('toggleAction');
                const actionColumn = table.getColumn('action');
                toggleAction.addEventListener('change', function() {
                    if (this.checked) {
                        actionColumn.show();
                    } else {
                        actionColumn.hide();
                    }
                    table.redraw(true); // Redraw to adjust column widths
                });

                // Channel filter functionality
                const channelFilter = document.getElementById("channelFilter");
                channelFilter.addEventListener("change", function () {
                    const selectedChannel = this.value;
                    if (selectedChannel) {
                        table.setFilter("channel", "=", selectedChannel);
                    } else {
                        table.clearFilter();
                    }
                });
            }

            // Modal form submit handler
            const form = document.getElementById('edit-form');
            form.addEventListener('submit', function(e){
                e.preventDefault();
                const formData = new FormData(form);
                const updateData = {
                    id: formData.get('id'),
                    channel_id: formData.get('channel_id'),
                    report_date: formData.get('report_date'),
                    account_health_links: formData.get('account_health_links'),
                    what: formData.get('what'),
                    why: formData.get('why'),
                    action: formData.get('action'),
                    c_action: formData.get('c_action'),
                };

                fetch("/AtoZClaims-rate/update", {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                            "X-CSRF-TOKEN": document.querySelector('meta[name="csrf-token"]').getAttribute("content")
                        },
                        body: JSON.stringify(updateData)
                    })
                    .then(res => res.json())
                    .then(response => {
                        console.log("Updated:", response);
                        // Update table row
                        const rowId = updateData.id;
                        table.updateData([updateData]); // Update the row in table
                        editModal.hide();
                    })
                    .catch(error => console.error("Update failed:", error));
            });

            function openEditModal(row){
                const data = row.getData();
                isEditMode = true;

                // Set title
                document.getElementById('modal-title').innerHTML = `<i class="fas fa-edit me-2"></i> Edit Report Details for ${data.channel}`;

                // Set channel - pre-select and disable
                const wrapper = document.querySelector(".custom-select-wrapper");
                const display = wrapper.querySelector(".custom-select-display");
                display.textContent = data.channel;
                document.getElementById('channel_id_hidden').value = data.channel_id || ''; // Assuming channel_id is in data
                wrapper.classList.add('disabled');
                const optionsContainer = wrapper.querySelector(".custom-select-options");
                optionsContainer.style.display = 'none'; // Ensure options are closed

                // Set other fields
                document.getElementById('report_date_input').value = data.report_date || '';
                document.getElementById('account_health_links_input').value = data.account_health_links || '';
                document.getElementById('what_input').value = data.what || '';
                document.getElementById('why_input').value = data.why || '';
                document.getElementById('action_input').value = data.action || '';
                document.getElementById('c_action_input').value = data.c_action || '';

                // Set hidden id
                document.getElementById('edit_id').value = data.id || '';

                // Set button text
                document.getElementById('submit-btn').innerHTML = '<i class="fas fa-save me-1"></i> Update Report';

                editModal.show();
            }

            // Reset modal on hide
            editModal._element.addEventListener('hidden.bs.modal', function () {
                const wrapper = document.querySelector(".custom-select-wrapper");
                wrapper.classList.remove('disabled');
                const display = wrapper.querySelector(".custom-select-display");
                display.textContent = 'Choose a channel...';
                document.getElementById('channel_id_hidden').value = '';
                isEditMode = false;
                // Reset other fields if needed
            });

            channelSelect();
        });

        function channelSelect() {
            const wrapper = document.querySelector(".custom-select-wrapper");
            const display = wrapper.querySelector(".custom-select-display");
            const optionsContainer = wrapper.querySelector(".custom-select-options");
            const searchInput = wrapper.querySelector(".custom-select-search");
            const hiddenInput = wrapper.querySelector("input[name='channel_id']");
            const options = wrapper.querySelectorAll(".custom-select-option");

            display.addEventListener("click", (e) => {
                if (wrapper.classList.contains('disabled')) {
                    return; // Prevent opening if disabled
                }
                e.stopPropagation();
                const isOpen = optionsContainer.style.display === "block";
                optionsContainer.style.display = isOpen ? "none" : "block";
                searchInput.value = "";
                filterOptions("");
                if (!isOpen) searchInput.focus();
            });

            options.forEach(option => {
                option.addEventListener("click", (e) => {
                    e.stopPropagation();
                    const value = option.dataset.value;         
                    const label = option.textContent.trim();    
                    display.textContent = label;                
                    hiddenInput.value = value;                 
                    optionsContainer.style.display = "none";
                });

            });


            searchInput.addEventListener("input", (e) => {
                filterOptions(e.target.value);
            });

            function filterOptions(query) {
                options.forEach(opt => {
                    const text = opt.textContent.toLowerCase();
                    opt.style.display = text.includes(query.toLowerCase()) ? "block" : "none";
                });
            }

            document.addEventListener("click", function(e) {
                if (!wrapper.contains(e.target)) {
                    optionsContainer.style.display = "none";
                }
            });
        }

    </script>
@endsection