@extends('layouts.vertical', ['title' => 'MFRG In Progress'])
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
    @include('layouts.shared.page-title', ['page_title' => 'MFRG In Progress', 'sub_title' => 'MFRG In Progress'])

    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body">

                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h4 class="mb-0">MFRG In Progress</h4>
                    </div>

                    <div class="row mb-4 g-3 align-items-end justify-content-between">
                        {{-- ▶️ Navigation --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">▶️ Navigation</label>
                            <div class="btn-group time-navigation-group" role="group">
                                <button id="play-backward" class="btn btn-light rounded-circle shadow-sm me-2"
                                    title="Previous parent">
                                    <i class="fas fa-step-backward"></i>
                                </button>
                                <button id="play-pause" class="btn btn-light rounded-circle shadow-sm me-2"
                                    style="display: none;" title="Pause">
                                    <i class="fas fa-pause"></i>
                                </button>
                                <button id="play-auto" class="btn btn-primary rounded-circle shadow-sm me-2" title="Play">
                                    <i class="fas fa-play"></i>
                                </button>
                                <button id="play-forward" class="btn btn-light rounded-circle shadow-sm"
                                    title="Next parent">
                                    <i class="fas fa-step-forward"></i>
                                </button>
                            </div>
                        </div>

                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">Pending Status</label>
                            <select id="row-data-pending-status" class="form-select border border-primary" style="width: 150px;">
                                <option value="">select color</option>
                                <option value="green">Green <span id="greenCount"></span></option>
                                <option value="yellow">Yellow <span id="yellowCount"></span></option>
                                <option value="red">Red <span id="redCount"></span></option>
                            </select>
                        </div>

                        {{-- total amount --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">💰 Total Amount</label>
                            <div id="totalAmount" class="fw-bold text-primary" style="font-size: 1.1rem;">
                                00
                            </div>
                        </div>

                        {{-- Ordered Items --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">Total Order Qty</label>
                            <div id="pendingItemsCount" class="fw-bold text-primary" style="font-size: 1.1rem;">
                                00
                            </div>
                        </div>

                        {{-- 📦 Total CBM --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">📦 Total CBM</label>
                            <div id="totalCBM" class="fw-bold text-success" style="font-size: 1.1rem;">
                                00
                            </div>
                        </div>

                        {{-- 🔍 Search --}}
                        <div class="col-auto">
                            <label for="search-input" class="form-label fw-semibold mb-1 d-block">🔍 Search</label>
                            <input type="text" id="search-input" class="form-control form-control-sm" placeholder="Search anything...">
                        </div>

                        {{-- 🗑️ Delete Selected --}}
                        <div class="col-auto">
                            <button class="btn btn-sm btn-danger d-none" id="delete-selected-btn">
                                <i class="fas fa-trash-alt me-1"></i> Delete
                            </button>
                        </div>
                    </div>
                    <div id="mfrg-table"></div>
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
            document.body.style.zoom = "80%";

            document.documentElement.setAttribute("data-sidenav-size", "condensed");

            const globalPreview = Object.assign(document.createElement("div"), {
                id: "image-hover-preview",
            });

            Object.assign(globalPreview.style, {
                position: "fixed",
                zIndex: 9999,
                border: "1px solid #ccc",
                background: "#fff",
                padding: "4px",
                boxShadow: "0 2px 8px rgba(0,0,0,0.2)",
                display: "none",
            });
            document.body.appendChild(globalPreview);

            let hideTimeout;
            let uniqueSuppliers = [];

            const table = new Tabulator("#mfrg-table", {
                ajaxURL: "/mfrg-in-progress/data",
                ajaxConfig: "GET",
                layout: "fitData",
                height: "700px",
                pagination: true,
                paginationSize: 100,
                paginationCounter: "rows",
                movableColumns: false,
                resizableColumns: true,
                columns: [
                    {
                        formatter: "rowSelection",
                        titleFormatter: "rowSelection",
                        hozAlign: "center",
                        width: 50,
                        headerSort: false
                    },
                    {
                        title: "#",
                        field: "Image",
                        headerSort: false,
                        formatter: (cell) => {
                            const url = cell.getValue();
                            return url ?
                                `<img src="${url}" data-full="${url}" class="hover-thumb" 
                                   style="width:30px;height:30px;border-radius:6px;object-fit:contain;
                                   box-shadow:0 1px 4px #0001;cursor: pointer;">` :
                                `<span class="text-muted">N/A</span>`;
                        },
                        cellMouseOver: (e, cell) => {
                            clearTimeout(hideTimeout);

                            const img = cell.getElement().querySelector(".hover-thumb");
                            if (!img) return;

                            globalPreview.innerHTML = `<img src="${img.dataset.full}" style="max-width:350px;max-height:350px;">`;
                            globalPreview.style.display = "block";
                            globalPreview.style.top = `${e.clientY + 15}px`;
                            globalPreview.style.left = `${e.clientX + 15}px`;
                        },
                        cellMouseMove: (e) => {
                            globalPreview.style.top = `${e.clientY + 15}px`;
                            globalPreview.style.left = `${e.clientX + 15}px`;
                        },
                        cellMouseOut: () => {
                            hideTimeout = setTimeout(() => {
                                globalPreview.style.display = "none";
                            }, 150);
                        },
                    },
                    {
                        title: "Parent",
                        field: "parent",
                        headerFilter: "input",
                        headerFilterPlaceholder: " Filter parent...",
                        width: 180,
                        headerFilterLiveFilter: true,
                    },
                    {
                        title: "SKU",
                        field: "sku", 
                        headerFilter: "input",
                        width: 180,
                        headerFilterPlaceholder: " Filter SKU...",
                        headerFilterLiveFilter: true,
                    },
                    {
                        title: "Order Qty",
                        field: "qty",
                        hozAlign: "center",
                        formatter: function (cell) {
                            const value = cell.getValue() || "";
                            
                            const html = `
                                    <div style="display:flex; justify-content:center; align-items:center; width:100%;">
                                        <input type="number" 
                                            class="form-control form-control-sm qty-input" 
                                            value="${value}" 
                                            min="0" max="99999" 
                                            style="width:80px; text-align:center;">
                                    </div>
                                `;

                            setTimeout(() => {
                                const input = cell.getElement().querySelector(".qty-input");
                                if (input) {
                                    input.addEventListener("change", function () {
                                        const newValue = this.value;
                                        saveLinkUpdate(cell, newValue);
                                    });
                                }
                            }, 10);

                            return html;
                        }
                    },
                    { 
                        title: "Rate", 
                        field: "rate", 
                        formatter: function(cell) {
                            const row = cell.getRow().getData();
                            const sku = row.sku || '';
                            const currency = row.rate_currency || 'USD';
                            const rate = row.rate || '';

                            return `
                                <div class="input-group input-group-sm" style="width:105px;">
                                    <span class="input-group-text" style="padding: 0 6px;">
                                        <select data-sku="${sku}" data-column="rate_currency" 
                                            class="form-select form-select-sm currency-select auto-save" 
                                            style="border: none; background: transparent; font-size: 13px; padding: 0 2px;">
                                            <option value="USD" ${currency === 'USD' ? 'selected' : ''}>$</option>
                                            <option value="CNY" ${currency === 'CNY' ? 'selected' : ''}>¥</option>
                                        </select>
                                    </span>
                                    <input data-sku="${sku}" data-column="rate" type="number" value="${rate}" 
                                        class="form-control form-control-sm amount-input auto-save" 
                                        style="background: #f9f9f9; font-size: 13px;" />
                                </div>
                            `;
                        },
                        hozAlign: "center",
                        headerHozAlign: "center",
                    },
                    {
                        title: "Supplier",
                        field: "supplier",
                        hozAlign: "center",
                        formatter: function(cell){
                            let value = cell.getValue() || "";
                            let options = uniqueSuppliers.map(supplier => {
                                let selected = (supplier === value) ? "selected" : "";
                                return `<option value="${supplier}" ${selected}>${supplier}</option>`;
                            }).join("");

                            return `
                                <select class="form-select form-select-sm editable-select" 
                                    data-sku="${cell.getRow().getData().SKU}" data-column="Supplier" style="width: 120px;">
                                    ${options}
                                </select>`;
                        }
                    },
                    {
                        title: "Order Date",
                        field: "created_at",
                        hozAlign: "center",
                        formatter: function (cell) {
                            const rawValue = cell.getValue() || "";
                            const formattedDate = rawValue ? new Date(rawValue).toISOString().split('T')[0] : "";
                            const rowData = cell.getRow().getData();

                            const html = `
                                <div style="display: flex; flex-direction: column; align-items: flex-start;">
                                    <input type="date" class="form-control form-control-sm order_date_input" value="${formattedDate}" style="width:85px;">
                                </div>
                            `;

                            // setTimeout(() => {
                            //     const input = cell.getElement().querySelector(".order_date_input");
                            //     if (input) {
                            //         input.addEventListener("change", function () {
                            //             const newValue = this.value;
                            //             saveLinkUpdate(cell, newValue);
                            //         });
                            //     }
                            // }, 10);

                            return html;
                        }
                    },
                    {
                        title: "Del Date",
                        field: "del_date",
                        hozAlign: "center",
                        formatter: function (cell) {
                            const rawValue = cell.getValue() || "";
                            const formattedDate = rawValue ? new Date(rawValue).toISOString().split('T')[0] : "";
                            const rowData = cell.getRow().getData();

                            const html = `
                                <div style="display: flex; flex-direction: column; align-items: flex-start;">
                                    <input type="date" class="form-control form-control-sm del_date_input" value="${formattedDate}" style="width:82px;">
                                </div>
                            `;

                            // setTimeout(() => {
                            //     const input = cell.getElement().querySelector(".del_date_input");
                            //     if (input) {
                            //         input.addEventListener("change", function () {
                            //             const newValue = this.value;
                            //             saveLinkUpdate(cell, newValue);
                            //         });
                            //     }
                            // }, 10);

                            return html;
                        }
                    },
                    {
                        title: "CBM",
                        field: "CBM",
                        hozAlign: "center",
                        formatter: function(cell){
                            const cellValue = cell.getValue();
                            const value = cellValue ? Number(cellValue).toFixed(4) : '0.0000';
                            return value;
                        }
                    },
                    {
                        title: "R2S",
                        field: "ready_to_ship",
                        hozAlign: "center",
                        formatter: function(cell) {
                            const value = cell.getValue() || "";
                            const rowData = cell.getRow().getData();

                            return `
                                <select class="form-select form-select-sm editable-select"
                                    data-column="ready_to_ship"
                                    data-sku='${rowData["SKU"]}'
                                    style="width: 75px;">
                                    <option value="No" ${value === "No" ? "selected" : ""}>No</option>
                                    <option value="Yes" ${value === "Yes" ? "selected" : ""}>Yes</option>
                                </select>
                            `;
                        },
                    },
                ],
                ajaxResponse: (url, params, response) => {
                    let data = response.data;

                    let filtered = data.filter(item => {
                        let qty = parseFloat(item.qty) || 0;
                        let isParent = item.sku && item.sku.startsWith("PARENT");
                        let isReadyToShip = item.ready_to_ship && item.ready_to_ship.trim().toLowerCase() === "yes";

                        return qty > 0 && !isParent && !isReadyToShip;
                    });

                    uniqueSuppliers = [...new Set(filtered.map(item => item.supplier))].filter(Boolean);
                    return filtered;
                },
            });

            table.on("rowSelectionChanged", function(data, rows) {
                if (data.length > 0) {
                    $('#delete-selected-btn').removeClass('d-none');
                } else {
                    $('#delete-selected-btn').addClass('d-none');
                }
            });
        });
    </script>
@endsection