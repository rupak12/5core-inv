@extends('layouts.vertical', ['title' => 'Purchase'])
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

        /* #image-hover-preview {
                transition: opacity 0.2s ease;
            } */
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
    @include('layouts.shared.page-title', ['page_title' => 'Purchase', 'sub_title' => 'Purchase'])

    <div class="row">
        <div class="col-12">
            <div class="card shadow-sm">
                <div class="card-body">

                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h4 class="mb-0">To Order Analysis</h4>
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
                            <label class="form-label fw-semibold mb-1 d-block">Parent / Sku</label>
                            <select id="row-data-type" class="form-select border border-primary" style="width: 150px;">
                                <option value="all">🔁 Show All</option>
                                <option value="sku">🔹 SKU (Child)</option>
                                <option value="parent">🔸 Parent</option>
                            </select>
                        </div>

                        {{-- 🕒 Pending Items --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">🕒 Pending Items</label>
                            <div class="fw-bold text-primary" style="font-size: 1.1rem;">
                                00
                            </div>
                        </div>

                        <div class="col-auto" hidden>
                            <label class="form-label fw-semibold mb-1 d-block">Total Approved Qty</label>
                            <div class="fw-bold text-primary" style="font-size: 1.1rem;">
                                00
                            </div>
                        </div>

                        {{-- 📦 Total CBM --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">📦 Total CBM</label>
                            <div class="fw-bold text-success" style="font-size: 1.1rem;">
                                00
                            </div>
                        </div>

                        {{-- 🎯 Stage Filter --}}
                        <div class="col-auto">
                            <label class="form-label fw-semibold mb-1 d-block">🎯 Stage</label>
                            <select id="stage-filter" class="form-select form-select-sm" style="min-width: 160px;">
                                <option value="">All Stages</option>
                                <option value="rfq sent">RFQ Sent</option>
                                <option value="analytics">Analytics</option>
                                <option value="to approve">To Approve</option>
                                <option value="approved">Approved</option>
                                <option value="advance">Advance</option>
                                <option value="mfrg progress">Mfrg Progress</option>
                            </select>
                        </div>

                        {{-- 🔍 Search --}}
                        <div class="col-auto">
                            <label for="search-input" class="form-label fw-semibold mb-1 d-block">🔍 Search</label>
                            <input type="text" id="search-input" class="form-control form-control-sm"
                                placeholder="Search suppliers...">
                        </div>
                    </div>
                    <div id="toOrderAnalysis-table"></div>
                </div>
            </div>
        </div>
    </div>
@endsection
@section('script')
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <script>
        document.addEventListener("DOMContentLoaded", function() {
            document.body.style.zoom = "75%";

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

            const table = new Tabulator("#toOrderAnalysis-table", {
                ajaxURL: "/to-order-analysis/data",
                ajaxConfig: "GET",
                layout: "fitData",
                height: "700px",
                pagination: true,
                paginationSize: 50,
                movableColumns: false,
                resizableColumns: true,
                columns: [
                    {
                        title: "#",
                        field: "Image",
                        headerSort: false,
                        formatter: (cell) => {
                            const url = cell.getValue();
                            return url ?
                                `<img src="${url}" data-full="${url}" 
                            class="hover-thumb" 
                            style="width:30px;height:30px;border-radius:6px;object-fit:contain;
                                   box-shadow:0 1px 4px #0001;cursor: pointer;">` :
                                `<span class="text-muted">N/A</span>`;
                        },
                        cellMouseOver: (e, cell) => {
                            clearTimeout(hideTimeout);

                            const img = cell.getElement().querySelector(".hover-thumb");
                            if (!img) return;

                            globalPreview.innerHTML = `<img src="${img.dataset.full}" 
                        style="max-width:350px;max-height:350px;">`;
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
                        field: "Parent"
                    },
                    {
                        title: "SKU",
                        field: "SKU"
                    },
                    {
                        title: "App. QTY",
                        field: "approved_qty",
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
                        title: "DOA",
                        field: "Date of Appr",
                        formatter: function (cell) {
                            const value = cell.getValue() || "";
                            const rowData = cell.getRow().getData();

                            let daysDiff = null;
                            let bgColor = "";

                            if (value) {
                                let doa = new Date(value);
                                let today = new Date();
                                let diffTime = today - doa;
                                daysDiff = Math.floor(diffTime / (1000 * 60 * 60 * 24));

                                if (daysDiff >= 14) {
                                    bgColor = "background-color:red; color:white;";
                                } else if (daysDiff >= 7) {
                                    bgColor = "background-color:yellow; color:black;";
                                }
                            }

                            const html = `
                                <div style="display: flex; flex-direction: column; align-items: flex-start;">
                                    <input type="date" class="form-control form-control-sm doa-input" value="${value}" style="width:82px; ${bgColor}">
                                </div>
                            `;

                            setTimeout(() => {
                                const input = cell.getElement().querySelector(".doa-input");
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
                        title: "Supplier",
                        field: "Supplier",
                        formatter: function(cell){
                            let value = cell.getValue() || "";
                            let options = uniqueSuppliers.map(supplier => {
                                let selected = (supplier === value) ? "selected" : "";
                                return `<option value="${supplier}" ${selected}>${supplier}</option>`;
                            }).join("");

                            return `
                                <select class="form-select form-select-sm stage-select" 
                                        data-sku="${cell.getRow().getData().SKU}" style="width: 120px;">
                                    ${options}
                                </select>`;
                        }
                    },
                    {
                        title: "Review",
                        field: "Review",
                        formatter: "html"
                    },
                    {
                        title: "RFQ Form",
                        field: "RFQ Form Link",
                        formatter: linkFormatter,
                        editor: "input",  
                        hozAlign: "center",
                        cellEdited: function(cell){
                            saveLinkUpdate(cell, cell.getValue());
                        }
                    },
                    {
                        title: "Rfq Report",
                        field: "Rfq Report Link",
                        formatter: linkFormatter,
                        editor: "input",         
                        hozAlign: "center",
                        cellEdited: function(cell){
                            saveLinkUpdate(cell, cell.getValue());
                        }
                    },
                    {
                        title: "Sheet",
                        field: "sheet_link",
                        formatter: "link",
                        formatterParams: {
                            target: "_blank"
                        },
                    },
                    {
                        title: "NRP",
                        field: "nrl",
                        formatter: function (cell) {
                            const row = cell.getRow();
                            const sku = row.getData().Sku;
                            return `
                                <select class="form-select form-select-sm editable-select" data-row-id="${sku}" data-type="NR"
                                    style="width: 75px;">
                                    <option value="REQ" ${cell.getValue() === 'REQ' ? 'selected' : ''}>RE</option>
                                    <option value="NR" ${cell.getValue() === 'NR' ? 'selected' : ''}>NR</option>
                                </select>
                            `;

                        },
                        hozAlign: "center"
                    },
                    {
                        title: "Stage",
                        field: "Stage",
                        formatter: function(cell) {
                            const value = cell.getValue() || "";
                            const rowData = cell.getRow().getData();

                            return `
                                <select class="form-select form-select-sm stage-select"
                                    data-type="stage"
                                    data-sku='${rowData["SKU"]}'
                                    style="width: 100px;>
                                    <option value="RFQ Sent" ${value === "RFQ Sent" ? "selected" : ""}>RFQ Sent</option>
                                    <option value="Analytics" ${value === "Analytics" ? "selected" : ""}>Analytics</option>
                                    <option value="To Approve" ${value === "To Approve" ? "selected" : ""}>To Approve</option>
                                    <option value="Approved" ${value === "Approved" ? "selected" : ""}>Approved</option>
                                    <option value="Advance" ${value === "Advance" ? "selected" : ""}>Advance</option>
                                    <option value="Mfrg Progress" ${value === "Mfrg Progress" ? "selected" : ""}>Mfrg Progress</option>
                                </select>
                            `;
                        },
                    },
                    {
                        title: "Adv date",
                        field: "Adv date"
                    },
                    {
                        title: "Order Qty",
                        field: "order_qty",
                        hozAlign: "center",
                        formatter: function (cell) {
                            const value = cell.getValue() || "";
                            
                            const html = `
                                    <div style="display:flex; justify-content:center; align-items:center; width:100%;">
                                        <input type="number" 
                                            class="form-control form-control-sm order_qty" 
                                            value="${value}" 
                                            min="0" max="99999" 
                                            style="width:80px; text-align:center;">
                                    </div>
                                `;

                            setTimeout(() => {
                                const input = cell.getElement().querySelector(".order_qty");
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
                ],
                ajaxResponse: (url, params, response) => {
                    let data = response.data;

                    let filtered = data.filter(item => {
                        let qty = parseFloat(item.approved_qty) || 0;
                        let isParent = item.SKU && item.SKU.startsWith("PARENT");
                        let isMfrg = item.Stage && item.Stage.trim().toLowerCase() === "mfrg progress";
                        let isNR = item.nrl && item.nrl.trim().toUpperCase() === "NR";

                        return qty > 0 && !isParent && !isMfrg && !isNR;
                    });

                    uniqueSuppliers = [...new Set(filtered.map(item => item.Supplier))].filter(Boolean);
                    return filtered;
                },
            });


            function linkFormatter(cell) {
                let url = cell.getValue() || "";
                if (url && url.trim() !== "") {
                    return `
                        <div style="align-items:center;">
                            <a href="${url}" target="_blank" class="btn btn-sm btn-outline-primary" 
                            title="Open Link">
                                <i class="mdi mdi-link"></i> Open
                            </a>
                        </div>
                    `;
                }
            }

            function saveLinkUpdate(cell, value) {
                let sku = cell.getRow().getData().SKU;
                let column = cell.getColumn().getField();

                fetch('/update-link', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-CSRF-TOKEN': document.querySelector('meta[name="csrf-token"]').content
                    },
                    body: JSON.stringify({
                        sku: sku,
                        column: column,
                        value: value
                    })
                })
                .then(res => res.json())
                .then(res => {
                    if (!res.success) {
                        alert("Error: " + res.message);
                    }
                })
                .catch(err => console.error(err));
            }

            globalPreview.addEventListener("mouseenter", () => clearTimeout(hideTimeout));
            globalPreview.addEventListener("mouseleave", () => {
                globalPreview.style.display = "none";
            });
        });
    </script>
@endsection
