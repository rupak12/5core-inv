@extends('layouts.vertical', ['title' => $form->name . ' Form Reports'])
@section('css')
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
<link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
<style>
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

    #img-popup {
        pointer-events: none;
    }

    .thumb-img {
        height: 30px;
        width: 30px;
        margin: 2px;
        border-radius: 4px;
        cursor: pointer;
        transition: all 0.2s ease-in-out;
        position: relative;
    }

    .thumb-img:hover::after {
        content: '';
        position: fixed;
        top: 50%;
        left: 50%;
        width: auto;
        height: auto;
        max-width: 80vw;
        max-height: 80vh;
        transform: translate(-50%, -50%);
        background: url(attr(src)) no-repeat center/contain;
        border: 2px solid #ccc;
        border-radius: 8px;
        z-index: 9999;
    }
</style>
@endsection
@section('content')
@include('layouts.shared.page-title', ['page_title' => $form->name . ' Form Reports', 'sub_title' => $form->name . ' Form Reports'])

@if (Session::has('flash_message'))
    <div class="alert alert-primary bg-primary text-white alert-dismissible fade show" role="alert"
        style="background-color: #03a744 !important; color: #fff !important;">
        {{ Session::get('flash_message') }}
        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="alert" aria-label="Close"></button>
    </div>
@endif

<div class="row">
    <div class="col-12">
        <div class="card shadow-sm">
            <div class="card-body">
                <div class="d-flex flex-wrap justify-content-end align-items-center mb-3 gap-2">
                    <div class="d-flex flex-wrap gap-2">
                        <button id="export-btn" class="btn btn-sm btn-success">
                            <i class="fas fa-file-excel"></i> Export Excel
                        </button>
                    </div>
                </div>
                <div id="form-reports-table"></div>
            </div>
        </div>
    </div>
</div>

@endsection
@section('script')
<script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
<!-- SheetJS for Excel Export -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
    document.addEventListener('DOMContentLoaded', function () {

        const table = new Tabulator("#form-reports-table", {
            ajaxURL: "/rfq-form/reports-data/{{ $form->id }}",
            ajaxConfig: "GET",
            layout: "fitData",
            pagination: true,
            paginationSize: 50,
            paginationMode: "local",
            movableColumns: false,
            resizableColumns: true,
            height: "500px",
            columns: [],
            ajaxResponse: function(url, params, response){
                if(!response.data || response.data.length === 0) return [];

                const sampleRow = response.data[0].data;

                const priorityKeys = ['additionalPhotos','supplierName','companyName','supplierLink','productName'];

                const dynamicColumns = [];

                // S.No column
                dynamicColumns.push({
                    title: "#",
                    formatter: "rownum",
                    hozAlign: "center",
                    width: 35
                });

                // Priority columns first
                priorityKeys.forEach(key => {
                    if(sampleRow.hasOwnProperty(key)){
                        if(key === 'supplierLink'){
                            dynamicColumns.push({
                                title: 'Supplier <i class="fa-solid fa-link"></i>',
                                field: key,
                                formatter: function(cell){
                                    const url = cell.getValue();
                                    if(!url) return "-";
                                    return `<a href="${url}" target="_blank" class="btn btn-sm btn-outline-info"><i class="fa-solid fa-link"></i></a>`;
                                },
                                hozAlign: "center"
                            });
                        } else if(key === 'additionalPhotos'){
                            dynamicColumns.push({
                                title: "Photos",
                                field: key,
                                formatter: function(cell){
                                    const photos = cell.getValue();
                                    if(!photos || photos.length === 0) return "-";

                                    return photos.map((img, index) => `
                                        <img src="/storage/${img}" class="thumb-img" data-img-id="${index}" style="width:30px; height:30px; margin:2px; border-radius:4px; cursor:pointer;">
                                    `).join('');
                                },
                                hozAlign: "center"
                            });
                        } else {
                            dynamicColumns.push({
                                title: key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase()),
                                field: key,
                                hozAlign: "left",
                            });
                        }
                    }
                });

                // Remaining dynamic columns
                Object.keys(sampleRow).forEach(key => {
                    if(!priorityKeys.includes(key)){
                        dynamicColumns.push({
                            title: key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase()),
                            field: key,
                            hozAlign: "left",
                        });
                    }
                });

                table.setColumns(dynamicColumns);

                return response.data.map(row => row.data);
            },
        });


        let popup;

        document.addEventListener('mouseover', function(e){
            if(e.target.classList.contains('thumb-img')){
                const src = e.target.src;

                if(popup) popup.remove();

                popup = document.createElement('div');
                popup.id = 'img-popup';
                popup.style.position = 'fixed';
                popup.style.top = '50%';
                popup.style.left = '50%';
                popup.style.transform = 'translate(-50%, -50%)';
                popup.style.zIndex = 9999;
                popup.style.padding = '10px';
                popup.style.background = '#fff';
                popup.style.border = '2px solid #ccc';
                popup.style.borderRadius = '8px';
                popup.innerHTML = `<img src="${src}" style="max-width:80vw; max-height:80vh;">`;

                document.body.appendChild(popup);
            }
        });

        document.addEventListener('mouseout', function(e){
            if(e.target.classList.contains('thumb-img')){
                if(popup) {
                    popup.remove();
                    popup = null;
                }
            }
        });

        document.getElementById("export-btn").addEventListener("click", function () {
            let allData = table.getData("active"); 

            if (allData.length === 0) {
                alert("No data available to export!");
                return;
            }

            let exportData = allData.map(row => ({ ...row }));
            let formName = "{{ $form->name }}";

            let ws = XLSX.utils.json_to_sheet(exportData);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Campaigns");

            XLSX.writeFile(wb, `${formName}_form_report.xlsx`);
        });


    });
</script>

@endsection