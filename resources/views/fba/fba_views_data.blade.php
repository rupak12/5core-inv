@extends('layouts.vertical', ['title' => 'FBA Sales Data', 'sidenav' => 'condensed'])

@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" rel="stylesheet">
    <link rel="stylesheet" href="{{ asset('assets/css/styles.css') }}">
@endsection
@section('script')
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
@endsection

@section('content')
    <div class="container-fluid">
        <div class="row">
            <div class="col-12">
                <div class="card">
                    <div class="card-header">
                        <h4>FBA Data </h4>

                    </div>
                    <div class="card-body">
                        <div id="fba-table"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>
@endsection

@section('script-bottom')
    <script>
        $(document).ready(function() {
            const table = new Tabulator("#fba-table", {
                ajaxURL: "/fba-data-json",
                layout: "fitData",
                pagination: true,
                paginationSize: 50,
                rowFormatter: function(row) {
                    if (row.getData().is_parent) {
                        row.getElement().classList.add("parent-row");
                    }
                },
                columns: [{
                        title: "Parent",
                        field: "Parent",
                        headerFilter: "input",
                        headerFilterPlaceholder: "Search Parent...",
                        cssClass: "text-primary",
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
                        frozen: true
                    },
                    {
                        title: "FBA SKU",
                        field: "FBA_SKU"
                    },
                  
                    {
                        title: "Shopify INV",
                        field: "Shopify_INV",
                        hozAlign: "center"
                    },
                      {
                        title: "FBA INV",
                        field: "FBA_Quantity",
                        hozAlign: "center"
                    },
                    {
                        title: "FBA Price",
                        field: "FBA_Price",
                        hozAlign: "center",
                        formatter: "dollar"
                    },
                    {
                        title: "L30 Units",
                        field: "l30_units",
                        hozAlign: "center"
                    },
                   
                    {
                        title: "L60 Units",
                        field: "l60_units",
                        hozAlign: "center"
                    },
               
                    {
                        title: "Views",
                        field: "Current_Month_Views",
                        hozAlign: "center"
                    },
                    {
                        title: "FBA Fee",
                        field: "Fulfillment_Fee",
                        hozAlign: "center"
                    },

                    {
                        title: "ASIN",
                        field: "ASIN"
                    },
                    {
                        
                        title: "Jan",
                        field: "Jan",
                        hozAlign: "center"
                    },
                    {
                        title: "Feb",
                        field: "Feb",
                        hozAlign: "center"
                    },
                    {
                        title: "Mar",
                        field: "Mar",
                        hozAlign: "center"
                    },
                    {
                        title: "Apr",
                        field: "Apr",
                        hozAlign: "center"
                    },
                    {
                        title: "May",
                        field: "May",
                        hozAlign: "center"
                    },
                    {
                        title: "Jun",
                        field: "Jun",
                        hozAlign: "center"
                    },
                    {
                        title: "Jul",
                        field: "Jul",
                        hozAlign: "center"
                    },
                    {
                        title: "Aug",
                        field: "Aug",
                        hozAlign: "center"
                    },
                    {
                        title: "Sep",
                        field: "Sep",
                        hozAlign: "center"
                    },
                    {
                        title: "Oct",
                        field: "Oct",
                        hozAlign: "center"
                    },
                    {
                        title: "Nov",
                        field: "Nov",
                        hozAlign: "center"
                    },
                    {
                        title: "Dec",
                        field: "Dec",
                        hozAlign: "center"
                    }
                ]
            });
        });
    </script>
@endsection
