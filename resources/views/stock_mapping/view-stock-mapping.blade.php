@extends('layouts.vertical', ['title' => 'View Stock Mapping', 'mode' => $mode ?? '', 'demo' => $demo ?? ''])
<meta name="csrf-token" content="{{ csrf_token() }}">

@section('css')
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
    <style>
        /* ========== TABLE STRUCTURE ========== */
        .table-container {
            overflow-x: auto;
            overflow-y: visible;
            position: relative;
            max-height: 600px;
        }

        .custom-resizable-table {
            width: 100%;
            border-collapse: collapse;
            margin: 0;
        }

        .custom-resizable-table th,
        .custom-resizable-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
            position: relative;
            white-space: nowrap;
            overflow: visible !important;
        }

        .custom-resizable-table th {
            background-color: #f8f9fa;
            font-weight: 600;
            user-select: none;
            position: sticky;
            top: 0;
            z-index: 10;
        }

        /* ========== RESIZABLE COLUMNS ========== */
        .resize-handle {
            position: absolute;
            top: 0;
            right: 0;
            width: 5px;
            height: 100%;
            background: rgba(0, 0, 0, 0.1);
            cursor: col-resize;
            z-index: 100;
        }

        .resize-handle:hover,
        .resize-handle.resizing {
            background: rgba(0, 0, 0, 0.3);
        }

        /* ========== TOOLTIP SYSTEM ========== */
        .tooltip-container {
            position: relative;
            display: inline-block;
            margin-left: 8px;
        }

        .tooltip-icon {
            cursor: pointer;
            transform: translateY(1px);
        }

        .tooltip {
            z-index: 9999 !important;
            pointer-events: none;
        }

        .tooltip-inner {
            transform: translate(-5px, -5px) !important;
            max-width: 300px;
            padding: 6px 10px;
            font-size: 13px;
        }

        .bs-tooltip-top .tooltip-arrow {
            bottom: 0;
        }

        .bs-tooltip-top .tooltip-arrow::before {
            transform: translateX(5px) !important;
            border-top-color: var(--bs-tooltip-bg);
        }

        /* ========== COLOR CODED CELLS ========== */
        .dil-percent-cell {
            padding: 8px 4px !important;
        }

        .dil-percent-value {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 4px;
            font-weight: bold;
        }

        .dil-percent-value.red {
            background-color: #dc3545;
            color: white;
        }

        .dil-percent-value.blue {
            background-color: #3591dc;
            color: white;
        }

        .dil-percent-value.yellow {
            background-color: #ffc107;
            color: #212529;
        }

        .dil-percent-value.green {
            background-color: #28a745;
            color: white;
        }

        .dil-percent-value.pink {
            background-color: #e83e8c;
            color: white;
        }

        .dil-percent-value.gray {
            background-color: #6c757d;
            color: white;
        }

        /* ========== TABLE CONTROLS ========== */
        .table-controls {
            position: sticky;
            bottom: 0;
            background: white;
            padding: 10px 0;
            border-top: 1px solid #ddd;
        }

        /* ========== SORTING ========== */
        .sortable {
            cursor: pointer;
        }

        .sortable:hover {
            background-color: #f1f1f1;
        }

        .sort-arrow {
            display: inline-block;
            margin-left: 5px;
        }

        /* ========== PARENT ROWS ========== */
        .parent-row {
            background-color: rgba(69, 233, 255, 0.1) !important;
        }

        /* ========== SKU TOOLTIPS ========== */
        .sku-tooltip-container {
            position: relative;
            display: inline-block;
        }

        .sku-tooltip {
            visibility: hidden;
            width: auto;
            min-width: 120px;
            background-color: #fff;
            color: #333;
            text-align: left;
            border-radius: 4px;
            padding: 8px;
            position: absolute;
            z-index: 1001;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            opacity: 0;
            transition: opacity 0.3s;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            border: 1px solid #ddd;
            white-space: nowrap;
        }

        .sku-tooltip-container:hover .sku-tooltip {
            visibility: visible;
            opacity: 1;
        }

        .sku-link {
            padding: 4px 0;
            white-space: nowrap;
        }

        .sku-link a {
            color: #0d6efd;
            text-decoration: none;
        }

        .sku-link a:hover {
            text-decoration: underline;
        }

        /* ========== DROPDOWNS ========== */
        .custom-dropdown {
            position: relative;
            display: inline-block;
        }

        .custom-dropdown-menu {
            display: none;
            position: absolute;
            background-color: white;
            min-width: 200px;
            box-shadow: 0px 8px 16px 0px rgba(0, 0, 0, 0.2);
            z-index: 1000;
            max-height: 300px;
            overflow-y: auto;
            border: 1px solid #ddd;
            border-radius: 4px;
        }

        .custom-dropdown-menu.show {
            display: block;
        }

        .column-toggle-item {
            padding: 8px 16px;
            cursor: pointer;
            display: flex;
            align-items: center;
        }

        .column-toggle-item:hover {
            background-color: #f8f9fa;
        }

        .column-toggle-checkbox {
            margin-right: 8px;
        }

        /* ========== LOADER ========== */
        .card-loader-overlay {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255, 255, 255, 0.8);
            z-index: 100;
            display: flex;
            justify-content: center;
            align-items: center;
            border-radius: 0.25rem;
        }

        .loader-content {
            text-align: center;
            padding: 20px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.1);
        }

        .loader-text {
            margin-top: 15px;
            font-weight: 500;
            color: #333;
        }

        .spinner-border {
            width: 3rem;
            height: 3rem;
        }

        /* ========== CARD BODY ========== */
        .card-body {
            position: relative;
        }

        /* ========== SEARCH DROPDOWNS ========== */
        .dropdown-search-container {
            position: relative;
        }

        .dropdown-search-results {
            position: absolute;
            width: 100%;
            max-height: 300px;
            overflow-y: auto;
            z-index: 1000;
            background: white;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            display: none;
        }

        .dropdown-search-item {
            padding: 8px 12px;
            cursor: pointer;
        }

        .dropdown-search-item:hover {
            background-color: #f8f9fa;
        }

        .no-results {
            color: #6c757d;
            font-style: italic;
        }

        /* ========== STATUS INDICATORS ========== */
        .status-circle {
            display: inline-block;
            width: 12px;
            height: 12px;
            border-radius: 50%;
            margin-right: 6px;
            vertical-align: middle;
            border: 1px solid #fff;
        }

        .status-circle.default {
            background-color: #6c757d;
        }

        .status-circle.red {
            background-color: #dc3545;
        }

        .status-circle.yellow {
            background-color: #ffc107;
        }

        .status-circle.blue {
            background-color: #007bff;
        }

        .status-circle.green {
            background-color: #28a745;
        }

        .status-circle.pink {
            background-color: #e83e8c;
        }

        /* ========== FILTER CONTROLS ========== */
        .d-flex.flex-wrap.gap-2 {
            gap: 0.5rem !important;
            margin-bottom: 1rem;
        }

        .btn-sm i.fas {
            margin-right: 5px;
        }

        .manual-dropdown-container {
            position: relative;
            display: inline-block;
        }

        .manual-dropdown-container .dropdown-menu {
            display: none;
            position: absolute;
            top: 100%;
            left: 0;
            z-index: 1000;
            min-width: 160px;
            padding: 5px 0;
            margin: 2px 0 0;
            background-color: #fff;
            border: 1px solid rgba(0, 0, 0, .15);
            border-radius: 4px;
            box-shadow: 0 6px 12px rgba(0, 0, 0, .175);
        }

        .manual-dropdown-container.show .dropdown-menu {
            display: block;
        }

        .dropdown-item {
            display: block;
            width: 100%;
            padding: 8px 16px;
            clear: both;
            font-weight: 400;
            color: #212529;
            text-align: inherit;
            white-space: nowrap;
            background-color: transparent;
            border: 0;
        }

        .dropdown-item:hover {
            color: #16181b;
            text-decoration: none;
            background-color: #f8f9fa;
        }

        /* ========== MODAL SYSTEM ========== */
        .custom-modal {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            z-index: 1050;
            overflow: hidden;
            outline: 0;
            pointer-events: none;
        }

        .custom-modal.show {
            display: block;
        }

        .custom-modal-dialog {
            position: fixed;
            width: auto;
            min-width: 600px;
            max-width: 90vw;
            margin: 1.75rem auto;
            pointer-events: auto;
            z-index: 1051;
            transition: transform 0.3s ease-out;
            background-color: white;
            border-radius: 0.3rem;
            box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15);
        }

        .custom-modal-content {
            pointer-events: auto;
        }

        .custom-modal-header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            padding: 1rem;
            border-bottom: 1px solid #dee2e6;
            border-top-left-radius: 0.3rem;
            border-top-right-radius: 0.3rem;
            background-color: #f8f9fa;
        }

        .custom-modal-title {
            margin-bottom: 0;
            line-height: 1.5;
            font-size: 1.25rem;
        }

        .custom-modal-close {
            padding: 0;
            background-color: transparent;
            border: 0;
            font-size: 1.5rem;
            font-weight: 700;
            line-height: 1;
            color: #000;
            text-shadow: 0 1px 0 #fff;
            opacity: 0.5;
            cursor: pointer;
        }

        .custom-modal-close:hover {
            opacity: 0.75;
        }

        .custom-modal-body {
            position: relative;
            flex: 1 1 auto;
            padding: 1rem;
            overflow-y: auto;
            max-height: 70vh;
        }

        /* Multiple Modal Stacking */
        .custom-modal:nth-child(1) .custom-modal-dialog {
            top: 20px;
            right: 20px;
            z-index: 1051;
        }

        .custom-modal:nth-child(2) .custom-modal-dialog {
            top: 40px;
            right: 40px;
            z-index: 1052;
        }

        .custom-modal:nth-child(3) .custom-modal-dialog {
            top: 60px;
            right: 60px;
            z-index: 1053;
        }

        .custom-modal:nth-child(4) .custom-modal-dialog {
            top: 80px;
            right: 80px;
            z-index: 1054;
        }

        .custom-modal:nth-child(5) .custom-modal-dialog {
            top: 100px;
            right: 100px;
            z-index: 1055;
        }

        /* For more than 5 modals - dynamic calculation */
        .custom-modal:nth-child(n+6) .custom-modal-dialog {
            top: calc(100px + (var(--modal-offset) * 20px));
            right: calc(100px + (var(--modal-offset) * 20px));
            z-index: calc(1055 + var(--modal-offset));
        }

        /* Animations */
        @keyframes modalSlideIn {
            from {
                transform: translateX(30px);
                opacity: 0;
            }

            to {
                transform: translateX(0);
                opacity: 1;
            }
        }

        @keyframes modalFadeIn {
            from {
                opacity: 0;
            }

            to {
                opacity: 1;
            }
        }

        .custom-modal.show .custom-modal-dialog {
            animation: modalSlideIn 0.3s ease-out;
        }

        .custom-modal-backdrop.show {
            display: block;
            animation: modalFadeIn 0.15s linear;
        }

        /* Body scroll lock */
        body.custom-modal-open {
            overflow: hidden;
            padding-right: 15px;
        }

        /* Responsive adjustments */
        @media (max-width: 768px) {
            .custom-modal-dialog {
                min-width: 95vw;
                max-width: 95vw;
                margin: 0.5rem auto;
            }

            .custom-modal:nth-child(1) .custom-modal-dialog,
            .custom-modal:nth-child(2) .custom-modal-dialog,
            .custom-modal:nth-child(3) .custom-modal-dialog,
            .custom-modal:nth-child(4) .custom-modal-dialog,
            .custom-modal:nth-child(5) .custom-modal-dialog,
            .custom-modal:nth-child(n+6) .custom-modal-dialog {
                top: 10px;
                right: 10px;
                left: 10px;
                margin: 0 auto;
            }
        }

        /* Status color overlays */
        .custom-modal .card.card-bg-red {
            background: linear-gradient(135deg, rgba(245, 0, 20, 0.69), rgba(255, 255, 255, 0.85));
            border-color: rgba(220, 53, 70, 0.72);
        }

        .custom-modal .card.card-bg-green {
            background: linear-gradient(135deg, rgba(3, 255, 62, 0.424), rgba(255, 255, 255, 0.85));
            border-color: rgba(40, 167, 69, 0.3);
        }

        .custom-modal .card.card-bg-yellow {
            background: linear-gradient(135deg, rgba(255, 193, 7, 0.15), rgba(255, 255, 255, 0.85));
            border-color: rgba(255, 193, 7, 0.3);
        }

        .custom-modal .card.card-bg-blue {
            background: linear-gradient(135deg, rgba(0, 123, 255, 0.15), rgba(255, 255, 255, 0.85));
            border-color: rgba(0, 123, 255, 0.3);
        }

        .custom-modal .card.card-bg-pink {
            background: linear-gradient(135deg, rgba(232, 62, 140, 0.15), rgba(255, 255, 255, 0.85));
            border-color: rgba(232, 62, 141, 0.424);
        }

        .custom-modal .card.card-bg-gray {
            background: linear-gradient(135deg, rgba(108, 117, 125, 0.15), rgba(255, 255, 255, 0.85));
            border-color: rgba(108, 117, 125, 0.3);
        }

        @keyframes slideInRight {
            from {
                transform: translateX(100%);
                opacity: 0;
            }

            to {
                transform: translateX(0);
                opacity: 1;
            }
        }

        .custom-modal.show .custom-modal-dialog {
            animation: slideInRight 0.3s ease-out;
        }

        /* Close All button */
        #close-all-modals {
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 1060;
        }

        .custom-modal-dialog {
            position: fixed !important;
            top: 20px;
            right: 20px;
            margin: 0 !important;
            transform: none !important;
            cursor: move;
        }

        .custom-modal-header {
            cursor: move;
        }


        /* ========== PLAY/PAUSE NAVIGATION BUTTONS ========== */
        .time-navigation-group {
            margin-left: 10px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
            border-radius: 50px;
            overflow: hidden;
            padding: 2px;
            background: #f8f9fa;
            display: inline-flex;
            align-items: center;
        }

        .time-navigation-group button {
            padding: 0;
            border-radius: 50% !important;
            width: 40px;
            height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 3px;
            transition: all 0.2s ease;
            border: 1px solid #dee2e6;
            background: white;
            cursor: pointer;
        }

        .time-navigation-group button:hover {
            background-color: #f1f3f5 !important;
            transform: scale(1.05);
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }

        .time-navigation-group button:active {
            transform: scale(0.95);
        }

        .time-navigation-group button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none !important;
            box-shadow: none !important;
        }

        .time-navigation-group button i {
            font-size: 1.1rem;
            transition: transform 0.2s ease;
        }

        /* Play button */
        #play-auto {
            color: #28a745;
        }

        #play-auto:hover {
            background-color: #28a745 !important;
            color: white !important;
        }

        /* Pause button */
        #play-pause {
            color: #ffc107;
            display: none;
        }

        #play-pause:hover {
            background-color: #ffc107 !important;
            color: white !important;
        }

        /* Navigation buttons */
        #play-backward,
        #play-forward {
            color: #007bff;
        }

        #play-backward:hover,
        #play-forward:hover {
            background-color: #007bff !important;
            color: white !important;
        }

        /* Button state colors - must come after hover styles */
        #play-auto.btn-success,
        #play-pause.btn-success {
            background-color: #28a745 !important;
            color: white !important;
        }

        #play-auto.btn-warning,
        #play-pause.btn-warning {
            background-color: #ffc107 !important;
            color: #212529 !important;
        }

        #play-auto.btn-danger,
        #play-pause.btn-danger {
            background-color: #dc3545 !important;
            color: white !important;
        }

        #play-auto.btn-light,
        #play-pause.btn-light {
            background-color: #f8f9fa !important;
            color: #212529 !important;
        }

        /* Ensure hover doesn't override state colors */
        #play-auto.btn-success:hover,
        #play-pause.btn-success:hover {
            background-color: #28a745 !important;
            color: white !important;
        }

        #play-auto.btn-warning:hover,
        #play-pause.btn-warning:hover {
            background-color: #ffc107 !important;
            color: #212529 !important;
        }

        #play-auto.btn-danger:hover,
        #play-pause.btn-danger:hover {
            background-color: #dc3545 !important;
            color: white !important;
        }

        /* Active state styling */
        .time-navigation-group button:focus {
            outline: none;
            box-shadow: 0 0 0 3px rgba(0, 123, 255, 0.25);
        }

        /* Responsive adjustments */
        @media (max-width: 768px) {
            .time-navigation-group button {
                width: 36px;
                height: 36px;
            }

            .time-navigation-group button i {
                font-size: 1rem;
            }
        }

        /* Add to your CSS file or style section */
        .hide-column {
            display: none !important;
        }

        /*popup modal style*/

        .choose-file {
            background-color: #ff6b2c;
            color: white;
            padding: 10px;
            border-radius: 8px;
            text-align: center;
            cursor: pointer;
            width: 100%;
            display: block;
            transition: background-color 0.3s;
        }

        .choose-file:hover {
            background-color: #e65c1e;
        }

        .modal-content {
            border-radius: 16px;
            padding: 25px;
            box-shadow: 0 0 30px rgba(0, 0, 0, 0.1);
        }

        .form-label {
            font-weight: 600;
        }

        .form-section {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 15px;
            margin-bottom: 15px;
        }

        option[value="Todo"] {
            background-color: #2196f3;
        }

        option[value="Not Started"] {
            background-color: #ffff00;
            color: #000;
        }

        option[value="Working"] {
            background-color: #ff00ff;
        }

        option[value="In Progress"] {
            background-color: #f1c40f;
            color: #000;
        }

        option[value="Monitor"] {
            background-color: #5c6bc0;
        }

        option[value="Done"] {
            background-color: #00ff00;
            color: #000;
        }

        option[value="Need Help"] {
            background-color: #e91e63;
        }

        option[value="Review"] {
            background-color: #ffffff;
            color: #000;
        }

        option[value="Need Approval"] {
            background-color: #d4ff00;
            color: #000;
        }

        option[value="Dependent"] {
            background-color: #ff9999;
        }

        option[value="Approved"] {
            background-color: #ffeb3b;
            color: #000;
        }

        option[value="Hold"] {
            background-color: #ffffff;
            color: #000;
        }

        option[value="Rework"] {
            background-color: #673ab7;
        }

        option[value="Urgent"] {
            background-color: #f44336;
        }

        option[value="Q-Task"] {
            background-color: #ff00ff;
        }


        /*popup modal style end */

        

        #ebay-table thead tr#summaryRow th {
            position: sticky;
            top: 0;
            background: #f8f9fa;
            z-index: 11;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
        }

        /* Make the main header row stick below summary row */
        #ebay-table thead tr:nth-child(2) th {
            position: sticky;
            top: 36px; /* height of summary row */
            background: #ffffff;
            z-index: 10;
            box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
        }

        /* Optional: ensure headers are tall enough */
        #ebay-table thead th {
            height: 36px;
            text-align: center;
            vertical-align: middle;
        }
        
        #ebay-table {
        color: #000 !important; /* Force dark black font */

        .img-wrapper {
            position: relative;
            display: inline-block;
        }

        .thumbnail-img {
            width: 30px;
            height: auto;
            border-radius: 4px;
            cursor: pointer;
        }

        .popup-img {
            display: none;
            position: absolute;
            top: -10px;
            left: 30px;
            z-index: 999;
            border: 1px solid #ccc;
            background: #fff;
            padding: 4px;
            box-shadow: 0 0 6px rgba(0,0,0,0.2);
        }

        .popup-img img {
            width: 300px;
            height: auto;
            border-radius: 4px;
        }

        .img-wrapper:hover .popup-img {
            display: block;
        }

    }

    </style>
@endsection

@section('content')
    @include('layouts.shared/page-title', ['page_title' => 'Stock Mapping', 'sub_title' => 'View Stock'])
  
    <div class="row">
        <div class="col-12">
            <div class="card">
                <div class="card-body">
                    <h4 class="header-title">View Stock Mapping</h4>

                   
                   
                    <!-- Controls row -->
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <!-- Left side controls -->
                        <div class="form-inline">
                            <div class="form-group mr-2">
                                <label for="row-data-type" class="mr-2">Filter Type:</label>
                                <select id="row-data-type" class="form-control form-control-sm">
                                    <option value="all">All</option>
                                    <option value="listed">Listed</option>
                                    <option value="notlisted">Not Listed</option>
                                    <option value="matching">Matching</option>
                                    <option value="notmatching">Mismatch</option>
                                  
                                </select>
                            </div>
                        </div>
                       

                    <!-- Search on right -->
                        <div class="form-inline">
                            <div class="form-group">
                                <label for="search-input" class="mr-2">Search:</label>
                                <input type="text" id="search-input" class="form-control form-control-sm"
                                    placeholder="Search all columns">
                            </div>
                        </div>
                    </div>
                    <div class="row">
                        {{-- inventory --}}
                        <div class="col-12">
                            <div class="table-container">
                        <table class="custom-resizable-table" id="inventory-table">
                            <thead>
                                <tr><th colspan="5" class="text-center text-bg-success"><b>Inventory</b></th></tr>                                
                                <tr>
                                    <th data-field="sl_no" style="max-width: 150px;">SL No. <span class="sort-arrow">↓</span></th>
                                    <th data-field="SKU" style="vertical-align: middle; white-space: nowrap;">
                                            <div class="d-flex flex-column align-items-center sortable">
                                                <div class="d-flex align-items-left"> SKU <span class="sort-arrow">↓</span></div>                                                                                       
                                            </div>                                        
                                    </th>

                                    <th data-field="TITLE">Title<span class="sort-arrow">↓</span></th>                                    
                                    <th data-field="INV_shopify" style="vertical-align: middle; white-space: nowrap;">
                                        <div class="d-flex flex-column align-items-center">
                                            <div class="d-flex align-items-center">
                                                Inventory Shopfiy <span class="sort-arrow">↓</span>
                                            </div>
                                        </div>
                                    </th>

                                  
                                  
                                     <th data-field="INV_amazon" style="vertical-align: middle; white-space: nowrap;">
                                        <div class="d-flex flex-column align-items-center">
                                            <div class="d-flex align-items-center">
                                                Inventory Amazon <span class="sort-arrow">↓</span>
                                            </div>
                                            
                                            <div class="notlisted-total">
                                                <span id="notlistedamazon" style="color:red"></span>
                                            </div>
                                        </div>
                                    </th>
                                </tr>
                            </thead>
                            <tbody>
                                <!-- Data will be populated by JavaScript -->
                            </tbody>
                        </table>
                    </div>

                    <!-- Pagination controls -->
                    <div class="pagination-controls mt-2">
                        <div class="form-group"> 
                            <span id="visible-rows" class="badge badge-light" style="color: #dc3545;">Showing 1-25 of
                                150</span>
                        </div>
                        
                        <button id="first-page" class="btn btn-sm btn-outline-secondary mr-1">First</button>
                        <button id="prev-page" class="btn btn-sm btn-outline-secondary mr-1">Previous</button>
                        <span id="page-info" class="mx-2">Page 1 of 6</span>
                        <button id="next-page" class="btn btn-sm btn-outline-secondary ml-1">Next</button>
                        <button id="last-page" class="btn btn-sm btn-outline-secondary ml-1">Last</button>
                    </div>

                    <div id="data-loader" class="card-loader-overlay" style="display: none;">
                        <div class="loader-content">
                            <div class="spinner-border text-primary" role="status">
                                <span class="visually-hidden">Loading...</span>
                            </div>
                            <div class="loader-text">Loading data...</div>
                        </div>
                    </div>
                        </div>
                        {{-- inventory --}}
                    </div>
                    
                </div>
            </div>
        </div>
    </div>
@endsection

@section('script')
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.6.2/dist/js/bootstrap.bundle.min.js"></script>
   
  <script>
    let isResizing = false;
    let currentSort = {
        field: null,
        direction: 1
    };
    let tableData = [];
    let filteredData = []; // Added
    let currentPage = 1;   // Added
   let isLoading = false;
   let notlisted=0;

    document.body.style.zoom = "80%";

    $(document).ready(function () {
        initTable();

        function initTable() {
            loadData().then(() => {
                filteredData = [...tableData]; // Initialize filteredData
                renderTable(filteredData);
                initResizableColumns();
                initSorting();
                initPagination();
                initSearch();
                initColumnToggle();
            });
        }

        function loadData() {
            showLoader();
            return $.ajax({
                url: '/stock/mapping/inventory/data',
                type: 'GET',
                dataType: 'json'
            }).then(response => {
                const sheetData = Array.isArray(response.data)
                    ? response.data
                    : Object.values(response.data || {});

                tableData = sheetData.map((item, index) => {
                    const sku = (item.sku || '').toUpperCase().trim();
                    const INV_shopify = parseFloat(item.inventory_shopify) || 0;
                    const INV_amazon = parseFloat(item.inventory_amazon) || item.inventory_amazon;
                        if(INV_amazon=='Not Listed'){notlisted +=1;}
                    return {
                        sl_no: index + 1,
                        SKU: sku,
                        TITLE: item.product_title || item.title,
                        INV_shopify: INV_shopify,
                        INV_amazon: INV_amazon,
                       is_notlisted: (() => {
                            // const invAmazon = INV_amazon?.toString().toLowerCase().trim();

                            return INV_amazon=='Not Listed'?'notlisted':'listed';
                        })(),

                                                matching: (() => {
                            return INV_shopify == INV_amazon ? 'matching' : 'notmatching';
                        })(),

                        raw_data: item
                    };
                });
                console.log(tableData);
                filteredData=tableData;
                renderTable(filteredData);
                hideLoader();
                return tableData;
            }).catch(error => {
                console.error(" Error loading data:", error);
                hideLoader();
                return [];
            });
        }

        function renderTable(data) {
            // console.log(notlisted);
            $('#notlistedamazon').text('(Not Listed:'+notlisted+')');
            const $tableBody = $('#inventory-table tbody');
            $tableBody.empty();

            if (!Array.isArray(data) || data.length === 0) {
                $tableBody.append('<tr><td colspan="5" class="text-center">No data found</td></tr>');
                return;
            }

            data.forEach(item => {
                // const isMismatch = item.INV_shopify !== item.INV_amazon!='Not Listed'?parseFloat(item.INV_amazon):item.INV_amazon;
               const isMismatch = item.INV_amazon !== 'Not Listed'
    ? item.INV_shopify !== parseFloat(item.INV_amazon)
    : item.INV_shopify !== item.INV_amazon;


                const rowClass = isMismatch ? 'style="color:red; text-align:center;"' : 'style="text-align:center;"';

                const row = `
                    <tr ${rowClass}>
                        <td>${item.sl_no}</td>
                        <td>${item.SKU}</td>
                        <td style="max-width:150px; overflow:hidden !important; text-overflow:ellipsis;">${item.TITLE}</td>
                        <td>${item.INV_shopify}</td>
                        <td>${item.INV_amazon}</td>
                    </tr>
                `;
                $tableBody.append(row);
            });
        }

        function initResizableColumns() {
            const $table = $('#inventory-table'); // Correct table ID
            const $headers = $table.find('th');

            $headers.each(function () {
                if (!$(this).find('.resize-handle').length) {
                    $(this).append('<div class="resize-handle"></div>');
                }
            });

            $table.off('mousedown', '.resize-handle').on('mousedown', '.resize-handle', function (e) {
                e.preventDefault();
                e.stopPropagation();
                isResizing = true;
                $(this).addClass('resizing');

                const $th = $(this).parent();
                const columnIndex = $th.index();
                const startX = e.pageX;
                const startWidth = $th.outerWidth();

                $('body').css('user-select', 'none');

                $(document).off('mousemove.resize').on('mousemove.resize', function (e) {
                    if (!isResizing) return;
                    const newWidth = startWidth + (e.pageX - startX);
                    $th.css({ width: newWidth, 'min-width': newWidth, 'max-width': newWidth });
                });

                $(document).off('mouseup.resize').on('mouseup.resize', function () {
                    if (!isResizing) return;
                    $('.resize-handle').removeClass('resizing');
                    $('body').css('user-select', '');
                    isResizing = false;
                    $(document).off('mousemove.resize mouseup.resize');
                });
            });
        }

        function initSorting() {
            const $table = $('#inventory-table');
            $table.find('th[data-field]').addClass('sortable').off('click').on('click', function (e) {
                if (isResizing) return e.stopPropagation();
                if ($(e.target).is('input, .resize-handle')) return;

                const field = $(this).data('field');
                if (!field) return;

                if (currentSort.field === field) {
                    currentSort.direction *= -1;
                } else {
                    currentSort.field = field;
                    currentSort.direction = 1;
                }

                $('.sort-arrow').text('↓');
                $(this).find('.sort-arrow').text(currentSort.direction === 1 ? '↑' : '↓');

                filteredData.sort((a, b) => {
                    let valA = a[field] ?? '';
                    let valB = b[field] ?? '';

                    const numericFields = ['sl_no', 'invshopify', 'invamazon','INV_shopify','INV_amazon'];
                    if (numericFields.includes(field)) {
                        valA = parseFloat(valA) || 0;
                        valB = parseFloat(valB) || 0;
                        return (valA - valB) * currentSort.direction;
                    }

                    // return String(valA).localeCompare(String(valB)) * currentSort.direction;
                    return String(valA).trim().toLowerCase().localeCompare(String(valB).trim().toLowerCase()) * currentSort.direction;
                });

                renderTable(filteredData);
            });
        }

        function initPagination() {
            // Optional: disable if showing all rows
            $('.pagination-controls').hide();
        }

        function initSearch() {
            $('#search-input').off('keyup').on('keyup', function () {
                const term = $(this).val().toLowerCase().trim();
                if (term) {
                    filteredData = tableData.filter(item =>
                        Object.values(item).some(val =>
                            val != null && val.toString().toLowerCase().includes(term)
                        )
                    );
                } else {
                    filteredData = [...tableData];
                }
                renderTable(filteredData);
            });
        }

        function initColumnToggle() {
            const $table = $('#inventory-table'); // Fixed from #ebay-table
            const $headers = $table.find('th[data-field]');
            const $menu = $('#columnToggleMenu');
            const $btn = $('#hideColumnsBtn');

            $menu.empty();
            $headers.each(function () {
                const field = $(this).data('field');
                const title = $(this).text().replace(/ ↑| ↓/g, '').trim();
                const id = `toggle-${field}`;
                $menu.append(`
                    <div class="column-toggle-item">
                        <input type="checkbox" class="column-toggle-checkbox" id="${id}" data-field="${field}" checked>
                        <label for="${id}">${title}</label>
                    </div>
                `);
            });

            $btn.off('click').on('click', e => {
                e.stopPropagation();
                $menu.toggleClass('show');
            });

            $(document).off('click.colToggle').on('click.colToggle', e => {
                if (!$(e.target).closest('.custom-dropdown').length) {
                    $menu.removeClass('show');
                }
            });

            $menu.off('change').on('change', '.column-toggle-checkbox', function () {
                const field = $(this).data('field');
                const visible = $(this).is(':checked');
                const idx = $headers.filter(`[data-field="${field}"]`).index();
                $table.find('tr').each(function () {
                    $(this).find(`td:eq(${idx}), th:eq(${idx})`).toggle(visible);
                });
            });

            $('#showAllColumns').off('click').on('click', () => {
                $menu.find('.column-toggle-checkbox').prop('checked', true).trigger('change');
                $menu.removeClass('show');
            });
        }

        function showLoader() {
            $('#data-loader').fadeIn();
        }

        function hideLoader() {
            $('#data-loader').fadeOut();
        }

        // Optional: add if you need totals (not used in render, but called)
        function calculateTotals() {
            // No-op or implement if needed
        }

        
                $('#row-data-type').on('change', function() {
                    const filterType = $(this).val();
                    applyRowTypeFilter(filterType);
                });
                
function applyRowTypeFilter(filterType) {
    filteredData = [...tableData];
    console.log(filterType);
    if (filterType === 'matching') {
        filteredData = filteredData.filter(item => item.matching === 'matching');
    } else if (filterType === 'notmatching') {
        filteredData = filteredData.filter(item => item.matching === 'notmatching');
    }

    if (filterType === 'listed') {
        filteredData = filteredData.filter(item => item.is_notlisted === 'listed');
    } else if (filterType === 'notlisted') {
        filteredData = filteredData.filter(item => item.is_notlisted === 'notlisted');
    }

    currentPage = 1;
    renderTable(filteredData);
    calculateTotals();
}
         
    });
</script>
@endsection
