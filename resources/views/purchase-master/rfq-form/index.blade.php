@extends('layouts.vertical', ['title' => 'RFQ Form'])
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
</style>
@endsection
@section('content')
@include('layouts.shared.page-title', ['page_title' => 'RFQ Form', 'sub_title' => 'RFQ Form'])

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
                        <button id="add-new-row" class="btn btn-sm btn-success" data-bs-toggle="modal" data-bs-target="#createRFQFormModal">
                            <i class="fas fa-plus-circle me-1"></i> Create RFQ Form
                        </button>
                    </div>
                </div>
                <div id="rfq-form-table"></div>
            </div>
        </div>
    </div>
</div>

{{-- add rfq form modal --}}
<div class="modal fade" id="createRFQFormModal" tabindex="-1" aria-labelledby="createRFQFormModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-centered shadow-none">
        <div class="modal-content border-0 shadow-lg">
            <div class="modal-header bg-primary text-white">
                <h5 class="modal-title fw-bold" id="createRFQFormModalLabel">
                    <i class="fas fa-file-invoice me-2"></i> Create RFQ Form
                </h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>

            <form id="rfqFormCreate" method="POST" action="{{ route('rfq-form.store') }}" enctype="multipart/form-data" autocomplete="off">
                @csrf
                <div class="modal-body">

                    <!-- Section 1: Basic Info -->
                    <div class="border p-3 rounded mb-3">
                        <h6 class="fw-bold mb-3">Basic Information</h6>

                        <div class="row">
                            <div class="col-md-4 mb-3">
                                <label for="rfq_form_name" class="form-label">RFQ Form Name <span class="text-danger">*</span></label>
                                <input type="text" name="rfq_form_name" id="rfq_form_name" class="form-control" placeholder="Stand" required>
                            </div>

                            <div class="col-md-4 mb-3">
                                <label for="title" class="form-label">Form Heading / Title <span class="text-danger">*</span></label>
                                <input type="text" name="title" id="title" class="form-control" placeholder="Enter form heading" required>
                            </div>

                            <div class="col-md-4 mb-3">
                                <label for="main_image" class="form-label">Form Image (optional)</label>
                                <input type="file" name="main_image" id="main_image" class="form-control" accept="image/*">
                                <img id="mainImagePreview" src="#" alt="Preview" class="img-fluid mt-2" style="display:none; max-height:150px;">
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-md-6">
                                <label for="subtitle" class="form-label">Form Subtitle / Description</label>
                                <textarea name="subtitle" id="subtitle" class="form-control" rows="3" placeholder="Enter form description"></textarea>
                            </div>
                            <div class="col-md-2 d-flex flex-column align-item-center justify-content-center">
                                <div class="form-check mb-2">
                                    <input type="hidden" name="dimension_inner" value="false">
                                    <input class="form-check-input" type="checkbox" value="true" id="checkbox1" name="dimension_inner">
                                    <label class="form-check-label" for="checkbox1">Dimension Inner Box</label>
                                </div>

                                <div class="form-check mb-2">
                                    <input type="hidden" name="product_dimension" value="false">
                                    <input class="form-check-input" type="checkbox" value="true" id="checkbox2" name="product_dimension">
                                    <label class="form-check-label" for="checkbox2">Product Dimension</label>
                                </div>

                                <div class="form-check mb-2">
                                    <input type="hidden" name="package_dimension" value="false">
                                    <input class="form-check-input" type="checkbox" value="true" id="checkbox3" name="package_dimension">
                                    <label class="form-check-label" for="checkbox3">Package Dimension</label>
                                </div>

                            </div>
                        </div>
                    </div>

                    <!-- Section 2: Dynamic Fields -->
                    <div class="border p-3 rounded">
                        <h6 class="fw-bold mb-3">Add Fields</h6>

                        <div id="dynamicFieldsWrapper">
                            <div class="row g-3 mb-2 field-item">
                                <div class="col-md-3">
                                    <input type="text" name="fields[0][label]" class="form-control field-label" placeholder="Field Label" required>
                                </div>
                                <div class="col-md-3">
                                    <input type="text" name="fields[0][name]" class="form-control field-name" placeholder="Field Name (auto)" readonly>
                                </div>
                                <div class="col-md-2">
                                    <select name="fields[0][type]" class="form-select field-type">
                                        <option value="text">Text</option>
                                        <option value="number">Number</option>
                                        <option value="select">Select</option>
                                    </select>
                                </div>
                                <div class="col-md-3 select-options-wrapper" style="display:none;">
                                    <input type="text" name="fields[0][options]" class="form-control" placeholder="Options (comma separated)">
                                </div>
                                <div class="col-md-1">
                                    <input type="checkbox" name="fields[0][required]" class="form-check-input mt-2" value="1"> Required
                                </div>
                                <div class="col-md-1">
                                    <button type="button" class="btn btn-danger btn-sm remove-field">X</button>
                                </div>
                            </div>
                        </div>

                        <button type="button" class="btn btn-success btn-sm mt-2" id="addFieldBtn">
                            <i class="fas fa-plus"></i> Add Field
                        </button>
                    </div>


                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                    <button type="submit" class="btn btn-primary">Create Form</button>
                </div>
            </form>
        </div>
    </div>
</div>

{{-- edit rfq form modal --}}
<div class="modal fade" id="editRFQFormModal" tabindex="-1" aria-labelledby="editRFQFormModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-xl modal-dialog-centered shadow-none">
        <div class="modal-content border-0 shadow-lg">
            <div class="modal-header bg-warning text-white">
                <h5 class="modal-title fw-bold" id="editRFQFormModalLabel">
                    <i class="fas fa-edit me-2"></i> Edit RFQ Form
                </h5>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>

            <form id="rfqFormEdit" method="POST" enctype="multipart/form-data" autocomplete="off">
                @csrf
                <div class="modal-body">
                    <input type="hidden" id="edit_form_id" name="form_id">

                    <!-- Similar fields as Create Modal -->
                    <div class="border p-3 rounded mb-3">
                        <h6 class="fw-bold mb-3">Basic Information</h6>
                        <div class="row">
                            <div class="col-md-4 mb-3">
                                <label for="edit_rfq_form_name" class="form-label">RFQ Form Name</label>
                                <input type="text" name="rfq_form_name" id="edit_rfq_form_name" class="form-control" required>
                            </div>
                            <div class="col-md-4 mb-3">
                                <label for="edit_title" class="form-label">Form Heading / Title</label>
                                <input type="text" name="title" id="edit_title" class="form-control" required>
                            </div>
                            <div class="col-md-4 mb-3">
                                <label for="edit_main_image" class="form-label">Form Image</label>
                                <input type="file" name="main_image" id="edit_main_image" class="form-control" accept="image/*">
                                <img id="editMainImagePreview" src="#" class="img-fluid mt-2" style="display:none; max-height:50px;">
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-6">
                                <label for="edit_subtitle" class="form-label">Form Subtitle</label>
                                <textarea name="subtitle" id="edit_subtitle" class="form-control" rows="3"></textarea>
                            </div>
                            <div class="col-md-2 d-flex flex-column align-item-center justify-content-center">
                                <div class="form-check mb-2">
                                    <input type="hidden" name="dimension_inner" value="false">
                                    <input class="form-check-input" type="checkbox" value="true" id="editDimensionInner" name="dimension_inner">
                                    <label class="form-check-label" for="editDimensionInner">Dimension Inner Box</label>
                                </div>

                                <div class="form-check mb-2">
                                    <input type="hidden" name="product_dimension" value="false">
                                    <input class="form-check-input" type="checkbox" value="true" id="editProductDimension" name="product_dimension">
                                    <label class="form-check-label" for="editProductDimension">Product Dimension</label>
                                </div>

                                <div class="form-check mb-2">
                                    <input type="hidden" name="package_dimension" value="false">
                                    <input class="form-check-input" type="checkbox" value="true" id="editPackageDimension" name="package_dimension">
                                    <label class="form-check-label" for="editPackageDimension">Package Dimension</label>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Dynamic Fields -->
                    <div class="border p-3 rounded">
                        <h6 class="fw-bold mb-3">Edit Fields</h6>
                        <div id="editDynamicFieldsWrapper"></div>
                        <button type="button" class="btn btn-success btn-sm mt-2" id="addEditFieldBtn">
                            <i class="fas fa-plus"></i> Add Field
                        </button>
                    </div>
                </div>

                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                    <button type="submit" class="btn btn-warning">Update Form</button>
                </div>
            </form>
        </div>
    </div>
</div>

@endsection
@section('script')
<script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
<script>
    document.addEventListener('DOMContentLoaded', function () {

        const table = new Tabulator("#rfq-form-table", {
            ajaxURL: "/rfq-form/data",
            ajaxConfig: "GET",
            layout: "fitColumns",
            pagination: true,
            paginationSize: 50,
            paginationMode: "local",
            movableColumns: false,
            resizableColumns: true,
            height: "500px",
            columns: [
                {
                    title: "S.No",
                    formatter: "rownum",
                    hozAlign: "center",
                    width: 80,
                    headerHozAlign: "center",
                },
                {
                    title: "Form Name",
                    field: "name",
                    formatter: function(cell){
                        let value = cell.getValue();
                        return value;
                    }
                },
                {
                    title: "Report Link",
                    field: "slug",
                    formatter: function(cell, formatterParams, onRendered){
                        const slug = cell.getValue();
                        if(!slug) return "";

                        const fullUrl = window.location.origin + `/rfq-form/reports/${slug}`;

                        return `
                            <div class="d-flex justify-content-center align-item-center">
                                <a href="${fullUrl}" target="_blank" class="btn btn-sm btn-info me-2">Report <i class="fa-solid fa-link"></i></a>
                            </div>
                        `;
                    },
                },
                {
                    title: "Form Link",
                    field: "slug",
                    formatter: function(cell, formatterParams, onRendered){
                        const slug = cell.getValue();
                        if(!slug) return "";

                        const fullUrl = window.location.origin + `/api/rfq-form/${slug}`;

                        return `
                            <div class="d-flex justify-content-center align-item-center">
                                <a href="${fullUrl}" target="_blank" class="btn btn-sm btn-outline-info me-2"><i class="fa-solid fa-link"></i></a>
                                <button class="btn btn-sm btn-outline-primary copy-btn" data-slug="${fullUrl}"><i class="fa-regular fa-copy"></i></button>
                            </div>
                        `;
                    },
                    cellClick: function(e, cell){
                        if(e.target.classList.contains('copy-btn')){
                            const slug = e.target.dataset.slug;
                            navigator.clipboard.writeText(slug).then(() => {
                                alert('Link copied: ' + slug);
                            }).catch(() => {
                                alert('Failed to copy slug');
                            });
                        }
                    }
                },
                {
                    title: "Created Date",
                    field: "created_at",
                    formatter: function(cell){
                        let value = cell.getValue();
                        return value ? moment(value).format('YYYY-MM-DD') : '';
                    }
                },
                {
                    title: "Updated Date",
                    field: "updated_at",
                    formatter: function(cell){
                        let value = cell.getValue();
                        return value ? moment(value).format('YYYY-MM-DD') : '';
                    }
                },
                {
                    title: "Action",
                    field: "name",
                    hozAlign: "center",
                    formatter: function(cell, formatterParams, onRendered) {
                        const rowData = cell.getData();
                        const editUrl = `/rfq-form/edit/${rowData.id}`;
                        const deleteUrl = `/rfq-form/delete/${rowData.id}`;

                        return `
                            <div class="d-flex justify-content-center align-item-center">
                                <a href="#" class="btn btn-sm btn-success me-2 edit-btn" data-id="${rowData.id}" title="Edit" style="cursor:pointer;">
                                    <i class="fa-solid fa-pen-to-square"></i>
                                </a>
                                <button class="btn btn-sm btn-danger delete-btn" data-id="${rowData.id}" title="Delete" style="cursor:pointer;">
                                    <i class="fa-solid fa-trash"></i>
                                </button>
                            </div>
                        `;
                    },
                    cellClick: function(e, cell) {
                        if(e.target.closest('.delete-btn')) {
                            const btn = e.target.closest('.delete-btn');
                            const id = btn.dataset.id;

                            if(confirm('Are you sure you want to delete this form?')) {
                                fetch(`/rfq-form/delete/${id}`, {
                                    method: 'DELETE',
                                    headers: {
                                        'X-CSRF-TOKEN': '{{ csrf_token() }}',
                                        'Content-Type': 'application/json'
                                    }
                                })
                                .then(r => r.json())
                                .then(r => {
                                    if(r.success) {
                                        cell.getRow().delete();
                                        alert('Form deleted successfully!');
                                    } else {
                                        alert('Failed to delete form: ' + r.message);
                                    }
                                })
                                .catch(() => alert('Error deleting form'));
                            }
                        }
                    }
                }

            ],
            ajaxResponse: function(url, params, response){
                return response.data;
            },
        });


        let fieldCount = 1;

        function slugify(text){
            return text.toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'');
        }

        function createFieldRow(index){
            return `
            <div class="row g-3 mb-2 field-item">
                <div class="col-md-3">
                    <input type="text" name="fields[${index}][label]" class="form-control field-label" placeholder="Field Label" required>
                </div>
                <div class="col-md-3">
                    <input type="text" name="fields[${index}][name]" class="form-control field-name" placeholder="Field Name (auto)" readonly>
                </div>
                <div class="col-md-2">
                    <select name="fields[${index}][type]" class="form-select field-type">
                        <option value="text">Text</option>
                        <option value="number">Number</option>
                        <option value="select">Select</option>
                    </select>
                </div>
                <div class="col-md-3 select-options-wrapper" style="display:none;">
                    <input type="text" name="fields[${index}][options]" class="form-control" placeholder="Options (comma separated)">
                </div>
                <div class="col-md-1">
                    <input type="checkbox" name="fields[${index}][required]" class="form-check-input mt-2" value="1"> Required
                </div>
                <div class="col-md-1">
                    <button type="button" class="btn btn-danger btn-sm remove-field">X</button>
                </div>
            </div>
            `;
        }

        // Add new field
        document.getElementById('addFieldBtn').addEventListener('click', function(){
            let wrapper = document.getElementById('dynamicFieldsWrapper');
            wrapper.insertAdjacentHTML('beforeend', createFieldRow(fieldCount));
            fieldCount++;
        });

        document.getElementById('addEditFieldBtn').addEventListener('click', function(){
            let wrapper = document.getElementById('editDynamicFieldsWrapper');
            wrapper.insertAdjacentHTML('beforeend', createFieldRow(fieldCount));
            fieldCount++;
        });

        // Remove field
        document.addEventListener('click', function(e){
            if(e.target && e.target.classList.contains('remove-field')){
                e.target.closest('.field-item').remove();
            }
        });

        // Show/hide options input if type is select
        document.addEventListener('change', function(e){
            if(e.target && e.target.classList.contains('field-type')){
                let optionsWrapper = e.target.closest('.field-item').querySelector('.select-options-wrapper');
                if(e.target.value === 'select'){
                    optionsWrapper.style.display = 'block';
                } else {
                    optionsWrapper.style.display = 'none';
                }
            }
        });

        // Auto-fill field name from label
        document.addEventListener('input', function(e){
            if(e.target && e.target.classList.contains('field-label')){
                let nameInput = e.target.closest('.field-item').querySelector('.field-name');
                nameInput.value = slugify(e.target.value);
            }
        });

        // edit rfq form
        document.addEventListener('click', function(e) {
            if(e.target.closest('.edit-btn')) {
                const id = e.target.closest('.edit-btn').dataset.id;

                fetch(`/rfq-form/edit/${id}`)
                .then(res => res.json())
                .then(res => {
                    if(res.success) {
                        const data = res.data;

                        document.getElementById('edit_form_id').value = data.id;
                        document.getElementById('edit_rfq_form_name').value = data.name;
                        document.getElementById('edit_title').value = data.title;
                        document.getElementById('edit_subtitle').value = data.subtitle;

                        document.getElementById("editDimensionInner").checked = data.dimension_inner === true || data.dimension_inner === 1 || data.dimension_inner === "true";
                        document.getElementById("editProductDimension").checked = data.product_dimension === true || data.product_dimension === 1 || data.product_dimension === "true";
                        document.getElementById("editPackageDimension").checked = data.package_dimension === true || data.package_dimension === 1 || data.package_dimension === "true";

                        if(data.main_image){
                            document.getElementById('editMainImagePreview').src = "/storage/" + data.main_image;
                            document.getElementById('editMainImagePreview').style.display = 'block';
                        }

                        // Dynamic fields
                        const wrapper = document.getElementById('editDynamicFieldsWrapper');
                        wrapper.innerHTML = '';
                        data.fields.forEach((field, index) => {
                            wrapper.insertAdjacentHTML('beforeend', `
                                <div class="row g-3 mb-2 field-item">
                                    <div class="col-md-3">
                                        <input type="text" name="fields[${index}][label]" class="form-control field-label" value="${field.label}" required>
                                    </div>
                                    <div class="col-md-3">
                                        <input type="text" name="fields[${index}][name]" class="form-control field-name" value="${field.name}" readonly>
                                    </div>
                                    <div class="col-md-2">
                                        <select name="fields[${index}][type]" class="form-select field-type">
                                            <option value="text" ${field.type === 'text' ? 'selected':''}>Text</option>
                                            <option value="number" ${field.type === 'number' ? 'selected':''}>Number</option>
                                            <option value="select" ${field.type === 'select' ? 'selected':''}>Select</option>
                                        </select>
                                    </div>
                                    <div class="col-md-3 select-options-wrapper" style="${field.type === 'select' ? 'display:block':'display:none'};">
                                        <input type="text" name="fields[${index}][options]" class="form-control" value="${field.options || ''}">
                                    </div>
                                    <div class="col-md-1">
                                        <input type="checkbox" name="fields[${index}][required]" class="form-check-input mt-2" value="1" ${field.required ? 'checked':''}>
                                    </div>
                                    <div class="col-md-1">
                                        <button type="button" class="btn btn-danger btn-sm remove-field">X</button>
                                    </div>
                                </div>
                            `);
                        });

                        let myModal = new bootstrap.Modal(document.getElementById('editRFQFormModal'));
                        myModal.show();
                    }
                });
            }
        });

        // Update Submit
        document.getElementById('rfqFormEdit').addEventListener('submit', function(e){
            e.preventDefault();

            const id = document.getElementById('edit_form_id').value;

            // Collect all dynamic fields properly
            let fields = [];
            document.querySelectorAll('#editDynamicFieldsWrapper .field-item').forEach((item, index) => {
                const label = item.querySelector('.field-label')?.value || '';
                const name = item.querySelector('.field-name')?.value || '';
                const type = item.querySelector('.field-type')?.value || 'text';
                const options = item.querySelector('[name*="[options]"]')?.value || '';
                const required = item.querySelector('[name*="[required]"]')?.checked ? 1 : 0;
                fields.push({ label, name, type, options, required, order: index + 1 });
            });

            const formData = new FormData(this);
            formData.append('fields_json', JSON.stringify(fields));

            fetch(`/rfq-form/update/${id}`, {
                method: 'POST',
                body: formData,
                headers: {
                    'X-CSRF-TOKEN': '{{ csrf_token() }}'
                }
            })
            .then(res => res.json())
            .then(res => {
                if(res.success){
                    alert("Form updated successfully!");
                    location.reload();
                } else {
                    alert("Update failed: " + res.message);
                }
            })
            .catch(() => alert("Error updating form"));
        });

    });
</script>

@endsection