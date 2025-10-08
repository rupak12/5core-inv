<?php

namespace App\Http\Controllers\PurchaseMaster;

use App\Http\Controllers\Controller;
use App\Models\Category;
use App\Models\RfqForm;
use App\Models\RfqSubmission;
use Illuminate\Http\Request;
use Illuminate\Support\Str;

class RFQController extends Controller
{
    public function index()
    {
        $categories = Category::all();
        return view('purchase-master.rfq-form.index', compact('categories'));
    }

    public function getRfqFormsData(){
        $forms = RfqForm::select(
            'id',
            'name',
            'title',
            'slug',
            'main_image',
            'subtitle',
            'fields',
            'dimension_inner',
            'product_dimension',
            'package_dimension',
            'created_at',
            'updated_at'
            
        )->get();

        return response()->json([
            'data' => $forms
        ]);
    }

    public function storeRFQForm(Request $request)
    {
        $request->validate([
            'rfq_form_name' => 'required|string',
            'title' => 'required|string',
            'fields' => 'required|array',
            'main_image' => 'nullable|image|max:2048'
        ]);

        $slug = Str::slug($request->rfq_form_name) . '-' . Str::random(5);

        $imagePath = null;
        if($request->hasFile('main_image')){
            $imagePath = $request->file('main_image')->store('rfq_forms', 'public');
        }
        $fields = collect($request->fields)->map(function($field, $index) {
            $field['order'] = $field['order'] ?? ($index + 1);
            return $field;
        })->toArray();

        RfqForm::create([
            'name' => $request->rfq_form_name,
            'title' => $request->title,
            'slug' => $slug,
            'main_image' => $imagePath,
            'subtitle' => $request->subtitle,
            'fields' => $fields,
            'dimension_inner' => $request->dimension_inner,
            'product_dimension' => $request->product_dimension,
            'package_dimension' => $request->package_dimension,
        ]);

        return redirect()->back()->with('flash_message', 'RFQ Form created successfully!');
    }

    public function edit($id)
    {
        $form = RfqForm::findOrFail($id);

        return response()->json([
            'success' => true,
            'data' => $form
        ]);
    }

    public function update(Request $request, $id)
    {
        $request->validate([
            'rfq_form_name' => 'required|string',
            'title' => 'required|string',
            'fields_json' => 'required|string',
            'main_image' => 'nullable|image|max:2048'
        ]);

        $form = RfqForm::findOrFail($id);

        $imagePath = $form->main_image;
        if ($request->hasFile('main_image')) {
            $imagePath = $request->file('main_image')->store('rfq_forms', 'public');
        }

        $fields = json_decode($request->fields_json, true) ?? [];

        $form->update([
            'name' => $request->rfq_form_name,
            'title' => $request->title,
            'main_image' => $imagePath,
            'subtitle' => $request->subtitle,
            'fields' => $fields,
            'dimension_inner' => $request->dimension_inner,
            'product_dimension' => $request->product_dimension,
            'package_dimension' => $request->package_dimension,
        ]);

        return response()->json([
            'success' => true,
            'message' => 'RFQ Form updated successfully!',
        ]);
    }

    public function destroy($id)
    {
        try {
            $form = RfqForm::findOrFail($id);
            $form->delete();

            return response()->json(['success' => true, 'message' => 'Form deleted successfully']);
        } catch (\Exception $e) {
            return response()->json(['success' => false, 'message' => 'Failed to delete form']);
        }
    }

    // Form Reports
    public function rfqReports($slug)
    {   
        $form = RfqForm::where('slug', $slug)->firstOrFail();
        return view('purchase-master.rfq-form.form-submit-reports', compact('form'));
    }

    public function getRfqReportsData($slug)
    {
        $form = RfqForm::where('id', $slug)->firstOrFail();
        $submissions = RfqSubmission::where('rfq_form_id', $form->id)
            ->select('id', 'data', 'created_at')
            ->get();

        return response()->json([
            'data' => $submissions
        ]);
    }

}
