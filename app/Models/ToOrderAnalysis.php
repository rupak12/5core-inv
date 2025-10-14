<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Factories\HasFactory;

class ToOrderAnalysis extends Model
{
    use HasFactory;

    protected $table = 'to_order_analysis';

    protected $fillable = [
        'parent',
        'sku',
        'rfq_form_link',
        'date_apprvl',
        'mail_sent',
        'rfq_report',
        'stage',
        'nrl',
        'sheet_link',
        'supplier_name',
        'advance_date',
        'order_qty',
    ];
    
    public function supplier()
    {
        return $this->belongsTo(Supplier::class);
    }
}
