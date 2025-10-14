<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class ReverbProduct extends Model
{
    use HasFactory;

    protected $table = 'reverb_products';

    protected $fillable = [
        'sku',
        'r_l30',
        'r_l60',
        'price',
        'views',        
        'status',
    ];
}
