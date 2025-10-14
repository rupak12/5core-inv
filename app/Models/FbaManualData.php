<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class FbaManualData extends Model
{
    use HasFactory;

    protected $fillable = ['sku', 'data'];

    protected $casts = [
        'data' => 'array', // Cast data to array
    ];
}
