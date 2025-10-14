<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class WmpMarkAsDone extends Model
{
    use HasFactory;

    protected $table = 'wmp_mark_as_dones';

    protected $fillable = [
        'parent',
        'sku',
        'done_date',
        'is_done',
    ];
}
