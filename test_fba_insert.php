<?php

require_once 'vendor/autoload.php';

use App\Models\FbaTable;
use Illuminate\Container\Container;
use Illuminate\Database\Capsule\Manager as Capsule;
use Illuminate\Events\Dispatcher;

// Load Laravel app
$app = require_once 'bootstrap/app.php';
$app->make('Illuminate\Contracts\Console\Kernel')->bootstrap();

// Read cached data
$csvData = file_get_contents(storage_path('app/fba_report_cache.csv'));
$lines = explode("\n", trim($csvData));
$headers = str_getcsv(array_shift($lines), "\t");

echo "📊 Records: " . count($lines) . " | Columns: " . count($headers) . "\n";
echo "📋 Headers: " . implode(', ', $headers) . "\n\n";

// Truncate table
FbaTable::truncate();
echo "🗑️  Truncated fba_table\n";

$inserted = 0;
foreach ($lines as $line) {
    if (!trim($line)) continue;
    $row = str_getcsv($line, "\t");
    if (count($row) < count($headers)) continue;
    $data = array_combine($headers, $row);

    try {
        FbaTable::create([
            'seller_sku' => $data['sku'] ?? '',
            'fulfillment_channel_sku' => $data['fnsku'] ?? '',
            'asin' => $data['asin'] ?? '',
            'condition_type' => $data['condition'] ?? '',
            'quantity_available' => (int)($data['afn-fulfillable-quantity'] ?? 0)
        ]);
        $inserted++;
    } catch (Exception $e) {
        echo "❌ Error inserting: " . $e->getMessage() . "\n";
    }
}

echo "✅ Inserted {$inserted} records into fba_table\n";