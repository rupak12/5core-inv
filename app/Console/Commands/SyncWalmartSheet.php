<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use App\Http\Controllers\ApiController;
use App\Models\WalmartProductSheet;



class SyncWalmartSheet extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'sync:walmart-sheet';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Sync Walmart product sheet';

    /**
     * Execute the console command.
     */
    public function handle()
    {
        $controller = new ApiController();
        $sheet = $controller->fetchDataFromWalmartListingDataSheet();
        $rows = collect($sheet->getData()->data ?? []);

        foreach ($rows as $row) {
            $sku = trim($row->{'(Child) sku'} ?? '');
            if (!$sku) continue;

            WalmartProductSheet::updateOrCreate(
                ['sku' => $sku],
                [
                    'price'     => $this->toDecimalOrNull($row->{'Price'} ?? null),
                    'pft'       => $this->toDecimalOrNull($row->{'Pft%'} ?? null),
                    'roi'       => $this->toDecimalOrNull($row->{'ROI%'} ?? null),
                    'l30'       => $this->toIntOrNull($row->{'WL30'} ?? null),
                    'dil'       => $this->toDecimalOrNull($row->{'Dil%'} ?? null),
                    'buy_link'  => trim($row->{'Buyer Link'} ?? ''),
                ]
            );
        }


        $this->info('Walmart sheet synced successfully!');

    }

    private function toDecimalOrNull($value)
    {
        return is_numeric($value) ? round((float)$value, 2) : null;
    }

    private function toIntOrNull($value)
    {
        return is_numeric($value) ? (int)$value : null;
    }
}
