<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use App\Models\FbaTable;
use App\Models\FbaPrice;
use App\Models\FbaMonthlySale;

class FetchFbaMonthlySales extends Command
{
    protected $signature = 'app:fetch-fba-monthly-sales
        {--year= : Year to fetch (default current year)}
        {--source=both : source of SKUs: inventory|prices|both}
        {--debug : show debug info}';

    protected $description = 'Fetch monthly sales (Jan-Dec) per SKU and insert/update fba_monthly_sales';

    public function handle()
    {
        $this->info('📈 Starting monthly sales fetch...');

        $year = intval($this->option('year') ?: date('Y'));
        $source = $this->option('source') ?: 'both';

        $accessToken = $this->getAccessToken();
        if (!$accessToken) {
            $this->error('❌ Could not obtain access token.');
            return 1;
        }

        $endpoint = env('SPAPI_ENDPOINT', 'https://sellingpartnerapi-na.amazon.com');
        $marketplace = env('SPAPI_MARKETPLACE_ID', 'ATVPDKIKX0DER');

        // Collect SKUs
        $skus = collect();

        if (in_array($source, ['inventory','both'])) {
            $skus = $skus->merge(FbaTable::pluck('seller_sku')->filter()->unique());
        }

        if (in_array($source, ['prices','both'])) {
            $skus = $skus->merge(FbaPrice::pluck('seller_sku')->filter()->unique());
        }

        $skus = $skus->unique()->values()->all();

        if (empty($skus)) {
            $this->warn('⚠️ No SKUs found to fetch.');
            return 0;
        }

        $this->info("ℹ️ SKUs to process: " . count($skus));

        // interval: full year in ISO with timezone offset (use -07:00 like your example; adjust if needed)
        $start = "{$year}-01-01T00:00:00-07:00";
        $end = "{$year}-12-31T23:59:59-07:00";
        $interval = "{$start}--{$end}";

        // Iterate SKUs
        foreach ($skus as $sku) {
            try {
                // Build URL - note: adjust sku param name if SP-API expects something else.
                $url = "{$endpoint}/sales/v1/orderMetrics?marketplaceIds={$marketplace}&interval={$interval}&granularity=Month&sku={$sku}";

                $res = Http::withHeaders([
                    'x-amz-access-token' => $accessToken,
                    'Content-Type' => 'application/json',
                ])->get($url);

                if ($res->failed()) {
                    $this->warn("⚠️ API failed for SKU {$sku}: " . $res->status());
                    if ($this->option('debug')) $this->line($res->body());
                    // continue to next SKU
                    continue;
                }

                $payload = $res->json()['payload'] ?? null;
                if (!$payload || !is_array($payload) || count($payload) === 0) {
                    $this->warn("ℹ️ No metrics returned for SKU {$sku} (may be zero or restricted)");
                    // ensure a 0 row exists
                    FbaMonthlySale::updateOrCreate(
                        ['seller_sku' => $sku, 'year' => $year],
                        ['total_units' => 0]
                    );
                    continue;
                }

                // Initialize months container (0-11 index)
                $months = array_fill(1, 12, 0);
                $totalUnits = 0;
                $totalRevenue = 0.0; // will be used for weighted avg

                foreach ($payload as $period) {
                    // defensive parsing: period may include 'interval' and metrics keys
                    $intervalStr = $period['interval'] ?? null;
                    $unitCount = intval($period['unitCount'] ?? 0);

                    // Some responses provide 'orderedProductSales' => ['amount' => X]
                    $orderedProductSales = 0;
                    if (!empty($period['orderedProductSales']['amount'])) {
                        $orderedProductSales = floatval($period['orderedProductSales']['amount']);
                    } elseif (!empty($period['orderedProductSales'])) {
                        // defensive: might be number directly
                        $orderedProductSales = floatval($period['orderedProductSales']);
                    }

                    // some APIs include SKU+ASIN in the payload; try to capture asin:
                    $asin = $period['asin'] ?? null;

                    if ($intervalStr) {
                        // interval like "2025-01-01T00:00:00-07:00--2025-01-31T23:59:59-07:00"
                        $parts = explode('--', $intervalStr);
                        $startPart = $parts[0] ?? null;
                        if ($startPart) {
                            $monthNum = intval(date('n', strtotime($startPart))); // 1..12
                            if ($monthNum >=1 && $monthNum <= 12) {
                                $months[$monthNum] += $unitCount;
                                $totalUnits += $unitCount;
                                $totalRevenue += $orderedProductSales;
                            }
                        }
                    } else {
                        // fallback: if there's a 'month' field
                        if (!empty($period['month'])) {
                            $m = intval($period['month']);
                            if ($m >=1 && $m <=12) {
                                $months[$m] += $unitCount;
                                $totalUnits += $unitCount;
                                $totalRevenue += $orderedProductSales;
                            }
                        }
                    }
                }

                // compute average price (weighted)
                $avgPrice = null;
                if ($totalUnits > 0 && $totalRevenue > 0) {
                    $avgPrice = round($totalRevenue / $totalUnits, 2);
                } else {
                    // fallback: use current price from fba_prices table if available
                    $price = \App\Models\FbaPrice::where('seller_sku', $sku)->value('price');
                    $avgPrice = $price ? floatval($price) : null;
                }

                // map to keys jan..dec
                $dataToSave = [
                    'asin' => $asin,
                    'year' => $year,
                    'jan' => $months[1] ?? 0,
                    'feb' => $months[2] ?? 0,
                    'mar' => $months[3] ?? 0,
                    'apr' => $months[4] ?? 0,
                    'may' => $months[5] ?? 0,
                    'jun' => $months[6] ?? 0,
                    'jul' => $months[7] ?? 0,
                    'aug' => $months[8] ?? 0,
                    'sep' => $months[9] ?? 0,
                    'oct' => $months[10] ?? 0,
                    'nov' => $months[11] ?? 0,
                    'dec' => $months[12] ?? 0,
                    'total_units' => $totalUnits,
                    'avg_price' => $avgPrice
                ];

                FbaMonthlySale::updateOrCreate(
                    ['seller_sku' => $sku, 'year' => $year],
                    $dataToSave
                );

                $this->info("✅ Saved monthly metrics for {$sku} (units: {$totalUnits})");

            } catch (\Throwable $e) {
                Log::error("Error for SKU {$sku}: " . $e->getMessage());
                $this->error("❌ Error processing SKU {$sku}");
            }

            // small delay to avoid throttling
            sleep(1);
        }

        $this->info('🏁 Done fetching monthly sales.');
        return 0;
    }

    private function getAccessToken()
    {
        try {
            $res = Http::asForm()->post('https://api.amazon.com/auth/o2/token', [
                'grant_type' => 'refresh_token',
                'refresh_token' => env('SPAPI_REFRESH_TOKEN'),
                'client_id' => env('SPAPI_CLIENT_ID'),
                'client_secret' => env('SPAPI_CLIENT_SECRET'),
            ]);

            if ($res->failed()) {
                Log::error('Access token request failed', $res->json());
                return null;
            }

            return $res->json()['access_token'] ?? null;
        } catch (\Throwable $e) {
            Log::error('Access token error: ' . $e->getMessage());
            return null;
        }
    }
}
