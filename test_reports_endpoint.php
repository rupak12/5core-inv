<?php

require_once 'vendor/autoload.php';

// Load environment variables
$dotenv = Dotenv\Dotenv::createImmutable(__DIR__);
$dotenv->load();

echo "🔍 Testing Amazon SP-API Reports Endpoint\n";
echo "==========================================\n\n";

// Get access token
function getAccessToken() {
    $response = \Illuminate\Support\Facades\Http::asForm()->post('https://api.amazon.com/auth/o2/token', [
        'grant_type' => 'refresh_token',
        'refresh_token' => $_ENV['SPAPI_REFRESH_TOKEN'],
        'client_id' => $_ENV['SPAPI_CLIENT_ID'],
        'client_secret' => $_ENV['SPAPI_CLIENT_SECRET'],
    ]);

    if ($response->failed()) {
        echo "❌ Failed to get access token\n";
        echo "Response: " . $response->body() . "\n";
        return null;
    }

    $data = $response->json();
    return $data['access_token'] ?? null;
}

// Alternative method using cURL for token
function getAccessTokenCurl() {
    $data = [
        'grant_type' => 'refresh_token',
        'refresh_token' => $_ENV['SPAPI_REFRESH_TOKEN'],
        'client_id' => $_ENV['SPAPI_CLIENT_ID'],
        'client_secret' => $_ENV['SPAPI_CLIENT_SECRET'],
    ];

    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, 'https://api.amazon.com/auth/o2/token');
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query($data));
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        'Content-Type: application/x-www-form-urlencoded'
    ]);

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($httpCode !== 200) {
        echo "❌ Failed to get access token (HTTP $httpCode)\n";
        echo "Response: $response\n";
        return null;
    }

    $data = json_decode($response, true);
    return $data['access_token'] ?? null;
}

echo "📍 Step 1: Getting access token...\n";
$accessToken = getAccessTokenCurl();

if (!$accessToken) {
    echo "❌ Cannot proceed without access token\n";
    exit(1);
}

echo "✅ Access token obtained successfully\n\n";

echo "📍 Step 2: Testing reports endpoint...\n";

// Test GET_SALES_AND_TRAFFIC_REPORT specifically
$reportTypes = [
    'GET_SALES_AND_TRAFFIC_REPORT',
    'GET_AFN_INVENTORY_DATA',
    'GET_MERCHANT_LISTINGS_ALL_DATA', 
    'GET_MERCHANT_LISTINGS_DATA',
    'GET_FBA_FULFILLMENT_INVENTORY_HEALTH_DATA'
];

$marketplaceIds = $_ENV['SPAPI_MARKETPLACE_ID'] ?? 'ATVPDKIKX0DER';

// Test with first report type
$reportType = $reportTypes[0];
$endpoint = 'https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports?' . http_build_query([
    'reportTypes' => $reportType,
    'marketplaceIds' => $marketplaceIds
]);

echo "🔍 Testing with report type: $reportType\n";
echo "📍 URL: $endpoint\n\n";

$ch = curl_init();
curl_setopt($ch, CURLOPT_URL, $endpoint);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_HTTPHEADER, [
    'x-amz-access-token: ' . $accessToken,
    'Content-Type: application/json'
]);

$response = curl_exec($ch);
$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
curl_close($ch);

echo "HTTP Status Code: $httpCode\n";
echo "Response:\n";
echo "=========\n";
echo $response . "\n\n";

// Pretty print JSON if it's valid JSON
$jsonData = json_decode($response, true);
if (json_last_error() === JSON_ERROR_NONE) {
    echo "📊 Formatted JSON Response:\n";
    echo "===========================\n";
    echo json_encode($jsonData, JSON_PRETTY_PRINT) . "\n\n";
    
    // Analyze the structure
    if (isset($jsonData['reports'])) {
        echo "📋 Found " . count($jsonData['reports']) . " reports\n";
        
        if (!empty($jsonData['reports'])) {
            $firstReport = $jsonData['reports'][0];
            echo "📝 First report structure:\n";
            foreach ($firstReport as $key => $value) {
                $type = is_array($value) ? 'array' : gettype($value);
                echo "  - $key: $type\n";
            }
        }
    }
    
    if (isset($jsonData['nextToken'])) {
        echo "🔄 Pagination available (nextToken present)\n";
    }
    
    // Test multiple report types
    if (isset($jsonData['reports']) && !empty($jsonData['reports'])) {
        echo "\n📊 Testing other report types...\n";
        echo "=================================\n";
        
        foreach ($reportTypes as $idx => $testReportType) {
            if ($idx === 0) continue; // Skip first one already tested
            
            $testEndpoint = 'https://sellingpartnerapi-na.amazon.com/reports/2021-06-30/reports?' . http_build_query([
                'reportTypes' => $testReportType,
                'marketplaceIds' => $marketplaceIds,
                'pageSize' => 5  // Limit results
            ]);
            
            $ch2 = curl_init();
            curl_setopt($ch2, CURLOPT_URL, $testEndpoint);
            curl_setopt($ch2, CURLOPT_RETURNTRANSFER, true);
            curl_setopt($ch2, CURLOPT_HTTPHEADER, [
                'x-amz-access-token: ' . $accessToken,
                'Content-Type: application/json'
            ]);
            
            $testResponse = curl_exec($ch2);
            $testHttpCode = curl_getinfo($ch2, CURLINFO_HTTP_CODE);
            curl_close($ch2);
            
            $testData = json_decode($testResponse, true);
            $reportCount = isset($testData['reports']) ? count($testData['reports']) : 0;
            
            echo "🔍 $testReportType: HTTP $testHttpCode, Reports: $reportCount\n";
            
            if ($idx >= 2) break; // Test only first 3 to save time
        }
    }
}

echo "\n🏁 Test completed!\n";