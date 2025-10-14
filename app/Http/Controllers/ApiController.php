<?php

namespace App\Http\Controllers;

use App\Models\ProductMaster;
use App\Models\ShopifySku;
use Illuminate\Http\Request;
use App\Models\User;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Validator;

class ApiController extends Controller
{
    // Fetch data
    public function getData()
    {
        return response()->json([
            'message' => 'API Working',
            'status' => 200
        ]);
    }



    public function updateVerifiedStock(Request $request)
    {
        $sku = $request->input('sku');
        $verifiedStock = $request->input('verified_stock');

        // Step 1: Find variant by SKU
        $productsResponse = Http::withHeaders([
            'X-Shopify-Access-Token' => env('SHOPIFY_PASSWORD'),
        ])->get("https://" . env('SHOPIFY_STORE_URL') . "/admin/api/2025-01/products.json", [
            'fields' => 'id,title,variants'
        ]);

        foreach ($productsResponse['products'] as $product) {
            foreach ($product['variants'] as $variant) {
                if ($variant['sku'] == $sku) {
                    $variantId = $variant['id'];

                    // Step 2: Create or update metafield
                    $metafieldResponse = Http::withHeaders([
                        'X-Shopify-Access-Token' => env('SHOPIFY_PASSWORD'),
                        'Content-Type' => 'application/json'
                    ])->post("https://" . env('SHOPIFY_STORE_URL') . "/admin/api/2025-01/metafields.json", [
                        'metafield' => [
                            'namespace' => 'custom',
                            'key' => 'verified_stock',
                            'value' => (int) $verifiedStock,
                            'type' => 'number_integer',
                            'owner_id' => $variantId,
                            'owner_resource' => 'variant'
                        ]
                    ]);

                    return response()->json(['success' => true, 'metafield' => $metafieldResponse->json()]);
                }
            }
        }

        return response()->json(['error' => 'SKU not found'], 404);
    }


    // Fetch data from Shopify b2c Apps Script
    public function fetchShopifyB2CListingData()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbxAltEznYWjY5ULkbsGoi6RxAE5Tk8bLg_aqBhQ1dHvHZpaF3NstWt6xJgEfh00BjH-HQ/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Ebay Apps Script
    public function fetchEbayListingData()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/a/macros/5core.com/s/AKfycbzq7k35vKiGMfBrhO-7_gGNja1am_H0zBQ7xTjcfTfABaf4Q-lTVlui_-x8KzLHksX0sw/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }









    public function fetchDataFromKwEbayGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbwXkAyXFDAtCnvMtzzqsVk1vqlWZAacPZOmuQso0vkUVOGAuMqhyKaJYwisqof_55eizg/exec?route=ebay';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    public function fetchDataFromKwWalmartGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbwXkAyXFDAtCnvMtzzqsVk1vqlWZAacPZOmuQso0vkUVOGAuMqhyKaJYwisqof_55eizg/exec?route=walmart';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    public function fetchDataFromKwAmazonGoogleSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbw7RpyRwIc5B3xC-k7bVt_1h3aro9J2z-XaZYAd53R_DE8S6seZcnCs8FVlm8KOCibBRw/exec?route=amazon';


        try {
            $response = Http::timeout(120)->get($url);

            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'Amazon KW data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch Amazon KW data. Response:', [$response->body()]);

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching Amazon KW data:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }


    public function fetchDataFromGoogleShoppingGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbwXkAyXFDAtCnvMtzzqsVk1vqlWZAacPZOmuQso0vkUVOGAuMqhyKaJYwisqof_55eizg/exec?route=googleShopping';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Amazon Apps Script
    public function fetchDataFromAmazonGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbx51o_K6TjYvs-mYtXq1_B_OGu_ojxRCErdWD063GK5lCe1siZGwREvnNitTEK46vKh/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();
                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Update Amazon column in Google Sheet
    public function updateAmazonColumn(Request $request)
    {
        try {
            // Validate request
            $validatedData = $request->validate([
                'slNo' => 'required|integer',
                'updates' => 'required|array',
            ]);

            // Prepare final data format 
            $data = [
                'task' => 'update_by_slno',
                'data' => array_merge(
                    ['SL NO' => $validatedData['slNo']],
                    $validatedData['updates']
                ),
            ];

            // endpoint
            $url = 'https://script.google.com/macros/s/AKfycbx51o_K6TjYvs-mYtXq1_B_OGu_ojxRCErdWD063GK5lCe1siZGwREvnNitTEK46vKh/exec';

            // Post request to Google Apps Script
            $response = Http::timeout(120)->post($url, $data);

            if ($response->successful()) {
                return response()->json([
                    'message' => 'Data updated successfully',
                    'data' => $response->json(),
                    'status' => 200
                ]);
            } else {
                return response()->json([
                    'message' => 'Failed to update Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            return response()->json([
                'message' => 'An error occurred while updating data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }


    // Fetch data from Amazon Apps Script
    public function fetchDataFromAmazonFBAGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbzWwqRpTmb8eq0Vp05kP63r02smPIWGsTdcNozqIH0kERoLWuhtTcrsSv4KEub8oeoLNw/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Update AmazonFBA column in Google Sheet
    public function updateAmazonFBAColumn(Request $request)
    {
        try {
            // Log received data for debugging
            // Log::info('Received request:', $request->all());

            // Validate request
            $validatedData = $request->validate([
                'slNo' => 'required|integer',
                'updates' => 'required|array',
                'updates.*' => 'required|string',
            ]);

            // Google Apps Script API URL
            $url = 'https://script.google.com/macros/s/AKfycbzWwqRpTmb8eq0Vp05kP63r02smPIWGsTdcNozqIH0kERoLWuhtTcrsSv4KEub8oeoLNw/exec';


            // Send request to Google Apps Script
            $response = Http::timeout(120)->post($url, $validatedData);

            // Check if the request was successful
            if ($response->successful()) {
                // Log::info('Data updated successfully:', $response->json());
                return response()->json([
                    'message' => 'Data updated successfully',
                    'data' => $response->json(),
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to update:', $response->body());
                return response()->json([
                    'message' => 'Failed to update Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception occurred:', ['error' => $e->getMessage()]);
            return response()->json([
                'message' => 'An error occurred while updating data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Update Ebay column in Google Sheet
    public function updateEbay2Column(Request $request)
    {
        try {
            // Validate request
            $validatedData = $request->validate([
                'slNo' => 'required|integer',
                'updates' => 'required|array',
            ]);

            // Prepare final data format 
            $data = [
                'task' => 'update_by_slno',
                'data' => array_merge(
                    ['Sl' => $validatedData['slNo']],
                    $validatedData['updates']
                ),
            ];

            // Google Apps Script API URL
            $url = 'https://script.google.com/macros/s/AKfycbw9lTqutA_Ndu-Kha29ZFxKzCFjJreklNTdjNxRZHDAJZayX2XO_Ss1WQraO6-ZhKY8/exec';


            // Post request to Google Apps Script
            $response = Http::timeout(120)->post($url, $data);

            if ($response->successful()) {
                return response()->json([
                    'message' => 'Data updated successfully',
                    'data' => $response->json(),
                    'status' => 200
                ]);
            } else {
                return response()->json([
                    'message' => 'Failed to update Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            return response()->json([
                'message' => 'An error occurred while updating data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    public function updateEbayColumn(Request $request)
    {
        try {
            // Validate request
            $validatedData = $request->validate([
                'slNo' => 'required|integer',
                'updates' => 'required|array',
            ]);

            // Prepare final data format 
            $data = [
                'task' => 'update_by_slno',
                'data' => array_merge(
                    ['Sl' => $validatedData['slNo']],
                    $validatedData['updates']
                ),
            ];

            // Google Apps Script API URL
            $url = 'https://script.google.com/macros/s/AKfycbzq7k35vKiGMfBrhO-7_gGNja1am_H0zBQ7xTjcfTfABaf4Q-lTVlui_-x8KzLHksX0sw/exec';


            // Post request to Google Apps Script
            $response = Http::timeout(120)->post($url, $data);

            if ($response->successful()) {
                return response()->json([
                    'message' => 'Data updated successfully',
                    'data' => $response->json(),
                    'status' => 200
                ]);
            } else {
                return response()->json([
                    'message' => 'Failed to update Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            return response()->json([
                'message' => 'An error occurred while updating data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Update shopify b2c column in Google Sheet
    public function updateShopifyB2CColumn(Request $request)
    {
        try {
            // Log received data for debugging
            Log::info('Received request:', $request->all());

            // Validate request
            $validatedData = $request->validate([
                'slNo' => 'required|integer',
                'updates' => 'required|array',
            ]);

            // Prepare final data format 
            $data = [
                'task' => 'update_by_slno',
                'data' => array_merge(
                    ['SL' => $validatedData['slNo']],
                    $validatedData['updates']
                ),
            ];

            // Google Apps Script API URL
            $url = 'https://script.google.com/macros/s/AKfycbxAltEznYWjY5ULkbsGoi6RxAE5Tk8bLg_aqBhQ1dHvHZpaF3NstWt6xJgEfh00BjH-HQ/exec';

            // Post request to Google Apps Script
            $response = Http::timeout(120)->post($url, $data);

            if ($response->successful()) {
                return response()->json([
                    'message' => 'Data updated successfully',
                    'data' => $response->json(),
                    'status' => 200
                ]);
            } else {
                return response()->json([
                    'message' => 'Failed to update Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            return response()->json([
                'message' => 'An error occurred while updating data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Store data
    public function storeData(Request $request)
    {
        $request->validate([
            'name' => 'required|string',
            'email' => 'required|email|unique:users',
        ]);

        $user = User::create([
            'name' => $request->name,
            'email' => $request->email
        ]);

        return response()->json([
            'message' => 'User Created Successfully',
            'user' => $user
        ], 201);
    }

    // Fetch data from external API
    public function fetchExternalData()
    {
        // Google Apps Script Web App URL
        $apiUrl = "https://script.googleusercontent.com/macros/echo?user_content_key=AehSKLiY9qL7crl0zj-LUKtEYcZzwxH_YZsb28LBkpB8eGiEk1FgHz-R9OayvPXVVrD3orbmmfBv1g60z7Vt7lwJ8YjKUMRrp7QhB4sd_ZU9gci1mu_kp1teJAx3uUJs8qJeiq6XjSOaBAIYY-LdubLg2u-dZlwrb96xnQIh7198eTaXtv-a5oTFfBsrGx338_SMp_UTGeAxip22bjmkM0Z60_qPK__k-GCOF_3oPrMnOuk6j-kuik3pxF0z0cXGBO8Itai5fLDjtis9j1HVs7_f32tCnCEfRg&lib=MJAG7bh-wNYBTHeoOP4Nr2_btUGP9QdF0";

        // ✅ Fetch data from Google Sheets API
        $response = Http::get($apiUrl);

        // ✅ Check if request was successful
        if ($response->successful()) {
            return response()->json($response->json());
        }

        return response()->json(['error' => 'Failed to fetch data'], 500);
    }

    public function fetchExternalData2()
    {
        try {
            $response = Http::get('https://script.googleusercontent.com/macros/echo?user_content_key=AehSKLgVificg1JU3-Jj937ixAuRikG5IrdbWhgPyj9toMYalxisNyiuca1Ei_TJFna55bXRhMUoblof7YY0LhGhYzRV9R_bNsIBjG5Ma2ZRNaWBiTfvIw0-SSIGPym3SsiEXqFSPjthhXTkqY24SHXmp_qxns9CsBwJnec46Py8VjP47GVbyqH53cI_EhhptFLDNbCJCH0iGfbkffxRr9lJTqQQczIZ9Ye458SVm6Q5sNYq56V5c96L7abf6jdNBZ22KZHfwzpZjocG5zOzSHX7D8MM0VQkbAJ7_R0mWJOP&lib=M_0Jj1VeQKN9QEjfrDsIi_23vj4aZjNIa');

            if ($response->successful()) {
                $data = $response->json();
                Log::info('Fetched data from external API', ['data' => $data]); // Log the fetched data
                return $data;
            }

            Log::error('Failed to fetch data from external API', [
                'status' => $response->status(),
                'response' => $response->body(),
            ]);
        } catch (\Exception $e) {
            Log::error('Exception occurred while fetching data', [
                'message' => $e->getMessage(),
            ]);
        }

        return []; // Return an empty array if no data is fetched
    }

    // Fetch data from listing master Apps Script
    public function fetchDataFromListingMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbxq-VcrvZzAyo8MsQLO0AhgHLuZFuwU-u1W1tGyd-9LssS_E3cabV4T0XX-acP_Rb9b/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Macy's Apps Script
    public function fetchMacyListingData()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbzEzP_1sM86PT512M3QFRYHDVJ-ZabnsczEh1_Eak8Eb1lZzZ4bX0yCGACxafEp8dbGng/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 180)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Update macy's column in Google Sheet
    public function updateMacyColumn(Request $request)
    {
        try {
            // Validate request
            $validatedData = $request->validate([
                'slNo' => 'required|integer',
                'updates' => 'required|array',
            ]);

            // Prepare final data format 
            $data = [
                'task' => 'update_by_slno',
                'data' => array_merge(
                    ['Sr no' => $validatedData['slNo']],
                    $validatedData['updates']
                ),
            ];

            // Google Apps Script API URL
            $url = 'https://script.google.com/macros/s/AKfycbzEzP_1sM86PT512M3QFRYHDVJ-ZabnsczEh1_Eak8Eb1lZzZ4bX0yCGACxafEp8dbGng/exec';

            // Post request to Google Apps Script
            $response = Http::timeout(120)->post($url, $data);

            if ($response->successful()) {
                return response()->json([
                    'message' => 'Data updated successfully',
                    'data' => $response->json(),
                    'status' => 200
                ]);
            } else {
                return response()->json([
                    'message' => 'Failed to update Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            return response()->json([
                'message' => 'An error occurred while updating data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from product master Apps Script
    public function fetchDataFromProductMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbzUsXOfhxZP49uTYfbxMBcx-NEr9TCPa8yQLJRUvQkxc10hZDzqM27XX0F_4EaDLHRlgQ/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 180)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Channel master Apps Script
    public function fetchDataFromChannelMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbxoaCdVqZ6CzP2vWKpLwBxbnHrDatgjy6uvH6LocFW4H-TQcTJBIX5bvgfnK5W84Aby/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Newegg B2C master Apps Script
    public function fetchDataFromNeweggB2CMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbw9nqPPDupI2oD_JjD2UjpvSDVZlvIoPg7VjdsIlNu7udf0JnRjf29HTifUsRAfdexQ/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }





     public function fetchDataFromSheinMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbwNwG2rdvGOK49cRJQan7-3MSR2DQ2S-H0bP8iYx-olcfwWn_pswO-q7RS7hcZ152y5/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Newegg B2C master Apps Script
    public function fetchDataFromWayfairMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbxKeM4icpQCr4j--rmuI3129Ch6ThwugkqEvbBweEKdu6WyXkq-Hka65QZzi1tKMYNI/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }


    // Fetch data from Newegg B2C master Apps Script
    public function fetchDataFromFairMasterGoogleSheet()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbx0s7-4f77mapuYJ_HzUyJsubuAsXbzKyo95emiC3VdprmUvDXrWsbIPnaVozygyOb9iQ/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }





    // Fetch data from Ebay Apps Script
    public function fetchDobaListingData()
    {
        // URL of the Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbwSe2tvYfvb5_0uWK6YWumP7x6lpW90jtkL2DYEMhjMH6uNzJB27qjwEYdVe4QK3vHeIg/exec';

        try {
            // Make a GET request to the Google Apps Script URL
            $response = Http::timeout(seconds: 120)->get($url);

            // Check if the request was successful
            if ($response->successful()) {
                // Decode the JSON response
                $data = $response->json();

                // Log the data for debugging (optional)
                // Log::info('Data fetched from Google Sheet:', $data);

                // Return the data as a JSON response
                return response()->json([
                    'message' => 'Data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                // Log the error if the request failed
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                // Return an error response
                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            // Log the exception if something goes wrong
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            // Return an error response
            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }


    public function fetchDataFromLqsGoogleSheet()
    {
        // URL of the LQS Google Apps Script web app
        $url = 'https://script.google.com/macros/s/AKfycbyzws0BIKrOip0WyghgCUJL1UtiMPnuq9tnafHr10AytAbLlw7npRGi-5SSuA8iX2EZ/exec';

        try {
            // Make a GET request to the LQS Google Apps Script URL
            $response = Http::timeout(120)->get($url);

            if ($response->successful()) {
                $data = $response->json();
                // dd($data);

                return response()->json([
                    'message' => 'LQS data fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch LQS data from Google Sheet. Response:', ['body' => $response->body()]);

                return response()->json([
                    'message' => 'Failed to fetch LQS data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching LQS data:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching LQS data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    // Fetch data from Temu Apps Script

    public function fetchDataFromTemuListingDataSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbwu03jI3tT65UlW053kcWsQcx3rcv9qu3FdgnW0DVyUrfY7wsd7E2-ftIRmP8azC-hC0w/exec';

        try {
            $response = Http::timeout(seconds: 120)->get($url);
            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'Temu fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }

    public function fetchDataFromEbay3ListingDataSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbx5N-YTV8z7jJ51TqIewNtS5AFPHIZM35Q4-Tqyji0MH9bmQXfIQJYPVrBTO6kntJnO/exec';

        try {
            $response = Http::timeout(seconds: 120)->get($url);
            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'EBay3 fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }



    public function fetchDataFromEbay2ListingDataSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbw9lTqutA_Ndu-Kha29ZFxKzCFjJreklNTdjNxRZHDAJZayX2XO_Ss1WQraO6-ZhKY8/exec';

        try {
            $response = Http::timeout(seconds: 120)->get($url);
            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'Ebay2 fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }


    public function fetchDataFromWalmartListingDataSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbwT1m7qsRdrXuYJB-HLkOb0m6wmfytcsyeLNUGAWbjCzwPiX7Pcj1I7xROZVIFtTUcDOQ/exec';

        try {
            $response = Http::timeout(seconds: 120)->get($url);
            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'Walmart fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }


    public function syncInvAndL30ToSheet()
    {
        // Fetch latest data from your database or logic used in DataTable
        $data = ShopifySku::select('sku', 'inv', 'quantity')->get();

        $formatted = $data->map(function ($item) {
            return [
                'SKU' => $item->sku,
                'INV' => $item->inv,
                'L30' => $item->quantity,
            ];
        });

        $payload = [
            'task' => 'bulk_update_inv_l30',
            'data' => $formatted->values()->all()
        ];

        $url = 'https://script.google.com/macros/s/AKfycbxTfqbIQtcSpvxNsXnbkjH-xDwk4kPYX_aTjBP39mvIhHDtvrk8paUCC9BAT25byM9D/exec';
        

        try {
            $response = Http::timeout(60)->post($url, $payload);

            if ($response->successful()) {
                Log::info('Sync successful:');
                Log::info(json_encode($formatted->values()->all()));

                // if ($response->successful()) {
                return response()->json(['message' => 'Sync successful', 'data' => $response->json()]);
            } else {
                Log::error('Sync failed:', ['response' => $response->body()]);
                return response()->json(['message' => 'Google Sheet sync failed', 'error' => $response->body()], 500);
            }
        } catch (\Exception $e) {
            Log::error('Request error:', ['error' => $e->getMessage()]);
            return response()->json(['message' => 'Request error', 'error' => $e->getMessage()], 500);
        }
    }



     public function fetchDataSheetListingDataSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbwVPMyz2x4Np4kyd3ejFHtPdkeaPRg2sJdOk9TMGbDqdn6puRVVtQ9tQrIsYb0hYZIV/exec';

        try {
            $response = Http::timeout(seconds: 120)->get($url);
            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'Walmart fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }



     public function fetchDataFromTiktokDataSheet()
    {
        $url = 'https://script.google.com/macros/s/AKfycbyj1Z0xGDKHOWZvqj1fdnBi02abq67NzwBc7fj0XckA9O3zGbZOyHnLLDXuOPnTLC3E/exec';

        try {
            $response = Http::timeout(seconds: 120)->get($url);
            if ($response->successful()) {
                $data = $response->json();

                return response()->json([
                    'message' => 'Tiktok fetched successfully',
                    'data' => $data,
                    'status' => 200
                ]);
            } else {
                Log::error('Failed to fetch data from Google Sheet. Response:', $response->body());

                return response()->json([
                    'message' => 'Failed to fetch data from Google Sheet',
                    'status' => $response->status()
                ], $response->status());
            }
        } catch (\Exception $e) {
            Log::error('Exception while fetching data from Google Sheet:', ['error' => $e->getMessage()]);

            return response()->json([
                'message' => 'An error occurred while fetching data',
                'error' => $e->getMessage(),
                'status' => 500
            ], 500);
        }
    }
}
