<?php

namespace App\Services;

use GuzzleHttp\Client;
use GuzzleHttp\Psr7\Request;
use Aws\Signature\SignatureV4;
use Aws\Credentials\Credentials;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class FaireService
{
    protected $clientId;
    protected $clientSecret;
    protected $refreshToken;
    protected $region;
    protected $marketplaceId;
    protected $awsAccessKey;
    protected $awsSecretKey;
    protected $endpoint;

    public function __construct()
    {
        $this->clientId     = env('FAIRE_APP_ID');
        $this->clientSecret = env('FAIRE_APP_SECRET');
        $this->redirectUrl  = env('FAIRE_REDIRECT_URL');
    }

    public function getInventory(){

    } 

}
