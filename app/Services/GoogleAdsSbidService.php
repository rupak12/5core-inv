<?php

namespace App\Services;

use Google\Auth\Credentials\UserRefreshCredentials;
use Google\Ads\GoogleAds\Lib\V20\GoogleAdsClientBuilder;
use Google\Ads\GoogleAds\V20\Services\AdGroupOperation;
use Google\Ads\GoogleAds\V20\Services\AdGroupCriterionOperation;
use Google\Ads\GoogleAds\V20\Resources\AdGroup;
use Google\Ads\GoogleAds\V20\Resources\AdGroupCriterion;
use Google\Protobuf\FieldMask;
use Google\Ads\GoogleAds\V20\Services\SearchGoogleAdsStreamRequest;
use Google\Ads\GoogleAds\V20\Services\MutateAdGroupsRequest;
use Illuminate\Support\Facades\Log;
use Google\Ads\GoogleAds\V20\Services\MutateAdGroupCriteriaRequest;

class GoogleAdsSbidService
{
    protected $client;

    public function __construct()
    {
        $this->client = $this->buildClient();
    }

    private function buildClient()
    {
        $oAuth2Credential = new UserRefreshCredentials(
            ['https://www.googleapis.com/auth/adwords'],
            [
                'client_id'     => env('GOOGLE_ADS_CLIENT_ID'),
                'client_secret' => env('GOOGLE_ADS_CLIENT_SECRET'),
                'refresh_token' => env('GOOGLE_ADS_REFRESH_TOKEN'),
            ]
        );

        return (new GoogleAdsClientBuilder())
            ->withDeveloperToken(env('GOOGLE_ADS_DEVELOPER_TOKEN'))
            ->withLoginCustomerId(env('GOOGLE_ADS_LOGIN_CUSTOMER_ID'))
            ->withOAuth2Credential($oAuth2Credential)
            ->build();
    }

    /**
     * Run GAQL query
     */
    public function runQuery($customerId, $query)
    {
        $googleAdsService = $this->client->getGoogleAdsServiceClient();

        $request = new SearchGoogleAdsStreamRequest([
            'customer_id' => $customerId,
            'query' => $query,
        ]);

        $stream = $googleAdsService->searchStream($request);

        $results = [];
        foreach ($stream->iterateAllElements() as $row) {
            $results[] = json_decode($row->serializeToJsonString(), true);
        }

        return $results;
    }


    /**
     * Update Ad Group SBID
     */
    public function updateAdGroupSbid($customerId, $adGroupResourceName, $newSbid)
{
  
    $adGroupService = $this->client->getAdGroupServiceClient();

    $adGroup = new AdGroup([
        'resource_name' => $adGroupResourceName,
        'cpc_bid_micros' => intval($newSbid * 1_000_000)// Direct value, no conversion
    ]);

    $operation = new AdGroupOperation();
    $operation->setUpdate($adGroup);
    $operation->setUpdateMask(new FieldMask(['paths' => ['cpc_bid_micros']]));

    $request = new MutateAdGroupsRequest([
        'customer_id' => $customerId,
        'operations' => [$operation]
    ]);

    try {
        $response = $adGroupService->mutateAdGroups($request);
        Log::info("AdGroup SBID updated", [
            'customer_id' => $customerId,
            'ad_group' => $adGroupResourceName,
            'new_sbid' => $newSbid
        ]);
        return $response;
    } catch (\Exception $e) {
        Log::error("Failed to update AdGroup SBID", [
            'customer_id' => $customerId,
            'ad_group' => $adGroupResourceName,
            'new_sbid' => $newSbid,
            'error' => $e->getMessage()
        ]);
        throw $e; // optional: rethrow if you want cron to fail
    }
}
public function updateProductGroupSbid($customerId, $productGroupResourceName, $newSbid)
{
    $adGroupCriterionService = $this->client->getAdGroupCriterionServiceClient();

    $criterion = new AdGroupCriterion([
        'resource_name' => $productGroupResourceName,
       'cpc_bid_micros' => intval($newSbid * 1_000_000) 
    ]);

    $operation = new AdGroupCriterionOperation();
    $operation->setUpdate($criterion);
    $operation->setUpdateMask(new FieldMask(['paths' => ['cpc_bid_micros']]));

    // Wrap operation in a proper request object
    $request = new MutateAdGroupCriteriaRequest([
        'customer_id' => $customerId,
        'operations' => [$operation]
    ]);

    try {
        $response = $adGroupCriterionService->mutateAdGroupCriteria($request);

        \Log::info("Product Group SBID updated", [
            'customer_id' => $customerId,
            'product_group' => $productGroupResourceName,
            'new_sbid' => $newSbid
        ]);

        return $response;
    } catch (\Exception $e) {
        \Log::error("Failed to update Product Group SBID", [
            'customer_id' => $customerId,
            'product_group' => $productGroupResourceName,
            'new_sbid' => $newSbid,
            'error' => $e->getMessage()
        ]);
        throw $e;
    }
}
     public function updateCampaignSbids($customerId, $campaignId, $sbidFactor)
    {
        $adGroupQuery = "
            SELECT ad_group.resource_name, metrics.clicks, metrics.cost_micros
            FROM ad_group
            WHERE ad_group.campaign = 'customers/{$customerId}/campaigns/{$campaignId}'
        ";

        $adGroups = $this->runQuery($customerId, $adGroupQuery);

        foreach ($adGroups as $row) {
            $adGroup = $row['adGroup'] ?? [];
            $metrics = $row['metrics'] ?? [];
            $adGroupResource = $adGroup['resourceName'] ?? null;

            if ($adGroupResource) {
                $this->updateAdGroupSbid($customerId, $adGroupResource, $sbidFactor);
            }

            $productGroupQuery = "
                SELECT ad_group_criterion.resource_name, 
                       ad_group_criterion.listing_group.type,
                       ad_group_criterion.negative
                FROM ad_group_criterion
                WHERE ad_group_criterion.ad_group = '{$adGroupResource}'
                  AND ad_group_criterion.listing_group.type = 'UNIT'
                  AND ad_group_criterion.negative = FALSE
            ";


            $productGroups = $this->runQuery($customerId, $productGroupQuery);

            foreach ($productGroups as $pgRow) {
                $pgResource = $pgRow['adGroupCriterion']['resourceName'] ?? null;
                if ($pgResource) {
                    $this->updateProductGroupSbid($customerId, $pgResource, $sbidFactor);
                }
            }
        }
    }

}
