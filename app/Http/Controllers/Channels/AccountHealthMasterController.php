<?php

namespace App\Http\Controllers\Channels;

use App\Http\Controllers\ApiController;
use App\Http\Controllers\Controller;
use App\Models\AccountHealthMaster;
use App\Models\AtoZClaimsRate;
use App\Models\ChannelMaster;
use App\Models\FullfillmentRate;
use App\Models\LateShipmentRate;
use App\Models\NegativeSellerRate;
use App\Models\OdrRate;
use App\Models\OnTimeDeliveryRate;
use App\Models\RefundRate;
use App\Models\ValidTrackingRate;
use App\Models\VoilanceRate;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Log;

class AccountHealthMasterController extends Controller
{
    protected $apiController;

    public function __construct(ApiController $apiController)
    {
        $this->apiController = $apiController;
    }

    public function test()
    {
        $response = $this->apiController->fetchDataFromChannelMasterGoogleSheet();
        if ($response->getStatusCode() === 200) {
            $apiData = $response->getData()->data;

            // foreach ($apiData as $row) {
            //     $channelName = trim($row->{'Channel'} ?? $row->{'Channel '} ?? null);

            //     if (!$channelName) continue;

            //     AccountHealthMaster::create([
            //         'channel' => $channelName,
            //         'l30_sales' => $row->{'L30 Sales'} ?? null,
            //         'l30_orders' => $row->{'L30 Orders'} ?? null,
            //         'account_health_links' => $row->{'Health'} ?? null,
            //         'created_by' => Auth::user()->id,
            //         'report_date' => now(),
            //     ]);
            // }
            echo "<pre>";
            print_r($apiData);
            return response()->json(['success' => true, 'message' => 'Data inserted from Google Sheet.']);
        }
        return response()->json(['success' => false, 'message' => 'Failed to fetch data.'], 500);
    }

    public function index()
    {
        $channels = ChannelMaster::all();
        $accountHealthMaster = AccountHealthMaster::all();
        return view('channels.account-health-master', [
            'accountHealthMaster' => $accountHealthMaster,
            'channels' => $channels,
        ]);
    }

    public function store(Request $request)
    {
        $request->validate([
            'channel' => 'required',
            'report_date' => 'required',
        ], [
            'channel.required' => 'Please select channel name',
        ]);

        AccountHealthMaster::create([
            'channel' => $request->channel,
            'l30_sales' => null,
            'l30_orders' => null,
            'account_health_links' => $request->account_health_links,
            'remarks' => $request->remarks,
            'pre_fulfillment_cancel_rate' => $request->pre_fulfillment_cancel_rate,
            'odr' => $request->odr_transaction_defect_rate,
            'fulfillment_rate' => $request->fulfillment_rate,
            'late_shipment_rate' => $request->late_shipment_rate,
            'valid_tracking_rate' => $request->valid_tracking_rate,
            'on_time_delivery_rate' => $request->on_time_delivery_rate,
            'negative_feedback' => $request->negative_feedback,
            'positive_feedback' => $request->positive_feedback,
            'guarantee_claims' => $request->guarantee_claims,
            'refund_rate' => $request->refund_rate,
            'avg_processing_time' => $request->avg_processing_time,
            'message_time' => $request->message_time,
            'overall' => null,
            'report_date' => $request->report_date,
            'created_by' => Auth::user()->id,
        ]);

        return redirect()->back()->with('success', 'Account Health Report saved successfully.');
    }

    public function updateLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        AccountHealthMaster::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    public function update(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
        ]);

        $health = AccountHealthMaster::find($request->id);

        $health->update($request->only([
            'l30_sales',
            'l30_orders',
            'remarks',
            'pre_fulfillment_cancel_rate',
            'odr',
            'fulfillment_rate',
            'late_shipment_rate',
            'valid_tracking_rate',
            'on_time_delivery_rate',
            'negative_feedback',
            'positive_feedback',
            'guarantee_claims',
            'refund_rate',
            'avg_processing_time',
            'message_time',
            'overall',
        ]));

        return response()->json(['success' => true, 'message' => 'Updated successfully']);
    }

    // odr rate master start
    public function odrRateIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.odr-rate', compact('channels'));
    }

    public function saveOdrRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        OdrRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchOdrRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();

        $odrRates = OdrRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateOdrRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        $odr = OdrRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($odr) {
            // Shift previous data
            $odr->prev_2 = $odr->prev_1;
            $odr->prev_2_date = $odr->prev_1_date;

            $odr->prev_1 = $odr->current;
            $odr->prev_1_date = $odr->report_date;

            // Update current with new data
            $odr->current = $request->current ?? $request->odr_rate;
            $odr->report_date = $nowDate;

            $odr->allowed = $request->allowed ?? $request->odr_rate_allowed;
            $odr->what = $request->what;
            $odr->why = $request->why;
            $odr->action = $request->action;
            $odr->c_action = $request->c_action;
            $odr->account_health_links = $request->account_health_links;

            $odr->save();
        } else {
            $odr = OdrRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current,
                'allowed' => $request->allowed,
                'what' => $request->what,
                'why' => $request->why,
                'action' => $request->action,
                'c_action' => $request->c_action,
                'account_health_links' => $request->account_health_links,
            ]);
        }

        return response()->json(['message' => 'ODR Rate updated successfully']);
    }

    public function updateOdrHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        OdrRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }
    // odr rate master end

    // fullfillment rate start
    public function fullfillmentRateIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.fullfillment-rate', compact('channels'));
    }

    public function saveFullfillmentRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        FullfillmentRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchFullfillmentRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();

        $odrRates = FullfillmentRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateFullfillmentRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        if (!$channelId) {
            return response()->json(['message' => 'Channel not found'], 404);
        }

        $fulfillment = FullfillmentRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($fulfillment) {
            // Only shift data if we're updating the current value
            if ($request->has('current') || $request->has('fulfillment_rate')) {
                $fulfillment->prev_2 = $fulfillment->prev_1;
                $fulfillment->prev_2_date = $fulfillment->prev_1_date;

                $fulfillment->prev_1 = $fulfillment->current;
                $fulfillment->prev_1_date = $fulfillment->report_date;

                // Update current with new data
                $fulfillment->current = $request->current ?? $request->fulfillment_rate;
                $fulfillment->report_date = $nowDate;
            }

            // Only update allowed if it's present in the request
            if ($request->has('allowed') || $request->has('fulfillment_rate_allowed')) {
                $fulfillment->allowed = $request->allowed ?? $request->fulfillment_rate_allowed;
            }

            // Only update other fields if they're present in the request
            if ($request->has('what')) {
                $fulfillment->what = $request->what;
            }
            if ($request->has('why')) {
                $fulfillment->why = $request->why;
            }
            if ($request->has('action')) {
                $fulfillment->action = $request->action;
            }
            if ($request->has('c_action')) {
                $fulfillment->c_action = $request->c_action;
            }
            if ($request->has('account_health_links')) {
                $fulfillment->account_health_links = $request->account_health_links;
            }

            $fulfillment->save();
        } else {
            $fulfillment = FullfillmentRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current ?? $request->fulfillment_rate ?? null,
                'allowed' => $request->allowed ?? $request->fulfillment_rate_allowed ?? null,
                'what' => $request->what ?? null,
                'why' => $request->why ?? null,
                'action' => $request->action ?? null,
                'c_action' => $request->c_action ?? null,
                'account_health_links' => $request->account_health_links ?? null,
            ]);
        }

        return response()->json(['success' => true, 'message' => 'Fullfillment Rate updated successfully']);
    }

    public function updateFullfillmentHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        FullfillmentRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // fullfillment rate end

    // validTracking rate start
    public function validTrackingRateIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.valid-tracking-rate', compact('channels'));
    }

    public function saveValidTrackingRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        ValidTrackingRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchValidTrackingRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = ValidTrackingRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateValidTrackingRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        if (!$channelId) {
            return response()->json(['message' => 'Channel not found'], 404);
        }

        $validTracking = ValidTrackingRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($validTracking) {
            // Only shift data if we're updating the current value
            if ($request->has('current') || $request->has('valid_tracking_rate')) {
                $validTracking->prev_2 = $validTracking->prev_1;
                $validTracking->prev_2_date = $validTracking->prev_1_date;

                $validTracking->prev_1 = $validTracking->current;
                $validTracking->prev_1_date = $validTracking->report_date;

                // Update current with new data
                $validTracking->current = $request->current ?? $request->valid_tracking_rate;
                $validTracking->report_date = $nowDate;
            }

            // Only update allowed if it's present in the request
            if ($request->has('allowed') || $request->has('valid_tracking_rate_allowed')) {
                $validTracking->allowed = $request->allowed ?? $request->valid_tracking_rate_allowed;
            }

            // Only update other fields if they're present in the request
            if ($request->has('what')) {
                $validTracking->what = $request->what;
            }
            if ($request->has('why')) {
                $validTracking->why = $request->why;
            }
            if ($request->has('action')) {
                $validTracking->action = $request->action;
            }
            if ($request->has('c_action')) {
                $validTracking->c_action = $request->c_action;
            }
            if ($request->has('account_health_links')) {
                $validTracking->account_health_links = $request->account_health_links;
            }

            $validTracking->save();
        } else {
            $validTracking = ValidTrackingRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current ?? $request->valid_tracking_rate ?? null,
                'allowed' => $request->allowed ?? $request->valid_tracking_rate_allowed ?? null,
                'what' => $request->what ?? null,
                'why' => $request->why ?? null,
                'action' => $request->action ?? null,
                'c_action' => $request->c_action ?? null,
                'account_health_links' => $request->account_health_links ?? null,
            ]);
        }

        return response()->json(['success' => true, 'message' => 'Valid Tracking Rate updated successfully']);
    }

    public function updateValidTrackingHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        ValidTrackingRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // validTracking rate end

    // lateShipment rate start
    public function lateShipmentRateIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.late-shipment', compact('channels'));
    }

    public function saveLateShipmentRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        LateShipmentRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchLateShipmentRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = LateShipmentRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateLateShipmentRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        $odr = LateShipmentRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($odr) {
            // Shift previous data
            $odr->prev_2 = $odr->prev_1;
            $odr->prev_2_date = $odr->prev_1_date;

            $odr->prev_1 = $odr->current;
            $odr->prev_1_date = $odr->report_date;

            // Update current with new data
            $odr->current = $request->current;
            $odr->report_date = $nowDate;

            $odr->allowed = $request->allowed;
            $odr->what = $request->what;
            $odr->why = $request->why;
            $odr->action = $request->action;
            $odr->c_action = $request->c_action;
            $odr->account_health_links = $request->account_health_links;

            $odr->save();
        } else {
            $odr = LateShipmentRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current,
                'allowed' => $request->allowed,
                'what' => $request->what,
                'why' => $request->why,
                'action' => $request->action,
                'c_action' => $request->c_action,
                'account_health_links' => $request->account_health_links,
            ]);
        }

        return response()->json(['message' => 'ODR Rate updated successfully']);
    }

    public function updateLateShipmentHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        LateShipmentRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // lateShipment rate end

    // onTimeDelivery rate start
    public function onTimeDeliveryIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.on-time-delivery', compact('channels'));
    }

    public function saveOnTimeDeliveryRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        OnTimeDeliveryRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchOnTimeDeliveryRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = OnTimeDeliveryRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateOnTimeDeliveryRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        if (!$channelId) {
            return response()->json(['message' => 'Channel not found'], 404);
        }

        $onTimeDelivery = OnTimeDeliveryRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($onTimeDelivery) {
            // Only shift data if we're updating the current value
            if ($request->has('current') || $request->has('on_time_delivery')) {
                $onTimeDelivery->prev_2 = $onTimeDelivery->prev_1;
                $onTimeDelivery->prev_2_date = $onTimeDelivery->prev_1_date;

                $onTimeDelivery->prev_1 = $onTimeDelivery->current;
                $onTimeDelivery->prev_1_date = $onTimeDelivery->report_date;

                // Update current with new data
                $onTimeDelivery->current = $request->current ?? $request->on_time_delivery;
                $onTimeDelivery->report_date = $nowDate;
            }

            // Only update allowed if it's present in the request
            if ($request->has('allowed') || $request->has('on_time_delivery_allowed')) {
                $onTimeDelivery->allowed = $request->allowed ?? $request->on_time_delivery_allowed;
            }

            // Only update other fields if they're present in the request
            if ($request->has('what')) {
                $onTimeDelivery->what = $request->what;
            }
            if ($request->has('why')) {
                $onTimeDelivery->why = $request->why;
            }
            if ($request->has('action')) {
                $onTimeDelivery->action = $request->action;
            }
            if ($request->has('c_action')) {
                $onTimeDelivery->c_action = $request->c_action;
            }
            if ($request->has('account_health_links')) {
                $onTimeDelivery->account_health_links = $request->account_health_links;
            }

            $onTimeDelivery->save();
        } else {
            $onTimeDelivery = OnTimeDeliveryRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current ?? $request->on_time_delivery ?? null,
                'allowed' => $request->allowed ?? $request->on_time_delivery_allowed ?? null,
                'what' => $request->what ?? null,
                'why' => $request->why ?? null,
                'action' => $request->action ?? null,
                'c_action' => $request->c_action ?? null,
                'account_health_links' => $request->account_health_links ?? null,
            ]);
        }

        return response()->json(['success' => true, 'message' => 'On Time Delivery Rate updated successfully']);
    }

    public function updateOnTimeDeliveryHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        OnTimeDeliveryRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // onTimeDelivery rate end

    // negativeSeller rate start
    public function negativeSellerIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.negative-seller', compact('channels'));
    }

    public function saveNegativeSellerRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        NegativeSellerRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchNegativeSellerRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = NegativeSellerRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateNegativeSellerRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        $odr = NegativeSellerRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($odr) {
            // Shift previous data
            $odr->prev_2 = $odr->prev_1;
            $odr->prev_2_date = $odr->prev_1_date;

            $odr->prev_1 = $odr->current;
            $odr->prev_1_date = $odr->report_date;

            // Update current with new data
            $odr->current = $request->current;
            $odr->report_date = $nowDate;

            $odr->allowed = $request->allowed;
            $odr->what = $request->what;
            $odr->why = $request->why;
            $odr->action = $request->action;
            $odr->c_action = $request->c_action;
            $odr->account_health_links = $request->account_health_links;

            $odr->save();
        } else {
            $odr = NegativeSellerRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current,
                'allowed' => $request->allowed,
                'what' => $request->what,
                'why' => $request->why,
                'action' => $request->action,
                'c_action' => $request->c_action,
                'account_health_links' => $request->account_health_links,
            ]);
        }

        return response()->json(['message' => 'ODR Rate updated successfully']);
    }

    public function updateNegativeSellerHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        NegativeSellerRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // negativeSeller rate end

    // a-z-Claims rate start
    public function aTozClaimsIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.a-z-claims', compact('channels'));
    }

    public function saveAtoZClaimsRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        AtoZClaimsRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchAtoZClaimsRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = AtoZClaimsRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateAtoZClaimsRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        if (!$channelId) {
            return response()->json(['message' => 'Channel not found'], 404);
        }

        $atozClaims = AtoZClaimsRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($atozClaims) {
            // Only shift data if we're updating the current value
            if ($request->has('current') || $request->has('atoz_claims_rate')) {
                $atozClaims->prev_2 = $atozClaims->prev_1;
                $atozClaims->prev_2_date = $atozClaims->prev_1_date;

                $atozClaims->prev_1 = $atozClaims->current;
                $atozClaims->prev_1_date = $atozClaims->report_date;

                // Update current with new data
                $atozClaims->current = $request->current ?? $request->atoz_claims_rate;
                $atozClaims->report_date = $nowDate;
            }

            // Only update allowed if it's present in the request
            if ($request->has('allowed') || $request->has('atoz_claims_rate_allowed')) {
                $atozClaims->allowed = $request->allowed ?? $request->atoz_claims_rate_allowed;
            }

            // Only update other fields if they're present in the request
            if ($request->has('what')) {
                $atozClaims->what = $request->what;
            }
            if ($request->has('why')) {
                $atozClaims->why = $request->why;
            }
            if ($request->has('action')) {
                $atozClaims->action = $request->action;
            }
            if ($request->has('c_action')) {
                $atozClaims->c_action = $request->c_action;
            }
            if ($request->has('account_health_links')) {
                $atozClaims->account_health_links = $request->account_health_links;
            }

            $atozClaims->save();
        } else {
            $atozClaims = AtoZClaimsRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current ?? $request->atoz_claims_rate ?? null,
                'allowed' => $request->allowed ?? $request->atoz_claims_rate_allowed ?? null,
                'what' => $request->what ?? null,
                'why' => $request->why ?? null,
                'action' => $request->action ?? null,
                'c_action' => $request->c_action ?? null,
                'account_health_links' => $request->account_health_links ?? null,
            ]);
        }

        return response()->json(['success' => true, 'message' => 'A-to-Z Claims Rate updated successfully']);
    }

    public function updateAtoZClaimsHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        AtoZClaimsRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // a-z-Claims rate end

    // voilation rate start
    public function voilationIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.voilation', compact('channels'));
    }

    public function saveVoilanceRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        VoilanceRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchVoilanceRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = VoilanceRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateVoilanceRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        $odr = VoilanceRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($odr) {
            // Shift previous data
            $odr->prev_2 = $odr->prev_1;
            $odr->prev_2_date = $odr->prev_1_date;

            $odr->prev_1 = $odr->current;
            $odr->prev_1_date = $odr->report_date;

            // Update current with new data
            $odr->current = $request->current;
            $odr->report_date = $nowDate;

            $odr->allowed = $request->allowed;
            $odr->what = $request->what;
            $odr->why = $request->why;
            $odr->action = $request->action;
            $odr->c_action = $request->c_action;
            $odr->account_health_links = $request->account_health_links;

            $odr->save();
        } else {
            $odr = VoilanceRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current,
                'allowed' => $request->allowed,
                'what' => $request->what,
                'why' => $request->why,
                'action' => $request->action,
                'c_action' => $request->c_action,
                'account_health_links' => $request->account_health_links,
            ]);
        }

        return response()->json(['message' => 'ODR Rate updated successfully']);
    }

    public function updateVoilanceHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        VoilanceRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // voilation rate end

    // refund rate start
    public function refundIndex()
    {
        $channels = ChannelMaster::all();
        return view('channels.account_health_master.refund', compact('channels'));
    }

    public function saveRefundRate(Request $request)
    {
        $request->validate([
            'channel_id' => 'required|string',
            'report_date' => 'required|date',
            'account_health_links' => 'nullable|string'
        ]);

        RefundRate::create([
            'channel_id' => $request->channel_id,
            'report_date' => $request->report_date,
            'account_health_links' => $request->account_health_links,
        ]);

        return redirect()->back()->with('flash_message', 'Report saved successfully.');
    }

    public function fetchRefundRates()
    {
        $excludedChannels = [
            'Shopify B2B',
            'Shopify B2C',
            'Newegg B2B',
            'Sears',
            'Zendrop',
            'Business 5Core',
            'Flea Market',
            'Music Schools',
            'Schools',
            'Sports Classes',
            'Sports Shops',
            'Marine Shops',
            'Instituional Sales',
            'Range Me',
            'Wholesale Central',
            'Wish'
        ];

        $channels = ChannelMaster::whereNotIn('channel', $excludedChannels)->get();
        $odrRates = RefundRate::with('channel')
            ->orderBy('report_date', 'desc')
            ->get()
            ->keyBy('channel_id');

        $formatted = $channels->map(function ($channel) use ($odrRates) {
            $odr = $odrRates[$channel->id] ?? null;

            return [
                'id' => $odr?->id ?? null,
                'channel' => $channel->channel,
                'allowed' => $odr?->allowed ?? '',
                'current' => $odr?->current ?? '',
                'report_date' => $odr?->report_date ?? '',
                'prev_1' => $odr?->prev_1 ?? '',
                'prev_1_date' => $odr?->prev_1_date ?? '',
                'prev_2' => $odr?->prev_1 ?? '',
                'prev_2_date' => $odr?->prev_2_date ?? '',
                'what' => $odr?->what ?? '',
                'why' => $odr?->why ?? '',
                'action' => $odr?->action ?? '',
                'c_action' => $odr?->c_action ?? '',
                'account_health_links' => $odr?->account_health_links ?? '',
            ];
        });

        return response()->json($formatted);
    }

    public function updateRefundRate(Request $request)
    {
        $channelId = ChannelMaster::where('channel', $request->channel)->value('id');

        $odr = RefundRate::where('channel_id', $channelId)->first();

        $nowDate = now()->toDateString();

        if ($odr) {
            // Shift previous data
            $odr->prev_2 = $odr->prev_1;
            $odr->prev_2_date = $odr->prev_1_date;

            $odr->prev_1 = $odr->current;
            $odr->prev_1_date = $odr->report_date;

            // Update current with new data
            $odr->current = $request->current;
            $odr->report_date = $nowDate;

            $odr->allowed = $request->allowed;
            $odr->what = $request->what;
            $odr->why = $request->why;
            $odr->action = $request->action;
            $odr->c_action = $request->c_action;
            $odr->account_health_links = $request->account_health_links;

            $odr->save();
        } else {
            $odr = RefundRate::create([
                'channel_id' => $channelId,
                'report_date' => $nowDate,
                'current' => $request->current,
                'allowed' => $request->allowed,
                'what' => $request->what,
                'why' => $request->why,
                'action' => $request->action,
                'c_action' => $request->c_action,
                'account_health_links' => $request->account_health_links,
            ]);
        }

        return response()->json(['message' => 'ODR Rate updated successfully']);
    }

    public function updateRefundHealthLink(Request $request)
    {
        $request->validate([
            'id' => 'required|exists:account_health_master,id',
            'account_health_links' => 'required|url',
        ]);

        RefundRate::where('id', $request->id)->update([
            'account_health_links' => $request->account_health_links,
        ]);

        return response()->json(['success' => true, 'message' => 'Link updated']);
    }

    // refund rate end
}
