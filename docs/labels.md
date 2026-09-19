# Rates and labels

{% hint style="warning" %}
These operations currently require a sandbox key. Live requests return `503 carrier_not_configured` without charging the account. Sandbox output is not carrier postage.
{% endhint %}

## Request rates

`POST /rates` requires `rates:read`, `Idempotency-Key`, and:

```json
{ "shipment_id": "shp_YOUR_SHIPMENT" }
```

It returns an array in `data`, containing a fixed sandbox rate:

```json
{
  "id": "rate_test_shp_YOUR_SHIPMENT",
  "shipment_id": "shp_YOUR_SHIPMENT",
  "carrier": "shipide_sandbox",
  "service": "sandbox_standard",
  "currency": "EUR",
  "amount": 5,
  "test": true
}
```

## Create a label

`POST /labels` requires `labels:write`, `Idempotency-Key`, and:

```json
{
  "shipment_id": "shp_YOUR_SHIPMENT",
  "rate_id": "rate_test_shp_YOUR_SHIPMENT",
  "format": "pdf"
}
```

Use the rate returned for that exact shipment. The response includes `id`, `shipment_id`, `tracking_number`, `download_path`, `status`, `format`, `test`, and `charge`. The sandbox charge is zero. Tracking numbers start with `TEST` and cannot be tracked through a carrier.

Only PDF is currently supported. The sandbox PDF is 4 by 6 inches and prominently marked as invalid for shipping. Non-ASCII address characters may be replaced on sandbox PDFs; carrier label rendering will be supplied by the live provider.

## Download or retrieve

```http
GET /labels/{id}
GET /labels/{id}/file
GET /labels?limit=25
```

Download requests still require the Bearer key. The file endpoint responds with `application/pdf`, not JSON. Proxy the download through your backend if it needs to reach an operator's browser; do not put your key in the download URL.

## Void

`POST /labels/{id}/void` accepts an empty JSON object `{}` and requires `labels:write` and `Idempotency-Key`. A voided label cannot be downloaded. Repeating a void with a new idempotency key returns the existing voided label without another state change.

## Tracking

`GET /tracking/{tracking_number}` requires `shipments:read`. Sandbox tracking returns `pre_transit` or `voided` with an empty events array. It does not simulate carrier movement, delivery scans, or estimated arrival times.
