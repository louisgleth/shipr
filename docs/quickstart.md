# Quickstart

This walkthrough creates a sandbox shipment and downloads a test PDF. It does not contact a carrier or charge your account.

## Create a sandbox key

Sign in to [Shipide](https://portal.shipide.com/developer.html), open **Developer**, and select **Create key**. Choose **Sandbox** and enable `shipments:write`, `shipments:read`, `rates:read`, `labels:write`, and `labels:read`. The full key is shown once. Store it in your backend's secret manager.

These shell examples use `curl` and `jq`. Set `SHIPIDE_API_KEY` to your sandbox key in your local environment. Do not put it in a website, mobile app, source repository, or URL.

## Create a shipment

```bash
export SHIPIDE_API_URL='https://portal.shipide.com/api/v1'

SHIPMENT=$(curl --fail-with-body -sS "$SHIPIDE_API_URL/shipments" \
  -H "Authorization: Bearer $SHIPIDE_API_KEY" \
  -H 'Content-Type: application/json' \
  -H 'Idempotency-Key: warehouse-order-1001-shipment' \
  -d '{
    "reference": "order-1001",
    "from": {
      "name": "Test warehouse",
      "street1": "Example street 1",
      "city": "Brussels",
      "postal_code": "1000",
      "country": "BE"
    },
    "to": {
      "name": "Test recipient",
      "street1": "Example street 2",
      "city": "Antwerp",
      "postal_code": "2000",
      "country": "BE"
    },
    "parcel": {
      "weight_kg": 1,
      "length_cm": 20,
      "width_cm": 15,
      "height_cm": 10
    }
  }')
SHIPMENT_ID=$(printf '%s' "$SHIPMENT" | jq -r '.data.id')
```

Successful creation returns `201 Created`. The shipment starts in `draft` state. Your reference is a business reference, not a deduplication key; use `Idempotency-Key` for deduplication.

## Get rates and create a label

```bash
RATES=$(curl --fail-with-body -sS "$SHIPIDE_API_URL/rates" \
  -H "Authorization: Bearer $SHIPIDE_API_KEY" \
  -H 'Content-Type: application/json' \
  -H 'Idempotency-Key: warehouse-order-1001-rates' \
  -d "$(jq -n --arg id "$SHIPMENT_ID" '{shipment_id:$id}')")
RATE_ID=$(printf '%s' "$RATES" | jq -r '.data[0].id')

LABEL=$(curl --fail-with-body -sS "$SHIPIDE_API_URL/labels" \
  -H "Authorization: Bearer $SHIPIDE_API_KEY" \
  -H 'Content-Type: application/json' \
  -H 'Idempotency-Key: warehouse-order-1001-label' \
  -d "$(jq -n --arg shipment "$SHIPMENT_ID" --arg rate "$RATE_ID" \
    '{shipment_id:$shipment,rate_id:$rate,format:"pdf"}')")
LABEL_ID=$(printf '%s' "$LABEL" | jq -r '.data.id')

curl --fail-with-body -sS "$SHIPIDE_API_URL/labels/$LABEL_ID/file" \
  -H "Authorization: Bearer $SHIPIDE_API_KEY" \
  -o sandbox-label.pdf
```

Sandbox rates are fixed test data. The returned rate is EUR 5.00, but the label's actual charge is always EUR 0.00. The PDF is marked **TEST LABEL - NOT VALID FOR SHIPPING**.

## Call Shipide from Node.js

```javascript
async function shipide(path, { body, idempotencyKey } = {}) {
  const response = await fetch(`https://portal.shipide.com/api/v1${path}`, {
    method: body ? 'POST' : 'GET',
    headers: {
      Authorization: `Bearer ${process.env.SHIPIDE_API_KEY}`,
      ...(body ? {
        'Content-Type': 'application/json',
        'Idempotency-Key': idempotencyKey,
      } : {}),
    },
    ...(body ? { body: JSON.stringify(body) } : {}),
  });
  const result = await response.json();
  if (!response.ok) {
    throw new Error(`${result.error.code}: ${result.error.message} (${result.error.request_id})`);
  }
  return result.data;
}

const shipment = await shipide(`/shipments/${shipmentId}`);
```

Next, implement [retry handling](errors.md) and [webhook verification](webhooks.md).
