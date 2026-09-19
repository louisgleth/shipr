# Webhooks

Shipide sends events to your backend over HTTPS. Subscriptions belong to an account and environment: test events never reach live subscriptions.

## Subscribe

```http
POST /webhooks
Authorization: Bearer YOUR_API_KEY
Content-Type: application/json
Idempotency-Key: warehouse-events-v1

{
  "url": "https://warehouse.your-company.com/shipide/events",
  "events": ["order.created", "shipment.created", "label.created", "label.voided"]
}
```

The response includes a signing `secret` starting with `whsec_`. Store it securely. It is omitted from reads; retrying the identical create request with its original idempotency key returns the original response, including the secret.

Use a public HTTPS hostname on port 443. Credentials, URL fragments, localhost, and IP-literal destinations are rejected. Redirect responses are treated as failures. Up to 10 active subscriptions are supported per environment.

## Event format

```json
{
  "id": "evt_EXAMPLE",
  "type": "shipment.created",
  "livemode": false,
  "created_at": "2026-09-19T12:00:00.000Z",
  "data": {
    "id": "shp_EXAMPLE",
    "object": "shipment",
    "status": "draft"
  }
}
```

`data` contains the resource at the time of the event. The example is abbreviated. A shipment and its queued event are committed atomically, so a successful creation cannot lose its event before delivery.

## Verify the signature

Each delivery includes:

```http
Shipide-Event-Id: evt_EXAMPLE
Shipide-Signature: t=1790000000,v1=HEX_HMAC
```

Compute HMAC-SHA256 over `timestamp + "." + raw_request_body`, using the complete `whsec_...` secret as the signing key. Compare digests in constant time and reject timestamps older than five minutes. Use the exact request bytes, before parsing or reserializing JSON.

```javascript
import { createHmac, timingSafeEqual } from 'node:crypto';

export function verifyShipideWebhook(rawBody, header, secret) {
  const match = /^t=(\d+),v1=([a-f0-9]{64})$/.exec(header || '');
  if (!match) return false;
  const [, timestamp, signature] = match;
  if (Math.abs(Date.now() / 1000 - Number(timestamp)) > 300) return false;
  const expected = createHmac('sha256', secret)
    .update(`${timestamp}.`)
    .update(rawBody)
    .digest();
  const received = Buffer.from(signature, 'hex');
  return received.length === expected.length && timingSafeEqual(received, expected);
}
```

## Acknowledge and deduplicate

After verifying the signature, save or enqueue the event durably, then respond with any `2xx` status within ten seconds. Deduplicate by event ID. Delivery is at least once, and ordering is not guaranteed.

The dispatcher runs every minute. Failures retry after approximately 1, 2, 4, 8, 16, 32, and 60 minutes, up to eight total attempts. Delivery stops after the eighth failure. Network failures, timeouts, redirects, and non-2xx responses all count as failures. Backlogs can delay delivery beyond those intervals.

## Manage subscriptions

```http
GET /webhooks
GET /webhooks/{id}
DELETE /webhooks/{id}
```

Deleting requires `webhooks:write` and an `Idempotency-Key`; no body is needed. It disables future delivery. A request already in flight can still arrive. Listing and retrieval require `webhooks:read`. Failed-delivery inspection and manual replay are not yet exposed; retain your own event logs and retrieve the current resource if reconciliation is needed.
