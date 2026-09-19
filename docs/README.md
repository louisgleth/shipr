---
description: Connect your warehouse, order system, or custom platform to Shipide.
---

# Shipide API

Send shipment details directly from your software to Shipide. Keep your existing order and warehouse workflow, and use the API to exchange shipment records, labels, and events.

**Base URL:** `https://portal.shipide.com/api/v1`

**Version:** `v1` | **Format:** JSON | **Authentication:** Bearer API key

{% hint style="warning" %}
Live carrier booking is awaiting activation. Live keys currently accept order and shipment records only. Rates, labels, voids, and tracking are available in the sandbox; sandbox labels cannot be shipped and never incur charges. Check `GET /account` for current capabilities.
{% endhint %}

## The shipping flow

1. Your order system sends an order, or your warehouse creates a shipment directly.
2. Your warehouse supplies the packed parcel's weight and dimensions.
3. Request rates for the shipment, then create a label using a returned rate.
4. Download the PDF. Store Shipide's shipment ID and tracking number in your system.
5. Receive signed webhook events, or retrieve the shipment by ID.

Start with the [quickstart](quickstart.md), create a key in the [Developer page](https://portal.shipide.com/developer.html), or browse the [API reference](reference.md).

## Conventions

Weights are kilograms and dimensions are centimetres. Amounts use the stated currency, in major units: `5.00` means EUR 5.00. Addresses use two-letter country codes. Timestamps are ISO 8601 UTC. IDs are opaque strings; do not derive business information from them.

Every write requires `Idempotency-Key`. Keep it stable when retrying the same operation. Successful JSON responses use `data`; failures use `error` and include a `request_id` for support.
