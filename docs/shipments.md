# Orders and shipments

An **order** records the shipping request from your order system. A **shipment** records a packed parcel. You may create shipments directly without creating orders first.

## Create

`POST /orders` and `POST /shipments` accept the same shipping fields. Only shipments accept `order_id`, which optionally links to an existing order in the same account and environment.

| Field | Required | Meaning |
| --- | --- | --- |
| `reference` | Yes | Your order or warehouse reference, up to 200 characters |
| `from` | Yes | Sender address |
| `to` | Yes | Recipient address |
| `parcel` | Yes | Packed parcel's weight and dimensions |
| `order_id` | No | Existing order ID, for shipment creation only |
| `metadata` | No | Custom JSON object, at most 4096 characters when serialized |

Each address requires `name`, `street1`, `city`, `postal_code`, and `country`. Optional fields: `company`, `street2`, `state`, `email`, `phone`. Address values are strings up to 200 characters; `country` must be a two-letter code, for example `BE`, `NL`, or `FR`.

The parcel requires numeric `weight_kg`, `length_cm`, `width_cm`, and `height_cm`, all greater than zero and at most 1000. These are input limits, not a guarantee of carrier acceptance. Unknown fields are rejected to catch mapping mistakes.

Orders begin in `received` state. Shipments begin in `draft`, become `label_created` when a label is issued, and become `voided` when that label is voided. Each shipment can have one label. Create a new shipment to replace a voided label.

The API currently accepts one parcel per shipment. For an order with several parcels, create a shipment per parcel with the same order ID and distinct idempotency keys. Customs documents, multi-parcel carrier bookings, manifests, pickups, and returns are not yet available through this version.

## Retrieve and list

```http
GET /orders/{id}
GET /shipments/{id}
GET /orders?limit=25
GET /shipments?limit=25&after=shp_PREVIOUS_CURSOR
```

List responses contain `data` and `next_cursor`. Pass `next_cursor` as `after` until it is `null`. `limit` defaults to 25 and supports 1 to 100. Results are ordered by opaque ID, not chronologically. This is not a snapshot: concurrent inserts can appear before your cursor. Use webhooks for ongoing intake events rather than relying on list polling to discover every new record.

Records created through this API are retrieved through the API. They do not yet appear in the portal's legacy label-generation history or monthly billing, because live label booking has not been activated.
