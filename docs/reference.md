# API reference

Base URL: `https://portal.shipide.com/api/v1`

The [OpenAPI specification](openapi.json) describes the request schemas, response models, authentication, and endpoint permissions. It can be imported into Postman or another OpenAPI-compatible client.

| Method | Path | Required scope | Purpose |
| --- | --- | --- | --- |
| GET | `/account` | `account:read` | Identity and capabilities |
| POST | `/orders` | `orders:write` | Create an order |
| GET | `/orders` | `orders:read` | List orders |
| GET | `/orders/{id}` | `orders:read` | Retrieve an order |
| POST | `/shipments` | `shipments:write` | Create a shipment |
| GET | `/shipments` | `shipments:read` | List shipments |
| GET | `/shipments/{id}` | `shipments:read` | Retrieve a shipment |
| POST | `/rates` | `rates:read` | Request sandbox rates |
| POST | `/labels` | `labels:write` | Create a sandbox label |
| GET | `/labels` | `labels:read` | List labels |
| GET | `/labels/{id}` | `labels:read` | Retrieve a label |
| GET | `/labels/{id}/file` | `labels:read` | Download a PDF |
| POST | `/labels/{id}/void` | `labels:write` | Void a sandbox label |
| GET | `/tracking/{tracking_number}` | `shipments:read` | Retrieve sandbox tracking |
| POST | `/webhooks` | `webhooks:write` | Create a subscription |
| GET | `/webhooks` | `webhooks:read` | List subscriptions |
| GET | `/webhooks/{id}` | `webhooks:read` | Retrieve a subscription |
| DELETE | `/webhooks/{id}` | `webhooks:write` | Disable a subscription |

All POST and DELETE operations require an `Idempotency-Key`. All JSON bodies must use `Content-Type: application/json`. PDF downloads are binary; all other successful responses use a `data` envelope.
