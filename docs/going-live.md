# Going live

The sandbox can be used now to develop and test your integration contract. Live keys support order and shipment intake. Carrier rates, billable labels, real tracking events, and wallet or invoice settlement for API labels still require a live carrier integration.

## Before production shipping

1. Complete the sandbox flow: shipment, rate, label, download, and void.
2. Test a timeout retry with the same idempotency key and verify no duplicate resource is created.
3. Verify webhook signatures, save event IDs, and handle duplicate delivery.
4. Confirm the carrier, supported services, credentials, and negotiated rates with Shipide.
5. Verify `GET /account` reports the required live capabilities before enabling automated shipping.
6. Create a live key with the scopes your backend needs and a separate live webhook subscription.

Do not treat a `draft` shipment as a purchased label. Do not print sandbox PDFs on real parcels. Live label requests currently fail explicitly and do not debit a wallet or create an invoice.

## Warehouse integration

Send the measured package data at the packing station. Store the returned Shipide ID against your internal order or parcel. Once live booking is enabled, the same flow can return the carrier label to the packing station and the tracking number to your order system.

Integrate from a trusted backend, middleware service, or warehouse server. API keys must remain on that server. A browser can ask your backend to create and download labels without receiving the API key itself.
