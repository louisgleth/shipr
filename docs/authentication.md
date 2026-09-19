# Authentication

Create and manage API keys in **Portal > Account > Developer**. Give each integration its own key so you can revoke one without interrupting others.

```http
Authorization: Bearer shipide_test_YOUR_KEY
```

`shipide_test_` keys access sandbox records. `shipide_live_` keys access live records. Both use the same API base URL. One customer's keys cannot access another customer's records, and sandbox records are isolated from live records.

The full key is shown only at creation. Shipide stores its SHA-256 hash, not the plaintext key. You can set an expiry when creating it; expired and revoked keys receive `401 invalid_api_key`.

## Permissions

| Scope | Access |
| --- | --- |
| `account:read` | Account identity, environment, and capabilities |
| `orders:read` | List and retrieve orders |
| `orders:write` | Create orders |
| `shipments:read` | List and retrieve shipments; retrieve tracking |
| `shipments:write` | Create shipments |
| `rates:read` | Request rates |
| `labels:read` | List and retrieve labels; download PDFs |
| `labels:write` | Create and void labels |
| `webhooks:read` | List and retrieve webhook subscriptions |
| `webhooks:write` | Create and disable webhook subscriptions |

Missing permissions return `403 insufficient_scope`. At most 20 active keys can be created per account. Key management requires a signed-in portal session; an API key cannot create more keys.

## Rotate a key

Create a replacement with the necessary scopes, update your backend secret, verify a request, then revoke the old key. Revocation prevents subsequent requests immediately. Never send a key in query parameters or browser-side JavaScript.
