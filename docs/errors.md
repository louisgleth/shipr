# Errors and retries

Errors use this structure:

```json
{
  "error": {
    "code": "invalid_request",
    "message": "Idempotency-Key must be a non-empty string of at most 128 characters.",
    "request_id": "req_EXAMPLE"
  }
}
```

Every API response includes `X-Request-Id`. Include this ID when contacting support. Never send your API key.

| Status | Meaning | Action |
| --- | --- | --- |
| 400 | Invalid JSON | Correct the request body |
| 401 | Missing, invalid, expired, or revoked key | Check your backend secret |
| 403 | Missing scope | Use a key with the required permission |
| 404 | Unknown route or inaccessible resource | Check the ID, account, and environment |
| 409 | Conflicting idempotency key or resource state | Inspect the existing operation; do not retry blindly |
| 413 | Body exceeds 64 KiB | Reduce request size |
| 415 | Unsupported content type | Send `application/json` |
| 422 | Invalid fields | Correct the named fields |
| 429 | Rate limit exceeded | Wait for `Retry-After` before retrying |
| 500 | Unexpected server error | Retry with the same idempotency key and backoff |
| 503 | API unavailable or carrier not configured | Inspect `error.code`; carrier activation is required for live labels |

## Safe retries

All POST and DELETE requests require `Idempotency-Key`, including rate requests. Use a unique, persistent string for each logical operation, up to 128 characters. For example:

```http
Idempotency-Key: warehouse-order-1001-label
```

Shipide stores the successful response with the operation. If the connection times out after success, retry the same method, path, JSON content, and key. The original status and response are returned with `Idempotency-Replayed: true`. Object key order does not affect matching. Changing the body, method, or path while reusing the key returns `409 idempotency_conflict`.

Keys are scoped to your account and environment, not to a single API credential. Rotating an API key does not remove deduplication. Successful idempotency records currently have no automatic expiry. Failed operations do not reserve the key.

An original response is a snapshot. If its resource later changes, a replay still returns the original response. Retrieve the resource to get its current state.

## Limits

Each API key permits 120 authenticated requests per minute, including reads and retries. Exceeding the limit returns `Retry-After: 60`. Use exponential backoff with jitter for transient failures. Do not retry validation errors unchanged, and do not repeatedly retry `carrier_not_configured`.
