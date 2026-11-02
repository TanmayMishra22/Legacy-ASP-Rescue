# Classic ASP API Diagnostics

A standalone proof-of-work example showing a practical approach to diagnosing HTTP/API failures from Classic ASP.

## Demonstrates
- HTTP status inspection
- Response body capture
- Transport-error handling
- Authentication/authorization diagnosis
- Timeout/error diagnosis
- Safe handling of credentials

## Diagnostic flow

Browser -> Classic ASP -> HTTP request -> API response -> status/body inspection -> diagnosis

## Common failures

| Symptom | Check |
|---|---|
| Timeout / no response | DNS, network, firewall, TLS, endpoint |
| 401 / 403 | Authentication, token, scopes |
| 400 | Request format, parameters, required fields |
| 404 | URL, endpoint, API version |
| 429 | Rate limits and retry behavior |
| 500+ | Remote service/server |
| 200 but bad data | Response validation / JSON parsing |

This example contains no client code, credentials, or proprietary data.
