# GraphBatchRequest - PowerShell Module

## Introduction

The `GraphBatchRequest` PowerShell module allows users to send batch requests to the Microsoft Graph API.
It supports automatic throttling handling, pagination, and can return results in either PowerShell object format or raw JSON.

This module is useful for executing multiple Microsoft Graph requests in a single API call, reducing network traffic and improving efficiency.

Note: Cleartext access tokens can be obtained, for example, using [EntraTokenAid](https://github.com/zh54321/EntraTokenAid).

## Parameters

| Parameter                    | Description                                                                                 |
| ---------------------------- | ------------------------------------------------------------------------------------------- |
| `-AccessToken` *(Mandatory)* | The OAuth access token to authenticate against Microsoft Graph API. Mutually exclusive with `-AccessTokenProvider`. |
| `-AccessTokenProvider`       | Script block invoked before every request (incl. pagination/retries). Must return a valid token and handle refresh itself. For calls that may outlive a single token. Mutually exclusive with `-AccessToken`. |
| `-Requests` *(Mandatory)*    | An array of request objects formatted for Microsoft Graph batch requests. Never modified by the module. |
| `-MaxRetries` *(Default: 6)* | Specifies the maximum number of retry attempts for failed requests. Sub-request failures and batch-request failures get independent budgets. |
| `-JsonDepthRequest` *(Default: 10)* | Specifies the depth for JSON conversion (request). Useful for deeply nested objects. |
| `-UserAgent`                 | Specifies the user agent string to use for the HTTP requests.                               |
| `-VerboseMode`               | Enables verbose output to give some information about the amount of sent requests.          |
| `-DebugMode`                 | Enables verbose logging to provide additional information about request processing.         |
| `-BetaAPI`                   | If specified, uses the Microsoft Graph `Beta` endpoint instead of `v1.0`.                   |
| `-MaxBatchSize` *(Default: 20)* | Specifies the maximum number of Graph subrequests per batch request. Valid range: `1..20`. |
| `-Proxy`                     | Specifies a web proxy to use for the HTTP request (e.g., http://proxyserver:8080).          |
| `-SkipCertificateCheck`      | If specified, skips TLS certificate validation (PS 7 only).                                 |
| `-RawJson`                   | If specified, returns the response as a raw JSON string instead of a PowerShell object.     |
| `-BatchDelay` *(Default: 0)* | Specifies a delay in seconds between each batch request to avoid throttling.                |
| `-QueryParameters`           | Query parameters (e.g., @{ '$select' = 'displayName'}) applied to all requests.             |
| `-Silent`                    | Suppresses error output (for example, when a sub-request returns an HTTP 400 error).        |
| `-DisablePagination`         | Prevents the function from automatically following @odata.nextLink. Affected results come back with `complete = $false` and a `nextLink` so you can page yourself. |
| `-JsonDepthResponse` *(Default: 10)* | Specifies the depth for JSON conversion (response). Useful for deeply nested objects in combination with `-RawJson`. |

## Results

You get **exactly one entry per requested ID**, in the order you submitted them, so you can validate
a run by comparing the returned ID set against the requested one.

| Field | Present | Meaning |
| ----- | ------- | ------- |
| `id`      | always | The request ID you supplied. |
| `status`  | always | The Graph status (`200`, `201`, `204`, `403`, …), or `"unknown"` if no response was seen. |
| `response`| on success | The response body. Collection endpoints keep the `value` wrapper; single-object GETs, creates and action results return the body as Graph sent it. |
| `errorCode` / `errorMessage` | on failure | The Graph error. |
| `complete` | always | `$true` only when the full result for that ID was retrieved. |
| `incompleteReason` | when `complete` is `$false` | `PaginationFailed`, `RequestFailed`, `NotAttempted`, or `PaginationDisabled`. |
| `nextLink` | when known and not followed | Continuation link, so you can resume. |

Because `complete` is on every entry, one uniform check finds every truncated or failed result:

```powershell
$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests
$Incomplete = $Response | Where-Object { -not $_.complete }
if ($Incomplete) {
    $Incomplete | ForEach-Object { Write-Warning "$($_.id): $($_.incompleteReason)" }
}
```

A batch request that cannot be delivered do not discard the run: everything already collected
is returned, and the requests that never went out are reported as `NotAttempted`.

## Examples

### Example 1: **Retrieve Groups and Users**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Requests = @(
    @{ 
        "id" = "1"
        "method" = "GET"
        "url" = "/groups" 
    },
    @{ 
        "id" = "2"
        "method" = "GET"
        "url" = "/users" 
    }
)

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests

#Show data
$Response

#Show the users and groups (all results)
$Response.response.value

#Show only the values of request 2 (users)
($Response | Where-Object { $_.id -eq 2 }).response.value
```

### Example 2: **Create a New Microsoft 365 Group**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Requests = @(
    @{ 
        "id" = "1"
        "method" = "POST"
        "url" = "/groups"
        "body" = @{ "displayName" = "New Group"; "mailEnabled" = $false; "mailNickname" = "whatever"; "securityEnabled" = $true }
        "headers" = @{"Content-Type"= "application/json"} 
    }
)

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -RawJson
$Response
```

### Example 3: **Global Query Parameters and Proxy Usage**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Requests = @(
    @{ "id" = "1"; "method" = "GET"; "url" = "/groups"},
    @{"id" = "2"; "method" = "GET"; "url" = "/users"}
)
$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -proxy http://127.0.0.1:8080 -QueryParameters @{'$select' = 'id,displayName' }
$Response.response
```

### Example 4: **Request-Level Query Parameters**

```powershell
    $AccessToken = "YOUR_ACCESS_TOKEN"
    $Requests = @(
        @{ id = "1"; method = "GET"; url = "/users"; queryParameters = @{ '$filter' = "startswith(displayName,'Adele')"; '$select' = 'displayName' } },
        @{ id = "2"; method = "GET"; url = "/groups"; queryParameters = @{ '$select' = 'id' } }
    )
$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests
$Response.response
```

### Example 5: **Generate Dynamic Requests**

Assuming you have an array of group objects stored in $groups
```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"

$RequestID = 0
$groups | ForEach-Object {
    $RequestID ++
    $Requests += @{
        "id"     = $RequestID  # Unique request ID
        "method" = "GET"
        "url"    = "/groups/$($_.id)"  # Graph API URL for each group
    }
}

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -VerboseMode
$Response.response
```

### Example 6: **Long-Running Collection With Token Renewal**

For runs that may outlive a single access token, pass a scriptblock instead of a string. It is
consulted before every HTTP request, including pagination. The module never parses or renews the
token itself; it just asks you for the current one.

```powershell
$Provider = {
    if (($script:Token.Expiration_time - [datetime]::Now).TotalMinutes -lt 30) {
        Update-MyToken   # your own renewal
    }
    $script:Token.access_token
}

$Response = Send-GraphBatchRequest -AccessTokenProvider $Provider -Requests $Requests
```

## Notes

- Ensure that you have **valid Microsoft Graph API permissions** before executing requests.
- The module automatically handles **the 429 throttling errors** using **exponential backoff**.
- Requests are **automatically split** into batches of up to **20 requests per API call** by default, as required by Microsoft Graph. Use `-MaxBatchSize` with a smaller value to reduce burst concurrency.
- The batch request itself is retried on transport failures (connection reset, DNS, timeout) and on retryable HTTP statuses. Statuses that cannot succeed on retry, such as `400` or `403`, fail immediately rather than sleeping through the backoff.
- With `-AccessTokenProvider`, HTTP `401` also becomes retryable, since a retry can pick up a freshly issued token.

### Behaviour changes to be aware of

Upgrading from an earlier version:

- Requests whose response is **not** a collection (single-object `GET`, `POST` create, action results) previously came back as `status = 200` with `response.value = @()`, silently discarding the body. They now return the real status and the real body.
- Successful results previously always reported `status = 200`. They now report the status Graph returned, so a create reports `201`. Range checks (`$_.status -ge 200 -and $_.status -lt 300`) are unaffected; equality checks against `200` are not.
- Results are returned in submission order rather than in hashtable enumeration order.
