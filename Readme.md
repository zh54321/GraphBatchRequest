# GraphBatchRequest - PowerShell Module

## Introduction

The `GraphBatchRequest` PowerShell module allows users to send batch requests to the Microsoft Graph API.
It supports automatic throttling handling, pagination, and can return results in either PowerShell object format or raw JSON.

This module is useful for executing multiple Microsoft Graph requests in a single API call, reducing network traffic and improving efficiency.

Note: Cleartext access tokens can be obtained, for example, using [EntraTokenAid](https://github.com/zh54321/EntraTokenAid).

## Parameters

| Parameter                    | Description                                                                                 |
| ---------------------------- | ------------------------------------------------------------------------------------------- |
| `-AccessToken` *(Mandatory)* | The OAuth access token to authenticate against Microsoft Graph API.                         |
| `-Requests` *(Mandatory)*    | An array of request objects formatted for Microsoft Graph batch requests.                   |
| `-MaxRetries` *(Default: 5)* | Specifies the maximum number of retry attempts for failed requests.                         |
| `-JsonDepthRequest` *(Default: 10)* | Specifies the depth for JSON conversion (request). Useful for deeply nested objects. |
| `-VerboseMode`               | Enables verbose logging to provide additional information about request processing.         |
| `-BetaAPI`                   | If specified, uses the Microsoft Graph `Beta` endpoint instead of `v1.0`.                   |
| `-RawJson`                   | If specified, returns the response as a raw JSON string instead of a PowerShell object.     |
| `-BatchDelay` *(Default: 0)* | Specifies a delay in seconds between each batch request to avoid throttling.                |
| `-JsonDepthResponse` *(Default: 10)* | Specifies the depth for JSON conversion (request). Useful for deeply nested objects in combination with `-RawJson` .  |

## Examples

### Example 1: **Retrieve All Groups**

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

### Example 3: **Use the Beta API Endpoint and verbode mode**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Requests = @(
    @{ "id" = "1"; "method" = "GET"; "url" = "/me" }
)

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -BetaAPI -VerboseMode
$Response.response
```

### Example 4: **Generate Dynamic Requests**

Asuming you have an array of group objects stored in $groups
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

## Notes

- Ensure that you have **valid Microsoft Graph API permissions** before executing requests.
- The module automatically handles **the 429 throttling errors** using **exponential backoff**.
- Requests are **automatically split** into batches of **20 requests per API call**, as required by Microsoft Graph.