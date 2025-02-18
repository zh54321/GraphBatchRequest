# GraphBatchRequest - PowerShell Module

## Introduction

The `GraphBatchRequest` PowerShell module allows to send batch requests to the Microsoft Graph API.
It supports automatic throttling handling, pagination, and can return results in either PowerShell object format or raw JSON.

This module is useful for executing multiple Microsoft Graph requests in a single API call, reducing network traffic and improving efficiency.

## Parameters

| Parameter                    | Description                                                                                 |
| ---------------------------- | ------------------------------------------------------------------------------------------- |
| `-AccessToken` *(Mandatory)* | The OAuth access token to authenticate against Microsoft Graph API.                         |
| `-Requests` *(Mandatory)*    | An array of request objects formatted for Microsoft Graph batch requests.                   |
| `-MaxRetries` *(Default: 5)* | Specifies the maximum number of retry attempts for failed requests.                         |
| `-JsonDepth` *(Default: 10)* | Specifies the depth for JSON conversion. Useful for deeply nested objects.                  |
| `-VerboseMode`               | Enables verbose logging to provide additional information about request processing.         |
| `-BetaAPI`                   | If specified, uses the Microsoft Graph `Beta` endpoint instead of `v1.0`.                 |
| `-RawJson`                   | If specified, returns the response as a raw JSON string instead of a PowerShell object. |
| `-DebugMode`                 | Enables detailed debugging output, including request and response details.                  |

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

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -VerboseMode

#Show data
$Response

#Show the users and groups
$Response.response.value

```

### Example 2: **Create a New Group**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Requests = @(
    @{ 
        "id" = "1"
        "method" = "POST"
        "url" = "/groups"
        "body" = @{ 
            "displayName" = "New Group"
            "mailEnabled" = $false
            "mailNickname" = "test213"
            "securityEnabled" = $true
        }
        "headers"= @{
        "Content-Type"= "application/json"
        }
    }
)

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -RawJson
$Response
```

### Example 3: **Use the Beta API Endpoint**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Requests = @(
    @{ "id" = "1"; "method" = "GET"; "url" = "/me" }
)

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -BetaAPI
$Response
```

### Example 4: **Generate Dynamic Requests**

Asuming you have an array of group objects stored in $GroupsInput
```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$GroupsInput | ForEach-Object {
    $RequestID++
    $Requests += @{
        "id"     = $RequestID  # Unique request ID
        "method" = "GET"
        "url"    = "/groups/$($_.id)"  # Graph API URL for each group
    }
}

$Response = Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests
$Response
```

## Notes

- Ensure that you have **valid Microsoft Graph API permissions** before executing requests.
- The module automatically handles **the 429 throttling errors** using **exponential backoff**.
- Requests are **automatically split** into batches of **20 requests per API call**, as required by Microsoft Graph.

