<#
.SYNOPSIS
    Sends a batch request to Microsoft Graph API.

.DESCRIPTION
    The Send-GraphBatchRequest function sends multiple Microsoft Graph API requests as a batch.
    It supports automatic throttling handling, pagination, and can return results in either
    PowerShell object format or raw JSON.

.PARAMETER AccessToken
    The OAuth access token to authenticate against Microsoft Graph API.

.PARAMETER Requests
    An array of request objects formatted for Microsoft Graph batch requests.

.PARAMETER VerboseMode
    Enables verbose logging to provide additional information about request processing.

.PARAMETER UserAgent
    Specifies the user agent string to be used in the HTTP requests. This can be customized to mimic specific browser or application behavior.
    Default: `python-requests/2.32.3`

.PARAMETER MaxRetries
    Specifies the maximum number of retry attempts for failed requests. Default is 5.

.PARAMETER Beta
    If specified, uses the Graph Beta endpoint instead of v1.0.

.PARAMETER RawJson
    If specified, returns the response as a raw JSON string instead of a PowerShell object.

.PARAMETER BatchDelay
    Specifies a delay in seconds between each batch request to avoid throttling. Default is 0 (no delay).

.PARAMETER Proxy
    Specifies a web proxy to use for the HTTP request (e.g., http://proxyserver:8080). Useful for debugging, traffic inspection.

.PARAMETER JsonDepthRequest
    Specifies the depth for JSON conversion in the request. Default is 10, but can be increased for complex objects.

.PARAMETER QueryParameters
    A hashtable of query parameters (e.g., @{ '$select' = 'displayName'; '$top' = '5' }) applied to all requests.
    Individual requests can override or add their own query parameters by including a `queryParameters` hashtable in the request object.

.PARAMETER JsonDepthResponse
    Specifies the depth for JSON conversion in the response (to use with -RawJson). Default is 10, but can be increased for complex objects.

.EXAMPLE
    $AccessToken = "YOUR_ACCESS_TOKEN"
    $Requests = @(
        @{ "id" = "1"; "method" = "GET"; "url" = "/groups" }
    )
    
    Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -VerboseMode

.EXAMPLE
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
    
    Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -RawJson

.EXAMPLE
    $AccessToken = "YOUR_ACCESS_TOKEN"
    $Requests = @(
        @{ "id" = "1"; "method" = "GET"; "url" = "/groups"},
        @{"id" = "2"; "method" = "GET"; "url" = "/users"}
    )
    Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -VerboseMode -proxy http://127.0.0.1:8080 -QueryParameters @{'$select' = 'displayName' }

.EXAMPLE
    $AccessToken = "YOUR_ACCESS_TOKEN"
    $Requests = @(
        @{ id = "1"; method = "GET"; url = "/users"; queryParameters = @{ '$filter' = "startswith(displayName,'Adele')"; '$select' = 'displayName' } },
        @{ id = "2"; method = "GET"; url = "/groups"; queryParameters = @{ '$select' = 'id' } }
    )
    Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests

.NOTES
    Author: ZH54321
    GitHub: https://github.com/zh54321/GraphBatchRequest
#>

function Send-GraphBatchRequest { 
    param (
        [Parameter(Mandatory = $true)]
        [string]$AccessToken,

        [Parameter(Mandatory = $true)]
        [array]$Requests,

        [int]$MaxRetries = 5,
        [int]$JsonDepthRequest = 10,
        [int]$JsonDepthResponse = 10,
        [string]$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Microsoft Windows 10.0.19045; en-us) PowerShell/7.5.0",
        [int]$BatchDelay = 0,
        [string]$Proxy,
        [hashtable]$QueryParameters,
        [switch]$VerboseMode,
        [switch]$BetaAPI,
        [switch]$RawJson
    )

    # Constants
    $ApiVersion = if ($BetaAPI) { "beta" } else { "v1.0" }
    $BatchUrl = "https://graph.microsoft.com/$ApiVersion/`$batch"
    $MaxBatchSize = 20
    
    # Validate Requests
    if (-not $Requests -or $Requests.Count -eq 0) {
        Write-Error "No requests provided."
        return
    }
    
    # Split Requests into Batches (max 20 per batch)
    $Batches = @()
    for ($i = 0; $i -lt $Requests.Count; $i += $MaxBatchSize) {
        $Batches += , ($Requests[$i..([math]::Min($i + $MaxBatchSize - 1, $Requests.Count - 1))])
    }
    
    # Execute Batches
    $Results = @()
    
    foreach ($Batch in $Batches) {
        $PendingRequests = $Batch
        $RetryCount = 0

        foreach ($req in $PendingRequests) {
            $effectiveParams = @{}
        
            # Use per-request queryParameters if defined
            if ($req.ContainsKey('queryParameters')) {
                $effectiveParams += $req.queryParameters
            }
        
            # Merge in global parameters if not already present
            if ($QueryParameters) {
                foreach ($key in $QueryParameters.Keys) {
                    if (-not $effectiveParams.ContainsKey($key)) {
                        $effectiveParams[$key] = $QueryParameters[$key]
                    }
                }
            }
        
            # Build and append query string to the URL
            if ($effectiveParams.Count -gt 0) {
                $queryString = ($effectiveParams.GetEnumerator() |
                    ForEach-Object {
                        "$($_.Key)=$([uri]::EscapeDataString($_.Value))"
                    }) -join '&'
        
                if ($req.url -notmatch "\?") {
                    $req.url = "$($req.url)?$queryString"
                } else {
                    $req.url = "$($req.url)&$queryString"
                }
            }
        }

        do {
            $BatchRequest = @{ requests = $PendingRequests }
            $Headers = @{
                "User-Agent" = $UserAgent
                "Authorization" = "Bearer $AccessToken"
                "Content-Type" = "application/json"
            }
            
            if ($VerboseMode) { Write-Host "Sending batch request: $($BatchRequest | ConvertTo-Json -Depth $JsonDepthRequest)" }

            try {

                # Request Parameters
                $irmParams = @{
                    Uri         = $BatchUrl
                    Method      = 'POST'
                    Headers     = $Headers
                    Body        = ($BatchRequest | ConvertTo-Json -Depth $JsonDepthRequest)
                    ErrorAction = 'Stop'
                }
            
                if ($Proxy) {
                    $irmParams['Proxy'] = $Proxy
                }
                
                # Send request
                $Response = Invoke-RestMethod @irmParams
                $PendingRequests = @()  # Reset failed requests
            } catch {
                Write-Error "Batch request failed: $_"
                return
            }

            # Process responses
            $FailedRequests = @()
            foreach ($Resp in $Response.responses) {
                if ($Resp.status -ge 200 -and $Resp.status -lt 300) {
                    # Handle pagination if needed
                    $ResultData = $Resp.body
                    while ($ResultData -and $ResultData.'@odata.nextLink') {
                        try {
                            if ($VerboseMode) { Write-Host "Fetching next page: $($ResultData.'@odata.nextLink')" }

                            # Request Parameters
                            $irmParams = @{
                                Uri         = $ResultData.'@odata.nextLink'
                                Headers     = $Headers
                                Method      = 'GET'
                                ErrorAction = 'Stop'
                            }
                            
                            if ($Proxy) {
                                $irmParams['Proxy'] = $Proxy
                            }
                            
                            # Send request
                            $NextResponse = Invoke-RestMethod @irmParams

                            $ResultData.value += $NextResponse.value
                            $ResultData.'@odata.nextLink' = $NextResponse.'@odata.nextLink'
                        } catch {
                            Write-Error "Failed to fetch next page: $_"
                            break
                        }
                    }
                    $Results += @{ id = $Resp.id; status = $Resp.status; response = $ResultData }
                } else {
                    $ErrorCode = $Resp.body.error.code
                    $ErrorMessage = $Resp.body.error.message
                    write-host "[!] Graph Batch Request: ID $($Resp.id) failed with status $($Resp.status): $ErrorCode - $ErrorMessage"
                    # Handle throttling & transient errors per request
                    if ($Resp.status -in @(429, 500, 502, 503, 504)) {
                        $RetryAfter = $null

                        # Try to parse "Retry-After" from error message
                        if ($ErrorMessage -match "try after (\d+) seconds") {
                            $RetryAfter = [int]$matches[1]
                        }
                        
                        if ($RetryAfter) {
                            if ($VerboseMode) {
                                write-host "[i] Retrying request $($Resp.id) after $RetryAfter seconds..."
                            } else {
                                write-host "[!] Request will be resent in $RetryAfter seconds..."
                            }
                            Start-Sleep -Seconds $RetryAfter
                        } else {
                            if ($RetryCount -eq 0) {
                                $Backoff = 0
                                write-host "[i] Retrying request $($Resp.id)..."
                            } else {
                                $Backoff = [math]::Pow(2, $RetryCount)
                                write-host "[!] Request will be resent in $Backoff seconds..."
                            }
                            Start-Sleep -Seconds $Backoff
                        }
                        # Add to failed requests for retry
                        $FailedRequests += $Batch | Where-Object { $_.id -eq $Resp.id }
                    } else {
                        # If it's a non-retryable error, log it and move on
                        $Results += @{ id = $Resp.id; status = $Resp.status; errorCode = $ErrorCode; errorMessage = $ErrorMessage }
                    }
                }
            }

            # Update pending requests for retry
            $PendingRequests = $FailedRequests
            $RetryCount++
        } while ($PendingRequests.Count -gt 0 -and $RetryCount -lt $MaxRetries)

        if ($BatchDelay -gt 0) {
            if ($VerboseMode) { Write-Host "Sleeping $BatchDelay seconds before next batch..." }
            Start-Sleep -Seconds $BatchDelay
        }
    }

    # Return JSON if -RawJson switch is used, otherwise return PowerShell object
    if ($RawJson) {
        return $Results | ConvertTo-Json -Depth $JsonDepthResponse
    } else {
        return $Results
    }
}
