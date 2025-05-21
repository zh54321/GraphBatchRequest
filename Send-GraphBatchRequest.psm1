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

.PARAMETER DebugMode
    Enables verbose debug logging to provide additional information about request processing.

.PARAMETER VerboseMode
    Enables verbose output to give some information about the amount of sent requests.

.PARAMETER UserAgent
    Specifies the user agent string to be used in the HTTP requests. This can be customized to mimic specific browser or application behavior.
    Default: `python-requests/2.32.3`

.PARAMETER MaxRetries
    Specifies the maximum number of retry attempts for failed requests. Default is 5.

.PARAMETER BetaAPI
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
    
    Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -DebugMode

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
    Send-GraphBatchRequest -AccessToken $AccessToken -Requests $Requests -DebugMode -proxy http://127.0.0.1:8080 -QueryParameters @{'$select' = 'displayName' }

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
        [switch]$DebugMode,
        [switch]$VerboseMode,
        [switch]$BetaAPI,
        [switch]$RawJson
    )

    $ApiVersion = if ($BetaAPI) { "beta" } else { "v1.0" }
    $BatchUrl = "https://graph.microsoft.com/$ApiVersion/`$batch"
    $MaxBatchSize = 20
    $HttpRequestCount = 0
    $SubRequestCount = 0

    if (-not $Requests -or $Requests.Count -eq 0) {
        Write-Error "No requests provided."
        return
    }

    $Batches = New-Object 'System.Collections.Generic.List[object]'
    for ($i = 0; $i -lt $Requests.Count; $i += $MaxBatchSize) {
        $Batches.Add($Requests[$i..([math]::Min($i + $MaxBatchSize - 1, $Requests.Count - 1))])
    }

    $Results = New-Object 'System.Collections.Generic.List[object]'
    $GlobalNextLinks = New-Object 'System.Collections.Generic.List[string]'
    $PagedResultsMap = @{}

    foreach ($Batch in $Batches) {
        $PendingRequests = $Batch
        $RetryCount = 0

        foreach ($req in $PendingRequests) {
            $effectiveParams = @{}
            if ($req.ContainsKey('queryParameters')) {
                $effectiveParams += $req.queryParameters
            }
            if ($QueryParameters) {
                foreach ($key in $QueryParameters.Keys) {
                    if (-not $effectiveParams.ContainsKey($key)) {
                        $effectiveParams[$key] = $QueryParameters[$key]
                    }
                }
            }
            if ($effectiveParams.Count -gt 0) {
                $queryString = ($effectiveParams.GetEnumerator() | ForEach-Object {
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

            $irmParams = @{
                Uri         = $BatchUrl
                Method      = 'POST'
                Headers     = $Headers
                Body        = ($BatchRequest | ConvertTo-Json -Depth $JsonDepthRequest)
                ErrorAction = 'Stop'
            }

            if ($Proxy) { $irmParams['Proxy'] = $Proxy }
            $HttpRequestCount++
            $SubRequestCount += $PendingRequests.Count

            try {
                $Response = Invoke-RestMethod @irmParams
                $PendingRequests = @()
            } catch {
                Write-Error "Batch request failed: $_"
                return
            }

            $FailedRequests = @()
            foreach ($Resp in $Response.responses) {
                if ($Resp.status -ge 200 -and $Resp.status -lt 300) {
                    $ResultData = $Resp.body
                    $PagedResultsMap[$Resp.id] = New-Object 'System.Collections.Generic.List[object]'
                    if ($ResultData.value) {
                        $ResultData.value | ForEach-Object { $PagedResultsMap[$Resp.id].Add($_) }
                    }
                    if ($ResultData.'@odata.nextLink') {
                        $GlobalNextLinks.Add("$($Resp.id)|$($ResultData.'@odata.nextLink')")
                    }
                } else {
                    $ErrorCode = $Resp.body.error.code
                    $ErrorMessage = $Resp.body.error.message
                    Write-Host "[!] Graph Batch Request: ID $($Resp.id) failed with status $($Resp.status): $ErrorCode - $ErrorMessage"
                    if ($Resp.status -in @(429, 500, 502, 503, 504)) {
                        $FailedRequests += $Batch | Where-Object { $_.id -eq $Resp.id }
                        Start-Sleep -Seconds ([math]::Pow(2, $RetryCount))
                    } else {
                        $Results.Add(@{ id = $Resp.id; status = $Resp.status; errorCode = $ErrorCode; errorMessage = $ErrorMessage })
                    }
                }
            }

            $PendingRequests = $FailedRequests
            $RetryCount++
        } while ($PendingRequests.Count -gt 0 -and $RetryCount -lt $MaxRetries)

        if ($BatchDelay -gt 0) {
            Start-Sleep -Seconds $BatchDelay
        }
    }

	while ($GlobalNextLinks.Count -gt 0) {
		$ToFetch = $GlobalNextLinks[0..([math]::Min(19, $GlobalNextLinks.Count - 1))]
		$GlobalNextLinks.RemoveRange(0, $ToFetch.Count)

		$Links = $ToFetch | ForEach-Object { ($_ -split '\|')[1] }
		$Ids   = $ToFetch | ForEach-Object { ($_ -split '\|')[0] }

		$BatchResult = Invoke-GraphNextLinkBatch -NextLinks $Links `
            -Ids $Ids `
            -AccessToken $AccessToken `
            -UserAgent $UserAgent `
            -JsonDepthRequest $JsonDepthRequest `
            -JsonDepthResponse $JsonDepthResponse `
            -Proxy $Proxy `
            -VerboseMode:$VerboseMode `
            -DebugMode:$DebugMode `
            -HttpRequestCount ([ref]$HttpRequestCount) `
            -SubRequestCount ([ref]$SubRequestCount)`
            -ApiVersion $ApiVersion

        foreach ($id in $BatchResult.values.Keys) {
            if (-not $PagedResultsMap.ContainsKey($id)) {
                Write-Warning ("[{0}] [!] Missing first-page data for ID {1} - initializing empty list." -f (Get-Date -Format "HH:mm:ss"), $id)
                $PagedResultsMap[$id] = New-Object 'System.Collections.Generic.List[object]'
            }
        
            $PagedResultsMap[$id].AddRange($BatchResult.values[$id])
        
            if ($BatchResult.nextLinks.ContainsKey($id)) {
                $GlobalNextLinks.Add("$id|$($BatchResult.nextLinks[$id])")
            }
        }
	}

    foreach ($id in $PagedResultsMap.Keys) {
        $Results.Add(@{ id = $id; status = 200; response = @{ value = $PagedResultsMap[$id].ToArray() } })
    }

    if ($VerboseMode) {
        Write-Host "[i] Total HTTP requests sent (including pagination): $HttpRequestCount"
        Write-Host "[i] Total Graph subrequests sent (individual operations): $SubRequestCount"
    }

    if ($RawJson) {
        return $Results | ConvertTo-Json -Depth $JsonDepthResponse
    } else {
        return $Results
    }
}


function Invoke-GraphNextLinkBatch {
    param (
        [string[]]$NextLinks,
        [string[]]$Ids,
        [string]$AccessToken,
        [string]$UserAgent,
        [int]$JsonDepthResponse = 10,
        [int]$JsonDepthRequest = 10,
        [string]$Proxy,
		[ref]$HttpRequestCount,
		[ref]$SubRequestCount,
        [switch]$VerboseMode,
		[switch]$DebugMode,
        [string]$ApiVersion
    )

    $ResultsList = New-Object 'System.Collections.Generic.List[object]'
    $MoreNextLinks = New-Object 'System.Collections.Generic.List[string]'

    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "User-Agent"    = $UserAgent
        "Content-Type"  = "application/json"
    }

    for ($i = 0; $i -lt $NextLinks.Count; $i += 20) {
        $BatchSet = $NextLinks[$i..([math]::Min($i + 19, $NextLinks.Count - 1))]
        $BatchRequests = @()
        $index = 0

        foreach ($link in $BatchSet) {
            $relativeUrl = $link -replace '^https://graph\.microsoft\.com/[^/]+', ''
            $BatchRequests += @{
                id     = "nl_$index"
                method = "GET"
                url    = $relativeUrl
            }
            $index++
        }

        $BatchBody = @{ requests = $BatchRequests }

        $irmParams = @{
            Uri         = "https://graph.microsoft.com/$ApiVersion/`$batch"
            Method      = 'POST'
            Headers     = $Headers
            Body        = ($BatchBody | ConvertTo-Json -Depth $JsonDepthRequest)
            ErrorAction = 'Stop'
        }

        if ($Proxy) { $irmParams['Proxy'] = $Proxy }

        try {
            if ($DebugMode) { Write-Host "[i] Sending nextLink batch request..." }
			$HttpRequestCount.Value++
			$SubRequestCount.Value += $BatchRequests.Count
            $BatchResp = Invoke-RestMethod @irmParams

            $AllValues = New-Object 'System.Collections.Generic.List[object]'
            $AllNextLinks = New-Object 'System.Collections.Generic.List[string]'

            foreach ($resp in $BatchResp.responses) {
                if ($resp.status -ge 200 -and $resp.status -lt 300) {
                    $data = $resp.body

                    # Ensure each slot is an array (even if null)
                    if ($data.value) {
                        $AllValues.Add(@($data.value))
                    } else {
                        $AllValues.Add(@())
                    }

                    if ($data.'@odata.nextLink') {
                        $AllNextLinks.Add($data.'@odata.nextLink')
                    } else {
                        $AllNextLinks.Add($null)
                    }
                } else {
                    Write-Warning "NextLink subrequest failed: ID $($resp.id) ($($resp.status))"
                    $AllValues.Add(@())        # Maintain index consistency
                    $AllNextLinks.Add($null)
                }
            }
        } catch {
            Write-Error "Failed nextLink batch: $_ "
        }
    }
    $ResultMap = @{}
    $MoreLinksMap = @{}
    
    foreach ($resp in $BatchResp.responses) {
        $i = [int]($resp.id -replace 'nl_', '')
        $realId = $Ids[$i]
    
        if (-not $ResultMap.ContainsKey($realId)) {
            $ResultMap[$realId] = New-Object 'System.Collections.Generic.List[object]'
        }
    
        if ($resp.body.value) {
            $ResultMap[$realId].AddRange(@($resp.body.value))
        }
    
        if ($resp.body.'@odata.nextLink') {
            $MoreLinksMap[$realId] = $resp.body.'@odata.nextLink'
        }
    }
    return @{
        values     = $ResultMap
        nextLinks  = $MoreLinksMap
    }
}
