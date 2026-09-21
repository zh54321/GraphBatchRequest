<#
.SYNOPSIS
    Sends a batch request to Microsoft Graph API.

.DESCRIPTION
    The Send-GraphBatchRequest function sends multiple Microsoft Graph API requests as a batch.
    It supports automatic throttling handling, pagination, and can return results in either
    PowerShell object format or raw JSON.

    Every requested ID is represented by exactly one result entry, and every entry carries a
    `complete` flag so a caller can tell a whole result from a truncated one without inspecting
    the warning stream.

.PARAMETER AccessToken
    The OAuth access token to authenticate against Microsoft Graph API.
    Mutually exclusive with -AccessTokenProvider.

.PARAMETER AccessTokenProvider
    A scriptblock returning the access token to use, consulted immediately before every HTTP
    request (initial batches and pagination batches alike). Use this for calls that may outlive a
    single token; the caller owns all renewal logic, the module only asks for a string.
    Mutually exclusive with -AccessToken.

    When a provider is supplied, HTTP 401 is treated as retryable, since a retry can pick up a
    freshly issued token.

.PARAMETER Requests
    An array of request objects formatted for Microsoft Graph batch requests.
    Request objects are never modified; query parameters are applied to an internal copy.

.PARAMETER DebugMode
    Enables verbose debug logging to provide additional information about request processing.

.PARAMETER VerboseMode
    Enables verbose output to give some information about the amount of sent requests.

.PARAMETER UserAgent
    Specifies the user agent string to be used in the HTTP requests. This can be customized to mimic specific browser or application behavior.
    Default: `Mozilla/5.0 (Windows NT 10.0; Microsoft Windows 10.0.19045; en-us) PowerShell/7.5.0`

.PARAMETER MaxRetries
    Specifies the maximum number of retry attempts for failed requests. Default is 6.
    Applies independently to sub-request failures and to transport/HTTP failures of the batch
    request itself, so a flaky connection cannot consume the budget a later 429 needs.

.PARAMETER BetaAPI
    If specified, uses the Graph Beta endpoint instead of v1.0.

.PARAMETER RawJson
    If specified, returns the response as a raw JSON string instead of a PowerShell object.

.PARAMETER BatchDelay
    Specifies a delay in seconds between each batch request to avoid throttling. Default is 0 (no delay).

.PARAMETER MaxBatchSize
    Specifies the maximum number of Graph subrequests to include in each batch request. Default is 20.

.PARAMETER Proxy
    Specifies a web proxy to use for the HTTP request (e.g., http://proxyserver:8080). Useful for debugging, traffic inspection.

.PARAMETER SkipCertificateCheck
    If specified, skips TLS certificate validation for Invoke-RestMethod calls (PS 7).

.PARAMETER JsonDepthRequest
    Specifies the depth for JSON conversion in the request. Default is 10, but can be increased for complex objects.

.PARAMETER QueryParameters
    A hashtable of query parameters (e.g., @{ '$select' = 'displayName'; '$top' = '5' }) applied to all requests.
    Individual requests can override or add their own query parameters by including a `queryParameters` hashtable in the request object.

.PARAMETER JsonDepthResponse
    Specifies the depth for JSON conversion in the response (to use with -RawJson). Default is 10, but can be increased for complex objects.

.PARAMETER Silent
    Suppresses error output (for example, when a sub-request returns an HTTP 400 error).

.PARAMETER DisablePagination
    If specified, prevents the function from automatically following @odata.nextLink for paginated results.
    Results that have a continuation link are returned with `complete = $false` and a `nextLink`
    field so the caller can drive paging itself.

.OUTPUTS
    One entry per requested ID. Successful entries carry the Graph status and the response body;
    failed entries carry `errorCode` / `errorMessage`. Every entry carries:

      complete         - $true only when the full result for that ID was retrieved
      incompleteReason - present when complete is $false, one of:
                           PaginationFailed   one or more continuation pages could not be retrieved
                           RequestFailed      the sub-request itself failed
                           NotAttempted       the request was never sent, or no response was seen
                           PaginationDisabled more data exists but -DisablePagination was used
      nextLink         - present when a continuation link is known but was not followed
      paginationFailureStatus    - HTTP status of a failed continuation request, when available
      paginationFailureErrorCode - error code of a failed continuation request

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

.EXAMPLE
    # Long-running collection that may outlive a single token.
    $Provider = {
        if (($script:Token.Expiration_time - [datetime]::Now).TotalMinutes -lt 30) { Update-MyToken }
        $script:Token.access_token
    }
    $Response = Send-GraphBatchRequest -AccessTokenProvider $Provider -Requests $Requests

    # Detect truncation without parsing warnings.
    $Incomplete = $Response | Where-Object { -not $_.complete }

.NOTES
    Author: ZH54321
    GitHub: https://github.com/zh54321/GraphBatchRequest
#>

function Send-GraphBatchRequest {
    [CmdletBinding(DefaultParameterSetName = 'Token')]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Token')]
        [string]$AccessToken,

        [Parameter(Mandatory = $true, ParameterSetName = 'Provider')]
        [scriptblock]$AccessTokenProvider,

        [Parameter(Mandatory = $true)]
        [array]$Requests,

        [int]$MaxRetries = 6,
        [int]$JsonDepthRequest = 10,
        [int]$JsonDepthResponse = 10,
        [string]$UserAgent = "Mozilla/5.0 (Windows NT 10.0; Microsoft Windows 10.0.19045; en-us) PowerShell/7.5.0",
        [double]$BatchDelay = 0,
        [ValidateRange(1, 20)]
        [int]$MaxBatchSize = 20,
        [string]$Proxy,
        [switch]$SkipCertificateCheck,
        [hashtable]$QueryParameters,
        [switch]$DebugMode,
        [switch]$VerboseMode,
        [switch]$Silent,
        [switch]$BetaAPI,
        [switch]$DisablePagination,
        [switch]$RawJson
    )

    $ApiVersion = if ($BetaAPI) { "beta" } else { "v1.0" }
    $BatchUrl = "https://graph.microsoft.com/$ApiVersion/`$batch"
    $HttpRequestCount = 0
    $SubRequestCount = 0
    $SupportsSkipCertificateCheck = (Get-Command Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck')

    if (-not $Requests -or $Requests.Count -eq 0) {
        Write-Error "No requests provided."
        return
    }

    if ($SkipCertificateCheck -and -not $SupportsSkipCertificateCheck -and -not $Silent) {
        Write-Warning "Current PowerShell does not support -SkipCertificateCheck on Invoke-RestMethod. The flag will be ignored."
    }

    # Retryable for both sub-request statuses and for the batch request itself. With a token
    # provider a 401 retry can pick up a fresh token, so it becomes worth retrying.
    $RetryableStatusCodes = @(429, 500, 502, 503, 504)
    if ($AccessTokenProvider) { $RetryableStatusCodes += 401 }

    # Requested IDs in submission order. Used to guarantee one entry per ID and to give the
    # result set a deterministic order.
    $RequestedIds = New-Object 'System.Collections.Generic.List[object]'
    foreach ($req in $Requests) { $RequestedIds.Add((Get-GraphMember -Object $req -Name 'id')) }

    $Batches = New-Object 'System.Collections.Generic.List[object]'
    for ($i = 0; $i -lt $Requests.Count; $i += $MaxBatchSize) {
        $Batches.Add($Requests[$i..([math]::Min($i + $MaxBatchSize - 1, $Requests.Count - 1))])
    }

    $ErrorEntries = @{}
    $GlobalNextLinks = New-Object 'System.Collections.Generic.List[string]'
    $PagedResultsMap = @{}
    $DirectResults = @{}
    $IncompleteIds = @{}
    $PaginationFailures = @{}
    $UnresolvedNextLinks = @{}
    $Aborted = $false

    foreach ($Batch in $Batches) {
        # Work on copies so the caller's request objects are never modified. Repeated invocation
        # with -QueryParameters must not accumulate query strings.
        $PendingRequests = @()
        foreach ($req in $Batch) {
            $PendingRequests += (Get-GraphEffectiveRequest -Request $req -QueryParameters $QueryParameters)
        }
        # Keyed by string: request IDs may be supplied as integers while Graph echoes them back
        # as JSON strings, and a hashtable lookup does not coerce the way -eq does.
        $BatchRequestsById = @{}
        foreach ($req in $PendingRequests) { $BatchRequestsById[[string]$req.id] = $req }

        $RetryCount = 0
        $LastRetryableErrors = @{}

        do {
            $BatchRequest = @{ requests = $PendingRequests }
            $Body = ($BatchRequest | ConvertTo-Json -Depth $JsonDepthRequest)

            # The batch request itself gets its own retry budget, so transport failures cannot
            # consume the attempts a later throttled sub-request needs.
            $Response = $null
            $TransportAttempt = 0
            $LastTransportError = $null
            while ($true) {
                $CurrentToken = Resolve-GraphAccessToken -AccessToken $AccessToken -AccessTokenProvider $AccessTokenProvider
                $irmParams = @{
                    Uri         = $BatchUrl
                    Method      = 'POST'
                    Headers     = @{
                        "User-Agent"    = $UserAgent
                        "Authorization" = "Bearer $CurrentToken"
                        "Content-Type"  = "application/json"
                    }
                    Body        = $Body
                    ErrorAction = 'Stop'
                }
                if ($Proxy) { $irmParams['Proxy'] = $Proxy }
                if ($SkipCertificateCheck -and $SupportsSkipCertificateCheck) { $irmParams['SkipCertificateCheck'] = $true }

                $HttpRequestCount++
                $SubRequestCount += $PendingRequests.Count

                try {
                    $Response = Invoke-RestMethod @irmParams
                    break
                } catch {
                    $LastTransportError = $_
                    $HttpStatus = Get-GraphErrorStatusCode -ErrorRecord $_
                    # No status means a transport-level failure (reset, DNS, timeout): retryable.
                    $IsRetryable = ($null -eq $HttpStatus) -or ($HttpStatus -in $RetryableStatusCodes)

                    if (-not $IsRetryable -or ($TransportAttempt + 1) -ge $MaxRetries) { break }

                    $TransportDelay = [math]::Pow(2, $TransportAttempt)
                    if (-not $Silent) {
                        Write-Host ("[i] Batch request failed ({0}). Retrying in {1}s (attempt {2}/{3})." -f `
                            $(if ($null -ne $HttpStatus) { $HttpStatus } else { 'transport error' }), $TransportDelay, ($TransportAttempt + 1), $MaxRetries)
                    }
                    Start-Sleep -Seconds $TransportDelay
                    $TransportAttempt++
                }
            }

            if ($null -eq $Response) {
                # The batch request could not be delivered. Stop, but keep everything collected so
                # far; the reconciliation pass below marks whatever was never answered.
                Write-Error "Batch request failed: $LastTransportError"
                $Aborted = $true
                break
            }

            $PendingRequests = @()
            $FailedRequests = @()
            $RetryDelaySeconds = [math]::Pow(2, $RetryCount)

            foreach ($Resp in $Response.responses) {
                if ($Resp.status -ge 200 -and $Resp.status -lt 300) {
                    $ResultData = $Resp.body
                    $HasValue = Test-GraphMember -Object $ResultData -Name 'value'
                    $NextLink = Get-GraphMember -Object $ResultData -Name '@odata.nextLink'

                    if (-not $HasValue -and -not $NextLink) {
                        # Not a collection (single object GET, create, action result). Return the
                        # body and the real status untouched instead of forcing it through the
                        # pagination map, which only ever carries `value` items.
                        $DirectResults[$Resp.id] = @{
                            id       = $Resp.id
                            status   = $Resp.status
                            response = $ResultData
                        }
                        continue
                    }

                    $PagedResultsMap[$Resp.id] = New-Object 'System.Collections.Generic.List[object]'
                    if ($ResultData.value) {
                        $PagedResultsMap[$Resp.id].AddRange(@($ResultData.value))
                    }

                    if ($NextLink) {
                        if ($DisablePagination) {
                            # More data exists. Say so, and hand back the link so the caller can
                            # continue on its own terms.
                            $IncompleteIds[$Resp.id] = 'PaginationDisabled'
                            $UnresolvedNextLinks[$Resp.id] = $NextLink
                        } else {
                            $GlobalNextLinks.Add("$($Resp.id)|$NextLink")
                        }
                    }
                } else {
                    $ErrorCode = $Resp.body.error.code
                    $ErrorMessage = $Resp.body.error.message

                    if ($Resp.status -in $RetryableStatusCodes) {
                        if ($BatchRequestsById.ContainsKey([string]$Resp.id)) {
                            $FailedRequests += $BatchRequestsById[[string]$Resp.id]
                        }
                        $LastRetryableErrors[[string]$Resp.id] = @{
                            status = $Resp.status
                            errorCode = $ErrorCode
                            errorMessage = $ErrorMessage
                        }

                        if (-not $Silent) {
                            if ($Resp.status -eq 429) {
                                Write-Host ("[i] Request ID {0} was throttled (429). Retrying automatically in {1}s (attempt {2}/{3}). No action needed." -f $Resp.id, $RetryDelaySeconds, ($RetryCount + 1), $MaxRetries)
                            } else {
                                Write-Host ("[i] Request ID {0} hit a temporary Graph error ({1}). Retrying automatically in {2}s (attempt {3}/{4})." -f $Resp.id, $Resp.status, $RetryDelaySeconds, ($RetryCount + 1), $MaxRetries)
                            }
                        }
                    } else {
                        if (-not $Silent) {
                            Write-Host "[!] Graph Batch Request: ID $($Resp.id) failed with status $($Resp.status): $ErrorCode - $ErrorMessage"
                        }
                        $ErrorEntries[[string]$Resp.id] = @{
                            id = $Resp.id
                            status = $Resp.status
                            errorCode = $ErrorCode
                            errorMessage = $ErrorMessage
                            complete = $false
                            incompleteReason = 'RequestFailed'
                        }
                    }
                }
            }

            $PendingRequests = $FailedRequests
            if ($PendingRequests.Count -gt 0 -and ($RetryCount + 1) -lt $MaxRetries) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
            $RetryCount++
        } while ($PendingRequests.Count -gt 0 -and $RetryCount -lt $MaxRetries)

        # break inside the do-loop only leaves the do-loop; stop the batch phase here.
        if ($Aborted) { break }

        if ($PendingRequests.Count -gt 0) {
            foreach ($PendingRequest in $PendingRequests) {
                $RequestId = $PendingRequest.id
                $LastError = $LastRetryableErrors[[string]$RequestId]
                $LastStatus = if ($null -ne $LastError) { $LastError.status } else { "unknown" }
                $LastErrorCode = if ($null -ne $LastError) { $LastError.errorCode } else { $null }
                $LastErrorMessage = if ($null -ne $LastError) { $LastError.errorMessage } else { "Retry attempts exhausted." }

                if (-not $Silent) {
                    if ($LastStatus -eq 429) {
                        Write-Warning ("[!] Request ID {0} remained throttled after {1} retries." -f $RequestId, $MaxRetries)
                    } else {
                        Write-Warning ("[!] Request ID {0} still failed with status {1} after {2} retries: {3} - {4}" -f $RequestId, $LastStatus, $MaxRetries, $LastErrorCode, $LastErrorMessage)
                    }
                }

                $ErrorEntries[[string]$RequestId] = @{
                    id = $RequestId
                    status = $LastStatus
                    errorCode = $LastErrorCode
                    errorMessage = $LastErrorMessage
                    complete = $false
                    incompleteReason = 'RequestFailed'
                }
            }
        }

        if ($BatchDelay -gt 0) {
            Start-Sleep -Seconds $BatchDelay
        }
    }

	 while (-not $Aborted -and -not $DisablePagination -and $GlobalNextLinks.Count -gt 0) {
		$ToFetch = @($GlobalNextLinks[0..([math]::Min($MaxBatchSize - 1, $GlobalNextLinks.Count - 1))])
		$GlobalNextLinks.RemoveRange(0, $ToFetch.Count)

		$Links = @($ToFetch | ForEach-Object { ($_ -split '\|', 2)[1] })
		$Ids   = @($ToFetch | ForEach-Object { ($_ -split '\|', 2)[0] })

		$BatchResult = Invoke-GraphNextLinkBatch -NextLinks $Links `
            -Ids $Ids `
            -AccessToken $AccessToken `
            -AccessTokenProvider $AccessTokenProvider `
            -UserAgent $UserAgent `
            -JsonDepthRequest $JsonDepthRequest `
            -JsonDepthResponse $JsonDepthResponse `
            -Proxy $Proxy `
            -VerboseMode:$VerboseMode `
            -DebugMode:$DebugMode `
            -Silent:$Silent `
            -SkipCertificateCheck:$SkipCertificateCheck `
            -MaxBatchSize $MaxBatchSize `
            -MaxRetries $MaxRetries `
            -RetryableStatusCodes $RetryableStatusCodes `
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

        # A page that could not be retrieved makes that ID's result truncated. Record it as a
        # field rather than leaving the caller to parse the warning stream.
        foreach ($id in $BatchResult.failures.Keys) {
            $IncompleteIds[$id] = 'PaginationFailed'
            $PaginationFailures[$id] = $BatchResult.failures[$id]
            $FailedLink = $BatchResult.failures[$id].nextLink
            if ($FailedLink) { $UnresolvedNextLinks[$id] = $FailedLink }
        }
	}

    # Anything still queued was never drained (aborted batch phase): those results are truncated.
    foreach ($Pending in $GlobalNextLinks) {
        $PendingParts = $Pending -split '\|', 2
        $IncompleteIds[$PendingParts[0]] = 'PaginationFailed'
        $UnresolvedNextLinks[$PendingParts[0]] = $PendingParts[1]
    }

    foreach ($id in $PagedResultsMap.Keys) {
        $Entry = @{ id = $id; status = 200; response = @{ value = $PagedResultsMap[$id].ToArray() } }
        if ($IncompleteIds.ContainsKey($id)) {
            $Entry['complete'] = $false
            $Entry['incompleteReason'] = $IncompleteIds[$id]
            if ($PaginationFailures.ContainsKey($id)) {
                $Entry['paginationFailureStatus'] = $PaginationFailures[$id].status
                $Entry['paginationFailureErrorCode'] = $PaginationFailures[$id].errorCode
            }
            if ($UnresolvedNextLinks.ContainsKey($id)) { $Entry['nextLink'] = $UnresolvedNextLinks[$id] }
        } else {
            $Entry['complete'] = $true
        }
        $DirectResults[$id] = $Entry
    }

    foreach ($id in $DirectResults.Keys) {
        if (-not $DirectResults[$id].ContainsKey('complete')) { $DirectResults[$id]['complete'] = $true }
    }

    # One entry per requested ID, in submission order. Anything never answered is reported
    # explicitly so a caller can validate by comparing ID sets rather than inferring.
    $Results = New-Object 'System.Collections.Generic.List[object]'
    $Emitted = @{}
    foreach ($id in $RequestedIds) {
        $Key = [string]$id
        if ($Emitted.ContainsKey($Key)) { continue }
        $Emitted[$Key] = $true

        if ($ErrorEntries.ContainsKey($Key)) {
            $Results.Add($ErrorEntries[$Key])
        } elseif ($DirectResults.ContainsKey($Key)) {
            $Results.Add($DirectResults[$Key])
        } else {
            $Results.Add(@{
                id = $id
                status = "unknown"
                errorCode = $null
                errorMessage = "Request was not attempted or no response was returned."
                complete = $false
                incompleteReason = 'NotAttempted'
            })
        }
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
        [scriptblock]$AccessTokenProvider,
        [string]$UserAgent,
        [int]$JsonDepthResponse = 10,
        [int]$JsonDepthRequest = 10,
        [int]$MaxRetries = 6,
        [int[]]$RetryableStatusCodes = @(429, 500, 502, 503, 504),
        [string]$Proxy,
        [switch]$SkipCertificateCheck,
        [ValidateRange(1, 20)]
        [int]$MaxBatchSize = 20,
		[ref]$HttpRequestCount,
		[ref]$SubRequestCount,
        [switch]$VerboseMode,
		[switch]$DebugMode,
        [switch]$Silent,
        [string]$ApiVersion
    )

    $ResultMap = @{}
    $MoreLinksMap = @{}
    $FailureMap = @{}

    $SupportsSkipCertificateCheck = (Get-Command Invoke-RestMethod).Parameters.ContainsKey('SkipCertificateCheck')

    for ($Offset = 0; $Offset -lt $NextLinks.Count; $Offset += $MaxBatchSize) {
        $BatchSet = @($NextLinks[$Offset..([math]::Min($Offset + $MaxBatchSize - 1, $NextLinks.Count - 1))])
        $BatchRequests = @()
        $index = 0

        foreach ($link in $BatchSet) {
            $relativeUrl = $link -replace '^https://graph\.microsoft\.com/[^/]+', ''
            # Carry the sub-batch offset in the ID so mapping back to the caller's IDs stays
            # correct when more than MaxBatchSize links are passed in one call.
            $BatchRequests += @{
                id     = "nl_$($Offset + $index)"
                method = "GET"
                url    = $relativeUrl
            }
            $index++
        }

        $BatchBody = @{ requests = $BatchRequests } | ConvertTo-Json -Depth $JsonDepthRequest

        $BatchResp = $null
        $Attempt = 0
        $LastTransportError = $null
        while ($true) {
            $CurrentToken = Resolve-GraphAccessToken -AccessToken $AccessToken -AccessTokenProvider $AccessTokenProvider
            $irmParams = @{
                Uri         = "https://graph.microsoft.com/$ApiVersion/`$batch"
                Method      = 'POST'
                Headers     = @{
                    "Authorization" = "Bearer $CurrentToken"
                    "User-Agent"    = $UserAgent
                    "Content-Type"  = "application/json"
                }
                Body        = $BatchBody
                ErrorAction = 'Stop'
            }
            if ($Proxy) { $irmParams['Proxy'] = $Proxy }
            if ($SkipCertificateCheck -and $SupportsSkipCertificateCheck) { $irmParams['SkipCertificateCheck'] = $true }

            try {
                if ($DebugMode) { Write-Host "[i] Sending nextLink batch request..." }
                $HttpRequestCount.Value++
                $SubRequestCount.Value += $BatchRequests.Count
                $BatchResp = Invoke-RestMethod @irmParams
                break
            } catch {
                $LastTransportError = $_
                $HttpStatus = Get-GraphErrorStatusCode -ErrorRecord $_
                $IsRetryable = ($null -eq $HttpStatus) -or ($HttpStatus -in $RetryableStatusCodes)

                if (-not $IsRetryable -or ($Attempt + 1) -ge $MaxRetries) { break }

                $Delay = [math]::Pow(2, $Attempt)
                if (-not $Silent) {
                    Write-Host ("[i] nextLink batch failed ({0}). Retrying in {1}s (attempt {2}/{3})." -f `
                        $(if ($null -ne $HttpStatus) { $HttpStatus } else { 'transport error' }), $Delay, ($Attempt + 1), $MaxRetries)
                }
                Start-Sleep -Seconds $Delay
                $Attempt++
            }
        }

        if ($null -eq $BatchResp) {
            # Whole sub-batch lost: every ID in it is truncated, and each keeps its continuation
            # link so the caller can resume.
            Write-Error "Failed nextLink batch: $LastTransportError "
            for ($k = 0; $k -lt $BatchSet.Count; $k++) {
                $realId = $Ids[$Offset + $k]
                $FailureMap[$realId] = @{
                    status = (Get-GraphErrorStatusCode -ErrorRecord $LastTransportError)
                    errorCode = 'BatchRequestFailed'
                    errorMessage = "$LastTransportError"
                    nextLink = $BatchSet[$k]
                }
            }
            continue
        }

        # Assemble inside the loop so every sub-batch contributes.
        foreach ($resp in $BatchResp.responses) {
            $SubIndex = [int]($resp.id -replace 'nl_', '')
            $realId = $Ids[$SubIndex]

            if ($resp.status -ge 200 -and $resp.status -lt 300) {
                if (-not $ResultMap.ContainsKey($realId)) {
                    $ResultMap[$realId] = New-Object 'System.Collections.Generic.List[object]'
                }
                if ($resp.body.value) {
                    $ResultMap[$realId].AddRange(@($resp.body.value))
                }
                $NextLink = Get-GraphMember -Object $resp.body -Name '@odata.nextLink'
                if ($NextLink) { $MoreLinksMap[$realId] = $NextLink }
            } else {
                if (-not $Silent) {
                    Write-Warning "NextLink subrequest failed: ID $($resp.id) ($($resp.status))"
                }
                $FailureMap[$realId] = @{
                    status = $resp.status
                    errorCode = $resp.body.error.code
                    errorMessage = $resp.body.error.message
                    nextLink = $NextLinks[$SubIndex]
                }
            }
        }
    }

    return @{
        values     = $ResultMap
        nextLinks  = $MoreLinksMap
        failures   = $FailureMap
    }
}


function Resolve-GraphAccessToken {
    param (
        [string]$AccessToken,
        [scriptblock]$AccessTokenProvider
    )

    if ($AccessTokenProvider) {
        $Token = [string](& $AccessTokenProvider)
        if ([string]::IsNullOrWhiteSpace($Token)) {
            throw "AccessTokenProvider returned an empty token."
        }
        return $Token
    }
    return $AccessToken
}


function Get-GraphErrorStatusCode {
    # Returns the HTTP status of a failed Invoke-RestMethod call, or $null for transport-level
    # failures. The exception type differs between Windows PowerShell 5.1 (WebException) and
    # PowerShell 7 (HttpResponseException / HttpRequestException), so match on the Response
    # property rather than on the type or the error id.
    param ($ErrorRecord)

    $Exception = $ErrorRecord.Exception
    if ($null -eq $Exception) { return $null }

    $ResponseProperty = $Exception.PSObject.Properties['Response']
    if ($null -eq $ResponseProperty -or $null -eq $ResponseProperty.Value) { return $null }

    $StatusProperty = $ResponseProperty.Value.PSObject.Properties['StatusCode']
    if ($null -eq $StatusProperty -or $null -eq $StatusProperty.Value) { return $null }

    try { return [int]$StatusProperty.Value } catch { return $null }
}


function Test-GraphMember {
    # Property presence, not truthiness: an empty collection is still a collection.
    param ($Object, [string]$Name)

    if ($null -eq $Object) { return $false }
    if ($Object -is [System.Collections.IDictionary]) { return $Object.Contains($Name) }
    return $null -ne $Object.PSObject.Properties[$Name]
}


function Get-GraphMember {
    param ($Object, [string]$Name)

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $Property = $Object.PSObject.Properties[$Name]
    if ($Property) { return $Property.Value }
    return $null
}


function Get-GraphEffectiveRequest {
    # Shallow copy with query parameters applied, so the caller's request objects are never
    # modified and the same request list can be submitted more than once.
    param ($Request, [hashtable]$QueryParameters)

    $Copy = @{}
    if ($Request -is [System.Collections.IDictionary]) {
        foreach ($Key in $Request.Keys) { $Copy[$Key] = $Request[$Key] }
    } else {
        foreach ($Property in $Request.PSObject.Properties) { $Copy[$Property.Name] = $Property.Value }
    }

    $RequestParams = $null
    if ($Copy.ContainsKey('queryParameters')) {
        $RequestParams = $Copy['queryParameters']
        # Not a Graph batch field; it must not be serialised into the request.
        $Copy.Remove('queryParameters')
    }

    $EffectiveParams = @{}
    if ($RequestParams) {
        foreach ($Key in $RequestParams.Keys) { $EffectiveParams[$Key] = $RequestParams[$Key] }
    }
    if ($QueryParameters) {
        foreach ($Key in $QueryParameters.Keys) {
            if (-not $EffectiveParams.ContainsKey($Key)) {
                $EffectiveParams[$Key] = $QueryParameters[$Key]
            }
        }
    }

    if ($EffectiveParams.Count -gt 0) {
        $QueryString = ($EffectiveParams.GetEnumerator() | ForEach-Object {
            "$($_.Key)=$([uri]::EscapeDataString($_.Value))"
        }) -join '&'

        if ($Copy.url -notmatch "\?") {
            $Copy.url = "$($Copy.url)?$QueryString"
        } else {
            $Copy.url = "$($Copy.url)&$QueryString"
        }
    }

    return $Copy
}


# Keep the public surface as it was; the helpers above are internal.
Export-ModuleMember -Function Send-GraphBatchRequest, Invoke-GraphNextLinkBatch
