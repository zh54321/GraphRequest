<#
.SYNOPSIS
    Sends requests to the Microsoft Graph API.

.DESCRIPTION
    Sends single requests to Microsoft Graph API with automatic pagination, retries, and exponential backoff.

.PARAMETER AccessToken
    OAuth token for Microsoft Graph authentication.

.PARAMETER Method
    HTTP method (GET, POST, PUT, PATCH, DELETE).

.PARAMETER Uri
    API resource URI (relative, e.g., /users).

.PARAMETER Body
    Request body as PowerShell hashtable/object.

.PARAMETER MaxRetries
    Maximum retry attempts. Default: 5.

.PARAMETER BetaAPI
    Switch to use Beta API endpoint.

.PARAMETER RawJson
    Returns the raw JSON string instead of a PowerShell object.

.PARAMETER Proxy
    Proxy URL (e.g., http://proxyserver:port).

.PARAMETER UserAgent
    Specify the User-Agent HTTP header.

.PARAMETER VerboseMode
    Enable detailed logging.

.PARAMETER AdditionalHeaders
    Add additional headers, example -AdditionalHeaders @{ 'ConsistencyLevel' = 'eventual' }

.PARAMETER QueryParameters
    Query parameters for more complex queries (e.g. -QueryParameters @{ '$filter' = "startswith(displayName,'Alex')" })

.EXAMPLE
    Send-MSGraphRequest -AccessToken $token -Method GET -Uri '/users'

.EXAMPLE
    Send-MSGraphRequest -AccessToken $token -Method GET -Uri '/groups?$select=displayName' -VerboseMode

.EXAMPLE
    Send-MSGraphRequest -AccessToken $token -Method GET -Uri '/groups?$select=displayName' -proxy "http://127.0.0.1:8080"

.EXAMPLE
    Send-MSGraphRequest -AccessToken $token -Method GET -Uri '/users' -AdditionalHeaders @{ 'ConsistencyLevel' = 'eventual' }

.EXAMPLE 
    Send-MSGraphRequest -AccessToken $token -Method GET -Uri '/users' -QueryParameters @{ '$filter' = "startswith(displayName,'Alex')" } -Proxy "http://127.0.0.1:8080"

.EXAMPLE 
    $Body = @{
        displayName     = "Test Security Group2"
        mailEnabled     = $false
        mailNickname    = "TestSecurityGroup$(Get-Random)"
        securityEnabled = $true
        groupTypes      = @()
    }
    $result = Send-MSGraphRequest -AccessToken $token -Method POST -Uri "/groups" -Body $Body -VerboseMode -Proxy "http://127.0.0.1:8080"
    

.NOTES
    Author: ZH54321
    GitHub: https://github.com/zh54321/GraphRequest
#>

function Send-MSGraphRequest {
    param (
        [Parameter(Mandatory)]
        [string]$AccessToken,

        [Parameter(Mandatory)]
        [ValidateSet("GET", "POST", "PATCH", "PUT", "DELETE")]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,
        [hashtable]$Body,
        [int]$MaxRetries = 5,
        [switch]$BetaAPI,
        [string]$UserAgent = 'PowerShell GraphRequest Module',
        [switch]$RawJson,
        [string]$Proxy,
        [switch]$VerboseMode,
        [hashtable]$QueryParameters,
        [hashtable]$AdditionalHeaders,
        [int]$JsonDepthResponse = 10
    )

    $ApiVersion = if ($BetaAPI) { "beta" } else { "v1.0" }
    $BaseUri = "https://graph.microsoft.com/$ApiVersion"
    $FullUri = "$BaseUri$Uri"

    #Add query parameters
    if ($QueryParameters) {
        $QueryString = ($QueryParameters.GetEnumerator() | 
            ForEach-Object { 
                "$($_.Key)=$([uri]::EscapeDataString($_.Value))" 
            }) -join '&'
        $FullUri = "$FullUri`?$QueryString"
    }
    

    #Define basic headers
    $Headers = @{
        Authorization  = "Bearer $AccessToken"
        'Content-Type' = 'application/json'
        'User-Agent'   = $UserAgent
    }

    #Add custom headers if required
    if ($AdditionalHeaders) {
        $Headers += $AdditionalHeaders
    }

    $RetryCount = 0
    $Results = @()

    # Prepare Invoke-RestMethod parameters
    $irmParams = @{
        Uri             = $FullUri
        Method          = $Method
        Headers         = $Headers
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }

    if ($Body) {
        $irmParams.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }

    if ($Proxy) {
        $irmParams.Proxy = $Proxy
    }

    do {
        try {
            if ($VerboseMode) { Write-Host "[*] Request [$Method]: $FullUri" }

            $Response = Invoke-RestMethod @irmParams

            if ($Response.value) {
                $Results += $Response.value
            } else {
                $Results += $Response
            }

            # Pagination handling
            while ($Response.'@odata.nextLink') {
                if ($VerboseMode) { Write-Host "[*] Following pagination link: $($Response.'@odata.nextLink')" }

                $irmParams.Uri = $Response.'@odata.nextLink'
                # Remove Body for paginated GET requests
                $irmParams.Remove('Body')

                $Response = Invoke-RestMethod @irmParams

                if ($Response.value) {
                    $Results += $Response.value
                } else {
                    $Results += $Response
                }
            }

            break
        }
        catch {
            $StatusCode = $_.Exception.Response.StatusCode.value__
            $StatusDesc = $_.Exception.Message

            Write-Host "[!] Error: [$StatusCode] $StatusDesc"

            if ($StatusCode -in @(429,500,502,503,504) -and $RetryCount -lt $MaxRetries) {
                $RetryAfter = $_.Exception.Response.Headers['Retry-After']
                if ($RetryAfter) {
                    Write-Host "[i] Throttled. Retrying after $RetryAfter seconds..."
                    Start-Sleep -Seconds ([int]$RetryAfter)
                } elseif ($RetryCount -eq 0) {
                    Write-Host "[*] Retrying immediately..."
                    Start-Sleep -Seconds 0
                } else {
                    $Backoff = 2 * $RetryCount
                    Write-Host "[*] Retrying in $Backoff seconds..."
                    Start-Sleep -Seconds $Backoff
                }
                $RetryCount++
            }
            else {
                Write-Host "[!] Failed after $RetryCount retries."
                return
            }
        }
    } while ($RetryCount -le $MaxRetries)

    if ($RawJson) {
        return $Results | ConvertTo-Json -Depth $JsonDepthResponse
    }
    else {
        return $Results
    }
}
