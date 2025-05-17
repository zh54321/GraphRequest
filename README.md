# GraphRequest - PowerShell Module

## Introduction

The `GraphRequest` PowerShell module allows to send single requests to the Microsoft Graph API.

Key features:
- Handles Microsoft Graph v1.0 and Beta APIs
- Automatic Pagination Support
- Retry Logic with exponential Backoff for transient errors (e.g. 429, 503)
- Custom Headers and Query Parameters for flexible API queries
- Optional Raw JSON Output
- Simple HTTP Proxy Support (for debugging)
- Verbose Logging Option
- User-Agent Customization

Note: 
- Cleartext access tokens can be obtained, for example, using [EntraTokenAid](https://github.com/zh54321/EntraTokenAid).
- Use [GraphBatchRequest](https://github.com/zh54321/GraphBatchRequest) to send Batch Request to the Graph API.
- Find first party apps with pre-consented scopes to bypass Graph API consent [GraphPreConsentExplorer](https://github.com/zh54321/GraphPreConsentExplorer)

## Parameters

| Parameter                    | Description                                                                                 |
| ---------------------------- | ------------------------------------------------------------------------------------------- |
| `-AccessToken` *(Mandatory)* | The OAuth access token to authenticate against Microsoft Graph API.                         |
| `-Method`      *(Mandatory)* | HTTP Method to use (GET, POST, PATCH, PUT, DELETE)                                          |
| `-Uri`         *(Mandatory)* | Relative Graph URI (e.g. /users)                                                            |
| `-VerboseMode`               | Enables verbose logging to provide additional information about request processing.         |
| `-UserAgent`                 | Custom UserAgent                                                                            |
| `-BetaAPI`                   | If specified, uses the Microsoft Graph `Beta` endpoint instead of `v1.0`.                   |
| `-RawJson`                   | If specified, returns the response as a raw JSON string instead of a PowerShell object.     |
| `-MaxRetries` *(Default: 5)* | Specifies the maximum number of retry attempts for failed requests.                         |
| `-Proxy`                     | Use a Proxy (e.g. http://127.0.0.1:8080)                                                    |
| `-Body`					   | Request body as PowerShell hashtable/object (will be converted to JSON).                    |
| `-QueryParameters`           | Query parameters for more complex queries                                                   |
| `-DisablePagination`         | Prevents the function from automatically following @odata.nextLink for paginated results.   |
| `-AdditionalHeaders`         | Add additional HTTP headers (e.g. for ConsistencyLevel)                                     |
| `-JsonDepthResponse` *(Default: 10)* | Specifies the depth for JSON conversion (request). Useful for deeply nested objects in combination with `-RawJson`.  |
| `-$Suppress404`              | Supress 404 Messages (example if a queried User object is not found in the tenant)          |          

## Examples

### Example 1: **Retrieve All Users**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Response = Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri '/users'

#Show Data
$Response
```

### Example 2: **Create a New Security Group**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Body = @{
	displayName     = "TestSecurityGroup"
	mailEnabled     = $false
	mailNickname    = "TestSecurityGroup$(Get-Random)"
	securityEnabled = $true
	groupTypes      = @()
}
$Response = Send-GraphRequest -AccessToken $AccessToken -Method POST -Uri "/groups" -Body $Body"

#Show Response
$Response
```

### Example 3: **Use the Beta API Endpoint and a simple query**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri '/groups?$select=displayName' -BetaAPI
```

### Example 4: **Using an advanced filter and a proxy**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri '/users' -QueryParameters @{ '$filter' = "startswith(displayName,'Alex')" } -Proxy "http://127.0.0.1:8080"
```

### Example 5: **Using an additional header**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri '/users' -AdditionalHeaders @{ 'ConsistencyLevel' = 'eventual' }
```

### Example 6: **Using an additional header to remove odata metadata**
```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$QueryParameters = @{
    '$select' = "Id,DisplayName,IsMemberManagementRestricted"
}
$headers = @{ 
    'Accept' = 'application/json;odata.metadata=none' 
}
Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri "/directory/administrativeUnits" -QueryParameters $QueryParameters -AdditionalHeaders $headers 
```
### Example 7: **Catch errors**
```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
try {
    Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri '/doesnotexist' -BetaAPI -ErrorAction Stop
} catch {
    $err = $_
    Write-Host "[!] Auth error occurred:"
    Write-Host "  Message     : $($err.Exception.Message)"
    Write-Host "  FullyQualifiedErrorId : $($err.FullyQualifiedErrorId)"
    Write-Host "  TargetURL: $($err.TargetObject)"
    Write-Host "  Category    : $($err.CategoryInfo.Category)"
    Write-Host "  Script Line : $($err.InvocationInfo.Line)"
}
```
### Example 8: **Get only one result by disabling pagination**
```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$QueryParameters = @{
    '$select' = "id,SignInActivity"
    '$top' = "1"
}
Send-GraphRequest -AccessToken $AccessToken -Method GET -Uri "/users" -QueryParameters $QueryParameters -DisablePagination
```


