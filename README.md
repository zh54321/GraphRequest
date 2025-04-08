# GraphRequest - PowerShell Module

## Introduction

The `GraphRequest` PowerShell module allows to send single requests to the Microsoft Graph API.
It supports automatic throttling handling, pagination, and can return results in either PowerShell object format or raw JSON.

Note: 
- Cleartext access tokens can be obtained, for example, using [EntraTokenAid](https://github.com/zh54321/EntraTokenAid).
- Use [GraphBatchRequest](https://github.com/zh54321/GraphBatchRequest) to send Batch Request to the Graph API.

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
| `-AdditionalHeaders`         | Add additional HTTP headers (e.g. for ConsistencyLevel)                                     |
| `-JsonDepthResponse` *(Default: 10)* | Specifies the depth for JSON conversion (request). Useful for deeply nested objects in combination with `-RawJson`.  |

## Examples

### Example 1: **Retrieve All Users**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Response = Send-MSGraphRequest -AccessToken $AccessToken -Method GET -Uri '/users'

#Show Data
$Response
```

### Example 2: **Create a New Microsoft 365 Group**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
$Body = @{
	displayName     = "TestSecurityGroup"
	mailEnabled     = $false
	mailNickname    = "TestSecurityGroup$(Get-Random)"
	securityEnabled = $true
	groupTypes      = @()
}
$Response = Send-MSGraphRequest -AccessToken $AccessToken -Method POST -Uri "/groups" -Body $Body"

#Show Response
$Response
```

### Example 3: **Use the Beta API Endpoint and a simple query**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
Send-MSGraphRequest -AccessToken $AccessToken -Method GET -Uri '/groups?$select=displayName' -BetaAPI -Proxy "http://127.0.0.1:8080"
```

### Example 4: **Using an advanced filter and a proxy**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
Send-MSGraphRequest -AccessToken $AccessToken -Method GET -Uri '/users' -QueryParameters @{ '$filter' = "startswith(displayName,'Alex')" } -Proxy "http://127.0.0.1:8080"
```

### Example 5: **Using an additional header**

```powershell
$AccessToken = "YOUR_ACCESS_TOKEN"
Send-MSGraphRequest -AccessToken $AccessToken -Method GET -Uri '/users' -AdditionalHeaders @{ 'ConsistencyLevel' = 'eventual' }
```

## Notes

- Ensure that you have **valid Microsoft Graph API permissions** before executing requests.
- The module automatically handles **the 429 throttling errors** using **exponential backoff**.
