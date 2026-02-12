[CmdletBinding()]
param(
    [string]$configFile = 'config.json'
)

#region helper functions
function Invoke-GraphAPI()
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$accessToken,
        [Parameter(Mandatory = $true)]
        [object]$ResourcePath,  # Can be string or string array for batch processing
        [Parameter()]
        [ValidateSet('v1.0', 'beta')]
        [string]$APIVersion = 'beta',
        [string]$method = 'get',
        [string]$Filter = $null,
        [string]$Search = $null,
        [string]$ExtraParameters = $null,
        $headers,
        [string]$body = $null,
        [switch]$consistencyLevel,
        [switch]$secureString
    )

    #region variables and logs
    $functionName = $MyInvocation.MyCommand.Name
    if ($accessToken)
    {
        Write-Verbose "[$functionName] Access token provided."
    }
    else
    {
        Write-Verbose "[$functionName] Access token not provided. Please provide a valid access token."
        return
    }

    # Check if ResourcePath is an array
    $isArrayInput = $ResourcePath -is [array]
    Write-Verbose "[$functionName] isArrayInput: $isArrayInput"
    # Handle single-item array
    if ($isArrayInput -and $ResourcePath.Count -eq 1)
    {
        Write-Verbose "[$functionName] Single-item array detected, processing as single request"
        $ResourcePath = $ResourcePath[0]
        $isArrayInput = $false
    }
    # Check if batch processing is requested (array with multiple items)
    $isBatchRequest = $isArrayInput -and $ResourcePath.Count -gt 1
    $batchThreshold = 1
    Write-Verbose "[$functionName] isBatchRequest: $isBatchRequest with a threshold of $batchThreshold"
    if ($isBatchRequest -and $ResourcePath.Count -ge $batchThreshold)
    {
        Write-Verbose "[$functionName] Batch request detected: $($ResourcePath.Count) resources"
        # Attempt to use native Graph API $batch endpoint
        # Graph API supports up to 20 requests per batch
        $maxBatchSize = 20
        $allResults = @()
        $successCount = 0
        $failureCount = 0
        # Split requests into batches of max 20
        $batches = @()
        for ($i = 0; $i -lt $ResourcePath.Count; $i += $maxBatchSize)
        {
            $batchSize = [Math]::Min($maxBatchSize, $ResourcePath.Count - $i)
            $batches += , @($ResourcePath[$i..($i + $batchSize - 1)])
        }
        Write-Verbose "[$functionName] Processing $($ResourcePath.Count) requests in $($batches.Count) batch(es)"
        $batchIndex = 0
        foreach ($batch in $batches)
        {
            # Build batch request body according to Graph API spec
            $batchRequests = @()
            $requestId = 1
            foreach ($path in $batch)
            {
                # Build full URL for the request
                $requestUrl = "/$path"
                # Handle filters, search, and extra parameters in the URL
                $queryParams = @()
                if ($Filter)
                {
                    $queryParams += "`$filter=$([uri]::EscapeUriString($Filter))"
                }
                if ($Search)
                {
                    $queryParams += "`$search=$([uri]::EscapeUriString($Search))"
                }
                if ($ExtraParameters)
                {
                    $queryParams += $ExtraParameters
                }
                if ($queryParams.Count -gt 0)
                {
                    $requestUrl += "?" + ($queryParams -join "&")
                }
                # Build request object
                $batchRequest = @{
                    id     = $requestId.ToString()
                    method = $method.ToUpper()
                    url    = $requestUrl
                }
                # Add headers if needed
                if ($consistencyLevel)
                {
                    $batchRequest['headers'] = @{
                        'ConsistencyLevel' = 'eventual'
                    }
                }
                # Add body if provided
                if ($body)
                {
                    $batchRequest['body'] = $body | ConvertFrom-Json
                    if (-not $batchRequest.ContainsKey('headers'))
                    {
                        $batchRequest['headers'] = @{}
                    }
                    $batchRequest['headers']['Content-Type'] = 'application/json'
                }
                $batchRequests += $batchRequest
                $requestId++
            }
            # Create batch request body
            $batchBody = @{
                requests = $batchRequests
            } | ConvertTo-Json -Depth 10
            # Send batch request to Graph API
            try
            {
                $batchHeaders = @{
                    'Authorization' = "Bearer $accessToken"
                    'Content-Type'  = 'application/json'
                }
                $batchUri = "https://graph.microsoft.com/$APIVersion/`$batch"
                $batchResponse = Invoke-RestMethod -Uri $batchUri -Method Post -Headers $batchHeaders -Body $batchBody -UseBasicParsing
                # Process batch responses
                # Renumber response IDs to be globally unique across all batches
                $globalIdOffset = $batchIndex * $maxBatchSize
                foreach ($response in $batchResponse.responses)
                {
                    # Adjust the response ID to be globally unique (1-240 instead of 1-20 per batch)
                    $globalId = ([int]$response.id) + $globalIdOffset
                    $response.id = $globalId

                    if ($response.status -ge 200 -and $response.status -lt 300)
                    {
                        # Preserve the entire response object so downstream code can match by id
                        $allResults += $response
                        $successCount++
                    }
                    else
                    {
                        # Include failed responses so downstream code can handle them properly
                        $allResults += $response
                        $failureCount++
                        $errorMsg = if ($response.body.error)
                        {
                            $response.body.error.message
                        }
                        else
                        {
                            "Unknown error"
                        }
                        Write-Verbose "[$functionName] Batch request ID $($response.id) failed with status $($response.status): $errorMsg"
                    }
                }
                $batchIndex++
            }
            catch
            {
                # Final fallback: process each resource path individually
                foreach ($path in $batch)
                {
                    # Recursive call with single resource path
                    $result = Invoke-GraphAPI -accessToken $accessToken -ResourcePath $path -APIVersion $APIVersion `
                        -method $method -Filter $Filter -Search $Search -ExtraParameters $ExtraParameters `
                        -body $body -consistencyLevel:$consistencyLevel -secureString:$secureString
                    # Check if result is an error status code (integer) or null
                    if ($null -eq $result -or $result -is [int])
                    {
                        $failureCount++
                    }
                    else
                    {
                        $allResults += $result
                        $successCount++
                    }
                }
            }
        }
        Write-Verbose "[$functionName] Batch processing completed: $successCount successful, $failureCount failed"
        # Return combined results
        return @{
            value          = $allResults
            batchProcessed = $true
            batchMethod    = if ($useBatchProcessor)
            {
                "GraphCore"
            }
            else
            {
                "NativeBatch"
            }
            successCount   = $successCount
            failureCount   = $failureCount
            totalCount     = $ResourcePath.Count
        }
    }

    # Single request processing (original behavior continues below)
    $uri = "https://graph.microsoft.com/$APIVersion/$ResourcePath"
    $statusCode = $null
    Write-Verbose "[$functionName] Uri: $uri"
    #endregion

    #region Encode filter and add headers
    if ($Filter)
    {
        Write-Verbose "[$functionName] Splitting filter by logical operators while preserving operators."
        $filterParts = [System.Collections.ArrayList]::new()
        $logicalOperators = [System.Collections.ArrayList]::new()
        # Pattern to match a logical operator with surrounding spaces
        $pattern = '\s+(and|or)\s+'
        $lastIndex = 0
        # Find all logical operators and their positions
        $logicalOperaterMatches = [regex]::Matches($Filter, $pattern)
        Write-Verbose "[$functionName] Found $($logicalOperaterMatches.Count) logical operators."
        # If no logical operators, process as a single condition
        if ($logicalOperaterMatches.Count -eq 0)
        {
            Write-Verbose "[$functionName] No logical operators found. Processing as a single filter condition."
            $processedFilter = ProcessFilterCondition -condition $Filter
            Write-Verbose "[$functionName] Processed single filter condition: $processedFilter"
            $encodedFilter = $processedFilter
            Write-Verbose "[$functionName] Encoded filter: $encodedFilter"
        }
        else
        {
            # Process each part of the filter
            Write-Verbose "[$functionName] Logical operators found. Processing filter as multiple conditions."
            foreach ($logicalOperatorMatch in $logicalOperaterMatches)
            {
                Write-Verbose "[$functionName] Processing filter condition before logical operator: $($Filter.Substring($lastIndex, $logicalOperatorMatch.Index - $lastIndex))"
                $condition = $Filter.Substring($lastIndex, $logicalOperatorMatch.Index - $lastIndex)
                Write-Verbose "[$functionName] Condition to process: $condition"
                [void]$filterParts.Add((ProcessFilterCondition -condition $condition))
                Write-Verbose "[$functionName] Processed filter condition: $($filterParts[$filterParts.Count - 1])"
                # Store the logical operator (and, or)
                [void]$logicalOperators.Add($logicalOperatorMatch.Value.Trim())
                $lastIndex = $logicalOperatorMatch.Index + $logicalOperatorMatch.Length
                Write-Verbose "[$functionName] Logical operators so far: $($logicalOperators -join ', ')"
            }
            # Don't forget the last part after the last logical operator
            if ($lastIndex -lt $Filter.Length)
            {
                Write-Verbose "[$functionName] Processing filter condition after the last logical operator."
                $condition = $Filter.Substring($lastIndex)
                [void]$filterParts.Add((ProcessFilterCondition -condition $condition))
                Write-Verbose "[$functionName] Processed filter condition: $($filterParts[$filterParts.Count - 1])"
            }
            # Rebuild the filter string with processed parts and original logical operators
            Write-Verbose "[$functionName] Rebuilding the filter string with processed parts and logical operators."
            $encodedFilter = $filterParts[0]
            for ($i = 0; $i -lt $logicalOperators.Count; $i++)
            {
                $encodedFilter += " $($logicalOperators[$i]) $($filterParts[$i+1])"
                Write-Verbose "[$functionName] Adding logical operator: $($logicalOperators[$i])"
            }
            Write-Verbose "[$functionName] Processed complex filter: $encodedFilter"
        }
        $encodedUri = "$uri`?`$filter=$([uri]::EscapeUriString($encodedFilter))"
        Write-Verbose "[$functionName] Uri after applying filters: $encodedUri"
    }
    else
    {
        Write-Verbose "[$functionName] No filter provided."
        $encodedUri = $uri
    }

    # Handle search parameter
    if ($Search)
    {
        Write-Verbose "[$functionName] Processing search parameter: $Search"
        # URL encode the search string
        $encodedSearch = [uri]::EscapeUriString($Search)
        Write-Verbose "[$functionName] Encoded search: $encodedSearch"
        # Add search parameter to URI
        if ($encodedUri.Contains("?"))
        {
            $encodedUri = "$encodedUri&`$search=$encodedSearch"
        }
        else
        {
            $encodedUri = "$encodedUri`?`$search=$encodedSearch"
        }
        Write-Verbose "[$functionName] Uri after applying search: $encodedUri"
    }
    else
    {
        Write-Verbose "[$functionName] No search parameter provided."
    }

    if ($extraParameters)
    {
        Write-Verbose "[$functionName] Extra parameters provided."
        # Initialize the parameter list
        $paramsList = @()
        # Split by ampersand to get individual key-value pairs
        $keyValuePairs = $extraParameters -split '&'
        Write-Verbose "[$functionName] Found $($keyValuePairs.Count) key-value pairs."
        foreach ($pair in $keyValuePairs)
        {
            Write-Verbose "[$functionName] Processing key-value pair: $pair"
            # Split each pair by equals sign to separate key and value
            $keyAndValue = $pair -split '=', 2
            if ($keyAndValue.Count -eq 2)
            {
                $key = $keyAndValue[0].Trim()
                $value = $keyAndValue[1].Trim()
                Write-Verbose "[$functionName] Key: $key"
                Write-Verbose "[$functionName] Value: $value"
                # Add the $ prefix to the key for OData parameters
                $formattedKey = "`$$key"
                Write-Verbose "[$functionName] Formatted Key with $ prefix: $formattedKey"
                # Add the formatted parameter to the list
                $paramsList += "$formattedKey=$value"
            }
            else
            {
                Write-Warning "Invalid parameter format: $pair - skipping"
            }
        }
        Write-Verbose "[$functionName] Final parameter list:"
        $paramsList | ForEach-Object { Write-Verbose "[$functionName] $_" }
        # Join the parameters with & to create a complete query string
        $queryString = $paramsList -join '&'
        Write-Verbose "[$functionName] Final query string: $queryString"
        # Append the extra parameters to the URI
        if ($filter -or $Search)
        {
            Write-Verbose "[$functionName] Adding extra parameters to the uri along with existing parameters."
            $encodedUri = "$encodedUri`&$queryString"
        }
        else
        {
            Write-Verbose "[$functionName] No filter or search provided. Adding extra parameters to the uri."
            $encodedUri = "$encodedUri`?$queryString"
        }
    }
    else
    {
        Write-Verbose "[$functionName] No extra parameters provided."
    }
    # Build default headers with Authorization and Content-Type
    if ($consistencyLevel)
    {
        Write-Verbose "[$functionName] Adding consistency level to the headers."
        $defaultHeaders = @{
            Authorization    = "Bearer $accessToken"
            'Content-Type'   = 'application/json'
            ConsistencyLevel = 'Eventual'
        }
    }
    else
    {
        Write-Verbose "[$functionName] No consistency level provided."
        $defaultHeaders = @{
            Authorization  = "Bearer $accessToken"
            'Content-Type' = 'application/json'
        }
    }

    # Merge custom headers if provided (custom headers take precedence)
    if ($headers)
    {
        Write-Verbose "[$functionName] Custom headers provided. Merging with default headers."
        foreach ($key in $headers.Keys)
        {
            $defaultHeaders[$key] = $headers[$key]
            Write-Verbose "[$functionName] Added/Overridden header: $key"
        }
    }
    #endregion

    #region prepare the call
    # Create parameter hashtable for splatting
    Write-Verbose "[$functionName] Preparing parameters for Invoke-RestMethod call."
    $restParams = @{
        Method          = $method
        Uri             = $encodedUri
        Headers         = $defaultHeaders
        UseBasicParsing = $true
    }
    #add headers parameter if it was passed
    if ($headers)
    {
        Write-Verbose "[$functionName] Headers provided. Adding to the request."
        $restParams['Headers'] = $headers
    }
    # Only add Body parameter if it exists
    if ($body)
    {
        Write-Verbose "[$functionName] Body parameter provided. Adding to the request."
        $restParams['Body'] = $body
    }
    #Add statusCodeVariable if we are running under powershell  7.0 or higher
    if ($PSVersionTable.PSVersion.Major -ge 7)
    {
        Write-Verbose "[$functionName] PowerShell version is $($PSVersionTable.PSVersion.Major ). Adding StatusCodeVariable to the request."
        $restParams['StatusCodeVariable'] = 'statusCode'
    }
    Write-Verbose "[$functionName] Making the following call to Microsoft Graph:"
    Write-Verbose "[$functionName] URI: $encodedUri"
    Write-Verbose "[$functionName] Method: $method"
    #endregion
    try
    {
        $response = Invoke-RestMethod @restParams
        Write-Verbose "[$functionName] NextLink: $($response.'@odata.nextLink')"
        Write-Verbose "[$functionName] Response count: $($response.value.count)"
        if ($response.'@odata.nextLink')
        {
            Write-Verbose "[$functionName] NextLink found. Fetching additional pages."
            # Initialize an array to hold all items
            $allItems = @()
            $allItems += $response.value
            $nextLink = $response.'@odata.nextLink'
            while ($nextLink)
            {
                $nextGroup = Invoke-RestMethod -Method $method -Uri $nextLink -Headers $defaultHeaders -UseBasicParsing
                Write-Verbose "[$functionName] Fetched next page with $($nextGroup.value.Count) items."
                if ($nextGroup.value)
                {
                    Write-Verbose "[$functionName] Adding items from next page to the collection."
                    $allItems += $nextGroup.value
                }
                $nextLink = $nextGroup.'@odata.nextLink'
            }
            # Optionally, reconstruct a response object if needed
            $response.value = $allItems
            Write-Verbose "[$functionName] All items collected. Total count: $($Response.value.Count)"
        }
        else
        {
            Write-Verbose "[$functionName] No nextLink found. Single page response received."
        }
        Write-Verbose "[$functionName] The call was successful."
        if ($response.count)
        {
            Write-Verbose "[$functionName] Total count of items: $($response.count)."
        }
        if ($response.value.Count)
        {
            Write-Verbose "[$functionName] Number of items returned: $($response.value.Count)."
        }
        if ($PSVersionTable.PSVersion.Major -ge 7)
        {
            Write-Verbose "[$functionName] Status code: $statusCode"
        }
    }
    catch
    {
        # Capture as much diagnostic information as possible about the failure
        Write-Verbose "[$functionName] An error occurred while calling the Graph API. Capturing diagnostics."
        Write-Verbose "[$functionName] Exception type: $($PSItem.Exception.GetType().FullName)"
        # Walk inner exceptions (if any)
        $inner = $PSItem.Exception.InnerException
        while ($null -ne $inner)
        {
            Write-Verbose "[$functionName] InnerException type: $($inner.GetType().FullName)"
            Write-Verbose "[$functionName] InnerException message: $($inner.Message)"
            $inner = $inner.InnerException
        }
        # Defaults
        $statusDescription = $null
        $statusMessage = $PSItem.Exception.Message
        $statusCodeMessage = $null
        # Try to extract status code from exception when available
        if ($null -eq $PSItem.Exception.statusCode)
        {
            # Fallback: try to parse from exception message
            $statusCode = [regex]::Match($PSItem.Exception.Message, '\d+').Value
            Write-Verbose "[$functionName] Status code (parsed): $statusCode"
            $statusCodeMessage = $PSItem.Exception | Out-String
            Write-Verbose "[$functionName] Status code message: $statusCodeMessage"
        }
        else
        {
            # PowerShell 5.1/7 HttpStatusCode
            try
            {
                $statusCode = $PSItem.Exception.statuscode.value__
            }
            catch
            {
                $statusCode = [int]$PSItem.Exception.statuscode
            }
            $statusCodeMessage = $PSItem.Exception.statuscode
            Write-Verbose "[$functionName] Status code (from exception): $statusCode"
        }

        # Attempt to extract response details (headers/body) across PS versions
        $responseBodyRaw = $null
        $responseJson = $null
        $requestId = $null
        $clientRequestId = $null
        $serverDate = $null
        $retryAfter = $null
        $diagHeader = $null
        $responseHeaders = @{}
        $resp = $PSItem.Exception.Response
        if ($null -ne $resp)
        {
            # Status description when available
            try
            {
                $statusDescription = $resp.StatusDescription
            }
            catch
            {
                $statusDescription = $null
            }

            # Headers (handle both WebHeaderCollection and IDictionary-like)
            try
            {
                if ($resp.Headers -and $resp.Headers -is [System.Net.WebHeaderCollection])
                {
                    foreach ($key in $resp.Headers.AllKeys)
                    {
                        $responseHeaders[$key] = $resp.Headers[$key]
                    }
                }
                elseif ($resp.Headers)
                {
                    foreach ($kvp in $resp.Headers.GetEnumerator())
                    {
                        $responseHeaders[$kvp.Key] = ($kvp.Value -join ',')
                    }
                }
            }
            catch
            {
                Write-Verbose "[$functionName] Failed to enumerate response headers: $($_.Exception.Message)"
                & $logWarn "[$functionName] Failed to enumerate response headers: $($_.Exception.Message)"
            }

            # Common Graph headers
            if ($responseHeaders.ContainsKey('request-id'))
            {
                $requestId = $responseHeaders['request-id']
            }
            if ($responseHeaders.ContainsKey('client-request-id'))
            {
                $clientRequestId = $responseHeaders['client-request-id']
            }
            if ($responseHeaders.ContainsKey('x-ms-ags-diagnostic'))
            {
                $diagHeader = $responseHeaders['x-ms-ags-diagnostic']
            }
            if ($responseHeaders.ContainsKey('Date'))
            {
                $serverDate = $responseHeaders['Date']
            }
            if ($responseHeaders.ContainsKey('Retry-After'))
            {
                $retryAfter = $responseHeaders['Retry-After']
            }
            # Body: handle HttpWebResponse stream and PS7 ErrorDetails fallbacks
            try
            {
                if ($resp -is [System.Net.HttpWebResponse])
                {
                    $errorResponse = $resp.GetResponseStream()
                    if ($errorResponse)
                    {
                        $streamReader = New-Object System.IO.StreamReader($errorResponse)
                        $responseBodyRaw = $streamReader.ReadToEnd()
                        $streamReader.Close()
                    }
                }
            }
            catch
            {
                Write-Verbose "[$functionName] Failed to read response stream: $($_.Exception.Message)"
            }
        }
        # Additional fallbacks commonly present in PS7
        if (-not $responseBodyRaw)
        {
            try
            {
                if ($PSItem.ErrorDetails -and $PSItem.ErrorDetails.Message)
                {
                    $responseBodyRaw = $PSItem.ErrorDetails.Message
                }
            }
            catch
            {
                Write-Verbose "[$functionName] Failed to retrieve error details message: $($_.Exception.Message)"
            }
        }
        if (-not $responseBodyRaw)
        {
            try
            {
                if ($PSItem.Exception.Response -and $PSItem.Exception.Response.Content)
                {
                    $responseBodyRaw = [string]$PSItem.Exception.Response.Content
                }
            }
            catch
            {
                Write-Verbose "[$functionName] Failed to retrieve response content: $($_.Exception.Message)"
            }
        }

        # Parse JSON body if it looks like JSON
        if ($responseBodyRaw)
        {
            Write-Verbose "[$functionName] Raw server response captured (truncated for display if large)."
            Write-Verbose "[$functionName] Server Response (raw): $responseBodyRaw"
            try
            {
                $responseJson = $responseBodyRaw | ConvertFrom-Json -ErrorAction Stop
            }
            catch
            {
                $responseJson = $null
            }
        }
        # Extract Graph error fields when available
        if ($null -ne $responseJson -and $responseJson.error)
        {
            $graphError = $responseJson.error
            $graphCode = $graphError.code
            $graphMessage = $graphError.message
            Write-Verbose "[$functionName] Graph error code: $graphCode"
            Write-Verbose "[$functionName] Graph error message: $graphMessage"
            if ($graphError.innerError)
            {
                $innerErr = $graphError.innerError
                # Newer Graph may use camelCase innerError fields; older uses innererror
                try
                {
                    if (-not $requestId -and $innerErr.'request-id')
                    {
                        $requestId = $innerErr.'request-id'
                    }
                }
                catch
                {
                    Write-Verbose "[$functionName] Failed to retrieve inner error request-id: $($_.Exception.Message)"
                }
                try
                {
                    if (-not $clientRequestId -and $innerErr.'client-request-id')
                    {
                        $clientRequestId = $innerErr.'client-request-id'
                    }
                }
                catch
                {
                    Write-Verbose "[$functionName] Failed to retrieve inner error client-request-id: $($_.Exception.Message)"
                }
                try
                {
                    if (-not $serverDate -and $innerErr.date)
                    {
                        $serverDate = $innerErr.date
                    }
                }
                catch
                {
                    Write-Verbose "[$functionName] Failed to retrieve inner error date: $($_.Exception.Message)"
                }
                Write-Verbose "[$functionName] Graph innerError: request-id=$requestId client-request-id=$clientRequestId date=$serverDate"
                # Some APIs include nested innererror with additional code/message
                if ($innerErr.innererror)
                {
                    Write-Verbose "[$functionName] Graph nested innererror: $($innerErr.innererror | ConvertTo-Json -Depth 5)"
                }
            }
        }

        # Summarize headers and identifiers (avoid logging Authorization)
        if ($responseHeaders.Count -gt 0)
        {
            Write-Verbose "[$functionName] Response headers:"
            foreach ($k in $responseHeaders.Keys | Sort-Object)
            {
                if ($k -ne 'Authorization')
                {
                    Write-Verbose "[$functionName]   $($k): $($responseHeaders[$k])"
                }
            }
        }
        if ($requestId)
        {
            Write-Verbose "[$functionName] Request-Id: $requestId"
        }
        if ($clientRequestId)
        {
            Write-Verbose "[$functionName] Client-Request-Id: $clientRequestId"
        }
        if ($diagHeader)
        {
            Write-Verbose "[$functionName] x-ms-ags-diagnostic: $diagHeader"
        }
        if ($serverDate)
        {
            Write-Verbose "[$functionName] Server Date: $serverDate"
        }
        if ($retryAfter)
        {
            Write-Verbose "[$functionName] Retry-After: $retryAfter"
        }
        try
        {
            # Build a consolidated diagnostic message
            $headersText = ''
            if ($responseHeaders.Count -gt 0)
            {
                $headersText = ($responseHeaders.GetEnumerator() | Where-Object { $_.Key -ne 'Authorization' } | Sort-Object Key | ForEach-Object { "${($_.Key)}: ${($_.Value)}" }) -join [Environment]::NewLine
            }
            $graphInnerDump = $null
            if ($responseJson -and $responseJson.error -and $responseJson.error.innerError)
            {
                try
                {
                    $graphInnerDump = ($responseJson.error.innerError | ConvertTo-Json -Depth 8)
                }
                catch
                {
                    $graphInnerDump = ($responseJson.error.innerError | Out-String)
                }
            }
            $rawBodyForLog = $responseBodyRaw
            # Optionally truncate extremely large bodies to keep logs manageable
            $maxBody = 50000
            if ($rawBodyForLog -and $rawBodyForLog.Length -gt $maxBody)
            {
                $rawBodyForLog = $rawBodyForLog.Substring(0, $maxBody) + "... (truncated; total length=$($responseBodyRaw.Length))"
            }
            $logMessage = @"
[$functionName] Graph API call failed.
ExceptionType: $($PSItem.Exception.GetType().FullName)
ExceptionMessage: $($PSItem.Exception.Message)
StatusCode: $statusCode
StatusDescription: $statusDescription
StatusCodeMessage: $statusCodeMessage
Request-Id: $requestId
Client-Request-Id: $clientRequestId
ServerDate: $serverDate
Retry-After: $retryAfter
Headers:
$headersText

GraphErrorCode: $graphCode
GraphErrorMessage: $graphMessage
GraphInnerError:
$graphInnerDump

ResponseBody:
$rawBodyForLog
"@

            Write-Verbose "[$functionName] (fallback) $logMessage"
        }
        catch
        {
            Write-Verbose "[$functionName] Failed to write diagnostics: $($_.Exception.Message)"
        }

        # Preserve existing switch logic for user-friendly messages
        $statusMessage = $statusMessage
        switch ($statusCode)
        {
            400
            {
                Write-Verbose "[$functionName] Bad request. Please check the resource name."
            }
            401
            {
                Write-Verbose "[$functionName] Unauthorized. Please check your access token."
            }
            403
            {
                Write-Verbose "[$functionName] Forbidden. You do not have permission to access this resource."
            }
            404
            {
                Write-Verbose "[$functionName] Not found. The resource does not exist."
            }
            default
            {
                Write-Verbose "[$functionName] An unknown error occurred. Please check the error message below."
                Write-Verbose "[$functionName] Error: $statusMessage"
                if ($statusCode)
                {
                    Write-Verbose "[$functionName] Status code: $statusCode"
                }
                if ($statusDescription)
                {
                    Write-Verbose "[$functionName] Status description: $statusDescription"
                }
                if ($statusCodeMessage)
                {
                    Write-Verbose "[$functionName] $statusCode indicates $statusCodeMessage"
                }
                Write-Verbose "[$functionName] Status message: $statusMessage"
                if ($requestId)
                {
                    Write-Verbose "[$functionName] Request-Id: $requestId"
                }
                if ($clientRequestId)
                {
                    Write-Verbose "[$functionName] Client-Request-Id: $clientRequestId"
                }
                if ($retryAfter)
                {
                    Write-Verbose "[$functionName] Retry-After: $retryAfter"
                }
                Write-Verbose "[$functionName] The full error message follows below:"
                Write-Verbose "[$functionName] ----------------------------------------------------------"
                Write-Verbose "[$functionName] $_"
                # Raw server body already logged above when available
            }
        }
        Write-Verbose "[$functionName] Failed to call the Graph API: $_"
        Write-Verbose "[$functionName] The status code is $statusCode"
        if ($statusCodeMessage)
        {
            Write-Verbose "[$functionName] $statusCode indicates $statusCodeMessage"
        }
        if ($statusDescription)
        {
            Write-Verbose "[$functionName] Status description: $statusDescription"
        }
        Write-Verbose "[$functionName] Status message: $statusMessage"
        Write-Verbose "[$functionName] The full error message follows below:"
        Write-Verbose "[$functionName] ----------------------------------------------------------"
        Write-Verbose "[$functionName] Error: $($_)"
        Write-Verbose "[$functionName] Exception message: $($PSItem.Exception.Message)"
        Write-Verbose "[$functionName] Exception response: $($PSItem.Exception.Response)"
        if ($responseBodyRaw)
        {
            Write-Verbose "[$functionName] Server Response (raw): $responseBodyRaw"
        }
        return $statusCode
        # return $null
    }
    Write-Verbose "[$functionName] Response: $($response)"
    Write-Verbose "[$functionName] Response value: $($response.value)"
    return $response
}

function Get-GraphAccessToken()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$tenantId,
        [Parameter(Mandatory = $true)]
        [string]$clientId,
        [Parameter(Mandatory = $false)]
        [string]$clientSecret,
        [Parameter(Mandatory = $false)]
        [string]$certificateThumbprint
    )

    #region helper functions
    function Get-TokenExpiryTime()
    {
        [CmdletBinding()]
        param(
            [object]$accessTokenObject
        )
        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Retrieving token expiry time"

        # Check if AbsoluteExpiryTime property exists and has a value
        if ($accessTokenObject.PSObject.Properties['AbsoluteExpiryTime'] -and $accessTokenObject.AbsoluteExpiryTime)
        {
            Write-Verbose "[$functionName] Token expires at: $($accessTokenObject.AbsoluteExpiryTime)"
            return $accessTokenObject.AbsoluteExpiryTime
        }

        Write-Verbose "[$functionName] No AbsoluteExpiryTime found in token object"
        return [datetime]::MinValue
    }

    function Save-TokenToCache()
    {
        [CmdletBinding()        ]
        param(
            [object]$tokenResponse
        )
        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Saving access token to memory cache"

        # Initialize global memory cache if it doesn't exist
        if (-not (Get-Variable -Name 'MemoryCache' -Scope Global -ErrorAction SilentlyContinue))
        {
            Write-Verbose "[$functionName] Initializing global memory cache"
            New-Variable -Name 'MemoryCache' -Scope Global -Value @{} -Force
        }

        # Calculate absolute expiry time
        $expiresInSeconds = $tokenResponse.expires_in
        $absoluteExpiryTime = (Get-Date).AddSeconds($expiresInSeconds)

        # Create a proper cache object with all token properties plus expiry time
        $cachedTokenObject = [PSCustomObject]@{
            access_token       = $tokenResponse.access_token
            token_type         = $tokenResponse.token_type
            expires_in         = $tokenResponse.expires_in
            ext_expires_in     = $tokenResponse.ext_expires_in
            AbsoluteExpiryTime = $absoluteExpiryTime
        }

        Write-Verbose "[$functionName] Token expires at: $absoluteExpiryTime (in $expiresInSeconds seconds)"

        $Global:MemoryCache['accessToken'] = $cachedTokenObject
        Write-Verbose "[$functionName] Token successfully saved to memory cache"
    }

    function New-JwtClientAssertion()
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
            [Parameter(Mandatory = $true)]
            [string]$ClientId,
            [Parameter(Mandatory = $true)]
            [string]$TokenEndpoint
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Creating JWT client assertion"

        $now = [Math]::Floor([decimal](Get-Date (Get-Date).ToUniversalTime() -UFormat "%s"))
        $exp = $now + 600 # Token valid for 10 minutes

        # Create JWT header
        $jwtHeader = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [Convert]::ToBase64String($Certificate.GetCertHash()) -replace '\+', '-' -replace '/', '_' -replace '='
        } | ConvertTo-Json -Compress

        # Create JWT payload
        $jwtPayload = @{
            aud = $TokenEndpoint
            iss = $ClientId
            sub = $ClientId
            jti = [guid]::NewGuid().ToString()
            nbf = $now
            exp = $exp
        } | ConvertTo-Json -Compress

        # Base64Url encode header and payload
        $jwtHeaderEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)) -replace '\+', '-' -replace '/', '_' -replace '='
        $jwtPayloadEncoded = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtPayload)) -replace '\+', '-' -replace '/', '_' -replace '='

        # Create signature
        $jwtToSign = "$jwtHeaderEncoded.$jwtPayloadEncoded"
        $jwtBytes = [System.Text.Encoding]::UTF8.GetBytes($jwtToSign)

        Write-Verbose "[$functionName] Signing JWT with certificate private key"
        $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)

        if (-not $privateKey)
        {
            throw "Failed to access certificate private key"
        }

        $signature = $privateKey.SignData($jwtBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $jwtSignatureEncoded = [Convert]::ToBase64String($signature) -replace '\+', '-' -replace '/', '_' -replace '='

        Write-Verbose "[$functionName] JWT client assertion created successfully"
        return "$jwtToSign.$jwtSignatureEncoded"
    }

    function Invoke-CertificateAuthentication()
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$CertificateThumbprint,
            [Parameter(Mandatory = $true)]
            [string]$ClientId,
            [Parameter(Mandatory = $true)]
            [string]$TokenEndpoint,
            [Parameter(Mandatory = $true)]
            [string]$Scope
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Attempting certificate-based authentication"

        # Get certificate from store
        $certificate = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -Recurse |
            Where-Object { $_.Thumbprint -eq $CertificateThumbprint } |
            Select-Object -First 1

        if (-not $certificate)
        {
            throw "Certificate with thumbprint $CertificateThumbprint not found in certificate stores"
        }

        Write-Verbose "[$functionName] Certificate found: Subject=$($certificate.Subject), NotAfter=$($certificate.NotAfter)"

        # Validate certificate
        if ($certificate.NotAfter -lt (Get-Date))
        {
            throw "Certificate has expired on $($certificate.NotAfter)"
        }

        if (-not $certificate.HasPrivateKey)
        {
            throw "Certificate does not have a private key"
        }

        # Create JWT client assertion
        $clientAssertion = New-JwtClientAssertion -Certificate $certificate -ClientId $ClientId -TokenEndpoint $TokenEndpoint

        # Build request body
        $body = @{
            client_id             = $ClientId
            scope                 = $Scope
            client_assertion      = $clientAssertion
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            grant_type            = 'client_credentials'
        }

        Write-Verbose "[$functionName] Sending token request with certificate authentication"

        $tokenResponse = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop

        Write-Verbose "[$functionName] Access token received successfully via certificate authentication"
        Write-Verbose "[$functionName] Token expires in: $($tokenResponse.expires_in) seconds"

        return $tokenResponse
    }

    function Invoke-ClientSecretAuthentication()
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$ClientId,
            [Parameter(Mandatory = $true)]
            [string]$ClientSecret,
            [Parameter(Mandatory = $true)]
            [string]$TokenEndpoint,
            [Parameter(Mandatory = $true)]
            [string]$Scope
        )

        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Attempting client secret authentication"

        $body = @{
            client_id     = $ClientId
            scope         = $Scope
            client_secret = $ClientSecret
            grant_type    = 'client_credentials'
        }

        Write-Verbose "[$functionName] Sending request to token endpoint"

        $tokenResponse = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop

        Write-Verbose "[$functionName] Access token received successfully via client secret authentication"
        Write-Verbose "[$functionName] Token expires in: $($tokenResponse.expires_in) seconds"

        return $tokenResponse
    }

    function Write-AuthenticationError()
    {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$FunctionName,
            [Parameter(Mandatory = $true)]
            [string]$AuthMethod,
            [Parameter(Mandatory = $true)]
            $Exception
        )

        Write-Verbose "[$FunctionName] $AuthMethod authentication failed: $Exception"

        if ($Exception.Exception.Response)
        {
            $statusCode = $Exception.Exception.Response.StatusCode
            Write-Verbose "[$FunctionName] Status code: $statusCode"

            try
            {
                $errorResponse = $Exception.Exception.Response.GetResponseStream()
                $streamReader = New-Object System.IO.StreamReader($errorResponse)
                $errorMessage = $streamReader.ReadToEnd()
                $streamReader.Close()

                $errorJson = $errorMessage | ConvertFrom-Json
                Write-Verbose "[$FunctionName] Error code: $($errorJson.error)"
                Write-Verbose "[$FunctionName] Error description: $($errorJson.error_description)"
            }
            catch
            {
                Write-Verbose "[$FunctionName] Could not parse error response"
            }
        }
    }
    #endregion helper functions

    $functionName = $MyInvocation.MyCommand.Name
    $renewalLeadTimeInMinutes = 5
    $timeBuffer = (Get-Date).AddMinutes($renewalLeadTimeInMinutes)
    Write-Verbose "[$functionName] Token renewal buffer time: $timeBuffer"
    if (-not (Get-Variable -Name 'MemoryCache' -Scope Global -ErrorAction SilentlyContinue))
    {
        Write-Verbose "[$functionName] Initializing global memory cache"
        New-Variable -Name 'MemoryCache' -Scope Global -Value @{} -Force
    }

    if ($Global:MemoryCache.ContainsKey('accessToken'))
    {
        Write-Verbose "[$functionName] Found token in memory cache"
        $tokenObject = $Global:MemoryCache['accessToken']
        $absoluteExpiryTime = Get-TokenExpiryTime -accessTokenObject $tokenObject
        Write-Verbose "[$functionName] Token expiry time: $absoluteExpiryTime"
        if ($absoluteExpiryTime -gt $timeBuffer)
        {
            Write-Verbose "[$functionName] Access token is valid until $absoluteExpiryTime"
            Write-Host "Valid Token retrieved from cache." -ForegroundColor Green
            [console]::beep(200, 200)
            return $tokenObject.access_token
        }
        else
        {
            Write-Verbose "[$functionName] Cached token has expired or expires soon, acquiring new token"
        }
    }
    else
    {
        Write-Verbose "[$functionName] No token found in memory cache"
    }

    # Validate that we have at least one authentication method
    if (-not $clientSecret -and -not $certificateThumbprint)
    {
        $errorMsg = "Either clientSecret or certificateThumbprint must be provided"
        Write-Verbose "[$functionName] $errorMsg"
        return $null
    }

    # Determine authentication strategy
    $useCertificate = $false
    $useClientSecret = $false
    $tryBothWithFallback = $false

    # Check for non-empty certificate and client secret
    $hasCertificate = -not [string]::IsNullOrWhiteSpace($certificateThumbprint)
    $hasClientSecret = -not [string]::IsNullOrWhiteSpace($clientSecret)

    if ($hasCertificate -and $hasClientSecret)
    {
        Write-Verbose "[$functionName] Both certificate and client secret provided - will try certificate first with fallback to secret"
        $tryBothWithFallback = $true
        $useCertificate = $true
    }
    elseif ($hasCertificate)
    {
        Write-Verbose "[$functionName] Using certificate-based authentication (certificate-only mode)"
        $useCertificate = $true
    }
    else
    {
        Write-Verbose "[$functionName] Using client secret authentication"
        $useClientSecret = $true
    }

    $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $scope = 'https://graph.microsoft.com/.default'
    Write-Verbose "[$functionName] Token endpoint: $tokenEndpoint"

    # Attempt certificate authentication
    if ($useCertificate)
    {
        try
        {
            $tokenResponse = Invoke-CertificateAuthentication -CertificateThumbprint $certificateThumbprint -ClientId $clientId -TokenEndpoint $tokenEndpoint -Scope $scope
            Save-TokenToCache -tokenResponse $tokenResponse
            return $tokenResponse.access_token
        }
        catch
        {
            Write-AuthenticationError -FunctionName $functionName -AuthMethod "Certificate" -Exception $_

            if ($tryBothWithFallback)
            {
                Write-Warning "[$functionName] Certificate authentication failed - falling back to client secret"
                $useCertificate = $false
                $useClientSecret = $true
            }
            else
            {
                return $null
            }
        }
    }

    # Attempt client secret authentication
    if ($useClientSecret)
    {
        try
        {
            $tokenResponse = Invoke-ClientSecretAuthentication -ClientId $clientId -ClientSecret $clientSecret -TokenEndpoint $tokenEndpoint -Scope $scope
            Save-TokenToCache -tokenResponse $tokenResponse
            return $tokenResponse.access_token
        }
        catch
        {
            Write-AuthenticationError -FunctionName $functionName -AuthMethod "Client secret" -Exception $_
            return $null
        }
    }
}
#endregion helper functions

$configPath = Join-Path -Path $PSScriptRoot -ChildPath $configFile
if (-not (Test-Path -Path $configPath))
{
    Write-Host "Config file not found at $configPath. Using default values." -ForegroundColor Red
    $config = @{}
}
else
{
    Write-Verbose "Config file found at $configPath. Loading configuration."
    $config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
}

#region Define variables
$tenantId = if ($config.tenantId)
{
    $config.tenantId
}
else
{
    $null
}
$clientId = if ($config.appId)
{
    $config.appId
}
else
{
    $null
}
$clientSecret = if ($config.AppSecret)
{
    $config.AppSecret
}
else
{
    $null
}
$certificateThumbprint = if ($config.thumbprint)
{
    $config.thumbprint
}
else
{
    $null
}
$managedAppUri = "deviceAppManagement/mobileApps"
$accessToken = Get-GraphAccessToken -tenantId $tenantId -clientId $clientId -clientSecret $clientSecret -certificateThumbprint $certificateThumbprint
$appTypes = @(
    @{
        AppType     = "Win32 App"
        ODataType   = "#microsoft.graph.win32LobApp"
        Description = "Traditional `.intunewin` wrapped apps"
    },
    @{
        AppType     = "Windows universal app"
        ODataType   = "#microsoft.graph.windowsUniversalAppX"
        Description = "Universal Windows Platform (UWP) apps"
    },
    @{
        AppType     = "MSI (LOB) app"
        ODataType   = "#microsoft.graph.windowsMobileMSI"
        Description = "Line-of-Business `.msi` files (Non-Win32 wrapped)"
    },
    @{
        AppType     = "WinGet App"
        ODataType   = "#microsoft.graph.winGetApp"
        Description = "Apps deployed via the Windows Package Manager"
    },
    @{
        AppType     = "MSIX / AppX"
        ODataType   = "#microsoft.graph.windowsAppX"
        Description = "Modern Windows package files (`.msix`, `.appx`, `.msixbundle`)"
    },
    @{
        AppType     = "Office Suite App"
        ODataType   = "#microsoft.graph.officeSuiteApp"
        Description = "Microsoft 365 Apps for enterprise (formerly Office 365 ProPlus) deployment configurations"
    },
    @{
        AppType     = "Web app"
        ODataType   = "#microsoft.graph.webApp"
        Description = "Web applications accessed via a URL and optionally integrated with Azure AD for single sign-on"
    },
    @{
        AppType     = "Windows store app"
        ODataType   = "#microsoft.graph.windowsStoreApp"
        Description = "Apps available in the Microsoft Store for Business and Education that can be deployed to devices"
    },
    @{
        AppType     = "Store App (New)"
        ODataType   = "#microsoft.graph.microsoftStoreForBusinessApp"
        Description = "Apps from the Microsoft Store (including the `"new`" Store experience)"
    }
)
$appArray = @()
#endregion Define variables

$apps = Invoke-GraphAPI -ResourcePath $managedAppUri -accessToken $AccessToken
foreach ($app in $apps.value)
{
    if ($app.'@odata.type' -in $appTypes.ODataType)
    {
        $matchedType = $appTypes | Where-Object { $_.'ODataType' -eq $app.'@odata.type' }
        $appObject = [PSCustomObject]@{
            displayName = $app.displayName
            type        = $matchedType.AppType
            description = $matchedType.Description
            id          = $app.id
        }
        $appArray += $appObject
    }
}
Write-Host "Processed $($appArray.Count) apps..." -ForegroundColor Green

# Use graphical interface to select app
if ($appArray.Count -eq 0)
{
    Write-Host "No eligible apps found. Exiting script." -ForegroundColor Red
    exit 1
}

$selectedApp = $appArray | Out-GridView -Title "Select an application to update" -OutputMode Single

if ($null -eq $selectedApp)
{
    Write-Host "No app selected. Exiting script." -ForegroundColor Yellow
    exit 0
}

Write-Host "Selected app: '$($selectedApp.displayName)'" -ForegroundColor Yellow

# Use File Open Dialog to select PNG file
Add-Type -AssemblyName System.Windows.Forms
$fileDialog = New-Object System.Windows.Forms.OpenFileDialog
$fileDialog.Title = "Select PNG Image File"
$fileDialog.Filter = "PNG Files (*.png)|*.png|All Files (*.*)|*.*"
$fileDialog.InitialDirectory = [Environment]::GetFolderPath('MyPictures')

$dialogResult = $fileDialog.ShowDialog()

if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK)
{
    Write-Host "No file selected. Exiting script." -ForegroundColor Yellow
    exit 0
}

$path = $fileDialog.FileName
Write-Host "Selected file: $path" -ForegroundColor Cyan
$imageBytes = [System.IO.File]::ReadAllBytes($path)
$base64Image = [Convert]::ToBase64String($imageBytes)

# Construct request body as hashtable, then convert to JSON
$params = @{
    largeIcon = @{
        "@odata.type" = "microsoft.graph.mimeContent"
        type          = "image/png"
        value         = $base64Image
    }
}

# Convert to JSON for API call
$bodyJson = $params | ConvertTo-Json -Depth 10

# Make API call with JSON body
Write-Host "Updating app icon..." -ForegroundColor Cyan
$APIResponse = Invoke-GraphAPI -ResourcePath "$managedAppUri/$($selectedApp.id)" -accessToken $AccessToken -Method PATCH -body $bodyJson

# Check for successful response
# Invoke-GraphAPI returns PSCustomObject on success, integer status code on error
if ($APIResponse -is [PSCustomObject])
{
    Write-Host "Successfully updated the app's large icon." -ForegroundColor Green
}
elseif ($APIResponse -is [int])
{
    Write-Host "Failed to update the app's large icon. Status code: $APIResponse" -ForegroundColor Red
}
else
{
    Write-Host "Failed to update the app's large icon. Unexpected response type." -ForegroundColor Red
}
