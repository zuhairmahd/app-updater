function Get-ClientCredentialsToken()
{
    <#
    .SYNOPSIS
    Acquires Microsoft Graph API access token using client credentials (application) authentication flow.

    .DESCRIPTION
    This function implements the OAuth 2.0 client credentials grant for non-delegated (application-only)
    authentication to Microsoft Graph. It supports both client secret and certificate-based authentication,
    with automatic fallback between methods. The function handles token caching, validation, and renewal
    to minimize redundant authentication requests.

    .PARAMETER tenantId
    The Azure AD tenant ID. This parameter is mandatory.

    .PARAMETER clientId
    The application (client) ID from Azure AD app registration. This parameter is mandatory.

    .PARAMETER clientSecret
    The client secret from Azure AD app registration. Either clientSecret or certificateThumbprint required.

    .PARAMETER certificateThumbprint
    The certificate thumbprint for certificate-based authentication. Either clientSecret or certificateThumbprint required.

    .PARAMETER domain
    The domain name for tenant-specific caching and logging.

    .PARAMETER cacheType
    The cache storage type: 'file' or 'memory'.

    .PARAMETER cacheTokenFile
    Path to the cache token file for file-based caching.

    .PARAMETER cacheFolder
    Path to the cache folder for file-based caching.

    .PARAMETER secureString
    When specified, returns access token as SecureString instead of plain text.

    .OUTPUTS
    System.String or System.Security.SecureString
    Returns the access token (plain text or SecureString based on -secureString switch), or $null on error.

    .EXAMPLE
    $token = Get-ClientCredentialsToken -tenantId $tid -clientId $cid -clientSecret $secret
    $token = Get-ClientCredentialsToken -tenantId $tid -clientId $cid -certificateThumbprint $thumb -cacheType 'memory'

    .NOTES
    Uses client credentials OAuth 2.0 flow for application-only permissions.
    Supports both client secret and certificate authentication.
    Implements automatic fallback: tries certificate first if both provided.
    Caches tokens to avoid redundant authentication.
    Validates cached tokens before reuse.
    Compatible with PowerShell 5.1.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$tenantId,
        [Parameter(Mandatory = $true)]
        [string]$clientId,
        [Parameter(Mandatory = $false)]
        [string]$clientSecret,
        [Parameter(Mandatory = $false)]
        [string]$certificateThumbprint,
        [Parameter(Mandatory = $false)]
        [string]$domain,
        [Parameter(Mandatory = $false)]
        [string]$cacheType,
        [Parameter(Mandatory = $false)]
        [string]$cacheTokenFile,
        [Parameter(Mandatory = $false)]
        [string]$cacheFolder,
        [Parameter(Mandatory = $false)]
        [switch]$secureString
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting Get-ClientCredentialsToken function"
    Write-Log -LogFile $LogFile -Module $functionName -Message "Starting Get-ClientCredentialsToken - tenantId=$tenantId, clientId=$clientId, domain=$domain, cacheType=$cacheType"
    Write-Verbose "[$functionName] Using non-delegated access (client credentials flow)"
    Write-Log -LogFile $LogFile -Module $functionName -Message "Using non-delegated access (client credentials flow)"

    # Validate that we have at least one authentication method
    if (-not $clientSecret -and -not $certificateThumbprint)
    {
        $errorMsg = "Either clientSecret or certificateThumbprint must be provided"
        Write-Error "[$functionName] $errorMsg"
        Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error
        return $null
    }

    # Determine authentication strategy
    $useCertificate = $false
    $useClientSecret = $false
    $tryBothWithFallback = $false

    # Check for non-empty certificate and client secret (not just presence)
    $hasCertificate = -not [string]::IsNullOrWhiteSpace($certificateThumbprint)
    $hasClientSecret = -not [string]::IsNullOrWhiteSpace($clientSecret)

    if ($hasCertificate -and $hasClientSecret)
    {
        Write-Verbose "[$functionName] Both certificate and client secret provided - will try certificate first with fallback to secret"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Both certificate and client secret provided - will try certificate first with fallback to secret" -LogLevel Warning
        $tryBothWithFallback = $true
        $useCertificate = $true
    }
    elseif ($hasCertificate)
    {
        Write-Verbose "[$functionName] Using certificate-based authentication (certificate-only mode)"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Using certificate-based authentication (certificate-only mode) with thumbprint: $certificateThumbprint"
        $useCertificate = $true
    }
    else
    {
        Write-Verbose "[$functionName] Using client secret authentication"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Using client secret authentication"
        $useClientSecret = $true
    }

    $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $scope = 'https://graph.microsoft.com/.default'
    Write-Verbose "[$functionName] Token endpoint: $tokenEndpoint"
    Write-Log -LogFile $LogFile -Module $functionName -Message "Token endpoint: $tokenEndpoint"
    Write-Verbose "[$functionName] Scope: $scope"
    Write-Log -LogFile $LogFile -Module $functionName -Message "Scope: $scope"

    # Certificate authentication attempt
    if ($useCertificate)
    {
        Write-Verbose "[$functionName] Attempting certificate-based authentication"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Attempting certificate-based authentication"

        try
        {
            # Get certificate from store
            Write-Verbose "[$functionName] Retrieving certificate with thumbprint: $certificateThumbprint"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Retrieving certificate from store with thumbprint: $certificateThumbprint"

            $certificate = Get-ChildItem -Path Cert:\CurrentUser\My, Cert:\LocalMachine\My -Recurse |
                Where-Object { $_.Thumbprint -eq $certificateThumbprint } |
                Select-Object -First 1

            if (-not $certificate)
            {
                $errorMsg = "Certificate with thumbprint $certificateThumbprint not found in certificate stores"
                Write-Verbose "[$functionName] $errorMsg"
                Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                if ($tryBothWithFallback)
                {
                    Write-Verbose "[$functionName] Certificate authentication failed - falling back to client secret"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate authentication failed - falling back to client secret" -LogLevel Warning
                    $useCertificate = $false
                    $useClientSecret = $true
                }
                else
                {
                    return $null
                }
            }

            if ($certificate)
            {
                Write-Verbose "[$functionName] Certificate found: Subject=$($certificate.Subject), NotAfter=$($certificate.NotAfter)"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate found: Subject=$($certificate.Subject), NotAfter=$($certificate.NotAfter)"

                # Check if certificate has expired
                if ($certificate.NotAfter -lt (Get-Date))
                {
                    $errorMsg = "Certificate has expired on $($certificate.NotAfter)"
                    Write-Verbose "[$functionName] $errorMsg"
                    Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                    if ($tryBothWithFallback)
                    {
                        Write-Verbose "[$functionName] Certificate expired - falling back to client secret"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate expired - falling back to client secret" -LogLevel Warning
                        $useCertificate = $false
                        $useClientSecret = $true
                        $certificate = $null
                    }
                    else
                    {
                        return $null
                    }
                }

                # Check if certificate has a private key
                if ($certificate -and -not $certificate.HasPrivateKey)
                {
                    $errorMsg = "Certificate does not have a private key"
                    Write-Verbose "[$functionName] $errorMsg"
                    Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                    if ($tryBothWithFallback)
                    {
                        Write-Verbose "[$functionName] Certificate has no private key - falling back to client secret"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate has no private key - falling back to client secret" -LogLevel Warning
                        $useCertificate = $false
                        $useClientSecret = $true
                        $certificate = $null
                    }
                    else
                    {
                        return $null
                    }
                }
            }

            if ($certificate)
            {
                # Create JWT client assertion
                Write-Verbose "[$functionName] Creating JWT client assertion"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Creating JWT client assertion"

                $now = [Math]::Floor([decimal](Get-Date (Get-Date).ToUniversalTime() -UFormat "%s"))
                $exp = $now + 600 # Token valid for 10 minutes

                # Create JWT header
                $jwtHeader = @{
                    alg = "RS256"
                    typ = "JWT"
                    x5t = [Convert]::ToBase64String($certificate.GetCertHash()) -replace '\+', '-' -replace '/', '_' -replace '='
                } | ConvertTo-Json -Compress

                # Create JWT payload
                $jwtPayload = @{
                    aud = $tokenEndpoint
                    iss = $clientId
                    sub = $clientId
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
                Write-Log -LogFile $LogFile -Module $functionName -Message "Signing JWT with certificate private key"
                $privateKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
                if (-not $privateKey)
                {
                    $errorMsg = "Failed to access certificate private key"
                    Write-Warning "[$functionName] $errorMsg"
                    Write-Log -LogFile $LogFile -Module $functionName -Message $errorMsg -LogLevel Error

                    if ($tryBothWithFallback)
                    {
                        Write-Warning "[$functionName] Cannot access private key - falling back to client secret"
                        Write-Log -LogFile $LogFile -Module $functionName -Message "Cannot access private key - falling back to client secret" -LogLevel Warning
                        $useCertificate = $false
                        $useClientSecret = $true
                        $certificate = $null
                        # Skip to client secret authentication
                        throw "PrivateKeyAccessFailed"
                    }
                    else
                    {
                        return $null
                    }
                }

                if ($privateKey -is [System.Security.Cryptography.RSACryptoServiceProvider])
                {
                    $signature = $privateKey.SignData($jwtBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                }
                else
                {
                    # For CNG keys
                    $signature = $privateKey.SignData($jwtBytes, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                }

                $jwtSignatureEncoded = [Convert]::ToBase64String($signature) -replace '\+', '-' -replace '/', '_' -replace '='
                $clientAssertion = "$jwtToSign.$jwtSignatureEncoded"

                Write-Verbose "[$functionName] JWT client assertion created successfully"
                Write-Log -LogFile $LogFile -Module $functionName -Message "JWT client assertion created successfully"

                # Build request body with certificate authentication
                $body = @{
                    client_id             = $clientId
                    scope                 = $scope
                    client_assertion      = $clientAssertion
                    client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                    grant_type            = 'client_credentials'
                }

                Write-Verbose "[$functionName] Sending token request with certificate authentication"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Sending token request with certificate authentication"

                $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop

                Write-Verbose "[$functionName] Access token received successfully via certificate authentication"
                Write-Verbose "[$functionName] Token expires in: $($tokenResponse.expires_in) seconds"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Access token received successfully via certificate authentication. Expires in: $($tokenResponse.expires_in) seconds"

                $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain
                Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
                Write-Log -LogFile $LogFile -Module $functionName -Message "Token cached successfully (type: $cacheType)"

                return Format-TokenOutput -token $tokenResponse.access_token -secureString $secureString
            }
        }
        catch
        {
            Write-Error "[$functionName] Certificate authentication failed: $_"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate authentication failed: $_" -LogLevel Error
            if ($_.Exception.Response)
            {
                Write-Verbose "[$functionName] Status code: $($_.Exception.Response.StatusCode)"
                Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP Status code: $($_.Exception.Response.StatusCode)" -LogLevel Error
                try
                {
                    $errorResponse = $_.Exception.Response.GetResponseStream()
                    $streamReader = New-Object System.IO.StreamReader($errorResponse)
                    $errorMessage = $streamReader.ReadToEnd()
                    $streamReader.Close()
                    Write-Verbose "[$functionName] Server Response: $errorMessage"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Server Response: $errorMessage" -LogLevel Error

                    $errorJson = $errorMessage | ConvertFrom-Json
                    Write-Verbose "[$functionName] Error code: $($errorJson.error)"
                    Write-Verbose "[$functionName] Error description: $($errorJson.error_description)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Error code: $($errorJson.error), Description: $($errorJson.error_description)" -LogLevel Error
                }
                catch
                {
                    Write-Verbose "[$functionName] Could not parse error response: $_"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Could not parse error response" -LogLevel Error
                }
            }

            if ($tryBothWithFallback)
            {
                Write-Warning "[$functionName] Certificate authentication failed - falling back to client secret"
                Write-Log -LogFile $LogFile -Module $functionName -Message "Certificate authentication failed - falling back to client secret" -LogLevel Warning
                $useCertificate = $false
                $useClientSecret = $true
            }
            else
            {
                return $null
            }
        }
    }

    # Client secret authentication attempt
    if ($useClientSecret)
    {
        Write-Verbose "[$functionName] Attempting client secret authentication"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Attempting client secret authentication"

        $body = @{
            client_id     = $clientId
            scope         = $scope
            client_secret = $clientSecret
            grant_type    = 'client_credentials'
        }

        Write-Verbose "[$functionName] Token request body: client_id=$clientId, scope=$scope, grant_type=client_credentials"
        Write-Log -LogFile $LogFile -Module $functionName -Message "Token request body prepared: client_id=$clientId, scope=$scope, grant_type=client_credentials"

        try
        {
            Write-Verbose "[$functionName] Sending request to token endpoint: $tokenEndpoint"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Sending token request with client secret authentication"

            $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -ContentType 'application/x-www-form-urlencoded' -Body $body -ErrorAction Stop

            Write-Verbose "[$functionName] Access token received successfully via client secret authentication"
            Write-Verbose "[$functionName] Token expires in: $($tokenResponse.expires_in) seconds"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Access token received successfully via client secret authentication. Expires in: $($tokenResponse.expires_in) seconds"

            $cachedToken = Get-TokenFromResponse -tokenResponse $tokenResponse -domain $domain
            Save-TokenToCache -cachedToken $cachedToken -cacheType $cacheType -cacheTokenFile $cacheTokenFile -cacheFolder $cacheFolder
            Write-Log -LogFile $LogFile -Module $functionName -Message "Token cached successfully (type: $cacheType)"

            return Format-TokenOutput -token $tokenResponse.access_token -secureString $secureString
        }
        catch
        {
            Write-Error "[$functionName] Failed to get access token via client secret: $_"
            Write-Log -LogFile $LogFile -Module $functionName -Message "Failed to get access token via client secret: $_" -LogLevel Error

            if ($_.Exception.Response)
            {
                Write-Verbose "[$functionName] Status code: $($_.Exception.Response.StatusCode)"
                Write-Log -LogFile $LogFile -Module $functionName -Message "HTTP Status code: $($_.Exception.Response.StatusCode)" -LogLevel Error

                try
                {
                    $errorResponse = $_.Exception.Response.GetResponseStream()
                    $streamReader = New-Object System.IO.StreamReader($errorResponse)
                    $errorMessage = $streamReader.ReadToEnd()
                    $streamReader.Close()
                    Write-Verbose "[$functionName] Server Response: $errorMessage"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Server Response: $errorMessage" -LogLevel Error

                    $errorJson = $errorMessage | ConvertFrom-Json
                    Write-Verbose "[$functionName] Error code: $($errorJson.error)"
                    Write-Verbose "[$functionName] Error description: $($errorJson.error_description)"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Error code: $($errorJson.error), Description: $($errorJson.error_description)" -LogLevel Error
                }
                catch
                {
                    Write-Verbose "[$functionName] Could not parse error response: $_"
                    Write-Log -LogFile $LogFile -Module $functionName -Message "Could not parse error response" -LogLevel Error
                }
            }

            return $null
        }
    }
}

