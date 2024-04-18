function Get-RequiredModules {
    param (
        [Parameter(Mandatory=$true)]
        [string]$moduleName
    )

    # Check if the module is installed
    $moduleInstalled = Get-Module -ListAvailable -Name $moduleName
    if (-not $moduleInstalled) 
    {
        Write-Host "The required module '$moduleName' is not installed. Trying to install it." -ForegroundColor Yellow

        try {
            Install-Module -Name $moduleName -Force -Scope CurrentUser
        } 
        
        catch {
            Write-Error "Could not install module '$moduleName' due to error: $_"
            return
        }
    }

    # Check if the module is imported
    $moduleImported = Get-Module -Name $moduleName
    if (-not $moduleImported) 
    {
        Write-Host "The required module '$moduleName' is not imported. Trying to import it." -ForegroundColor Yellow

        try {
            Import-Module -Name $moduleName
        } 
        
        catch {
            Write-Error "Could not import module '$moduleName' due to error: $_"
        }
    }
}

function Get-AccessTokenMSAL-ApplicationPermission {
    param (
        [Parameter(Mandatory=$true)]
        [string]$clientId,
        [Parameter(Mandatory=$true)]
        [System.Security.SecureString]$clientSecret,
        [Parameter(Mandatory=$true)]
        [string]$tenantId
    )


    #Checking if required Modules are installed
    $RequiredModulesInstalled = Get-RequiredModules -moduleName "MSAL.PS"

    $scope = "https://graph.microsoft.com/.default"
    $tokenResult = Get-MsalToken -ClientId $clientId -ClientSecret $clientSecret -TenantId $tenantId -Scopes $scope
    
    return $tokenResult.AccessToken
    
}

function Get-AccessTokenMSAL-DelegatedPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$clientId,
        [Parameter(Mandatory=$true)]
        [string]$tenantId
    )

    $RequiredModulesInstalled = Get-RequiredModules -moduleName "MSAL.PS"
    $scope = "https://graph.microsoft.com/.default"
    $tokenResult = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes $scope -RedirectUri "http://localhost"

    return $tokenResult
}

function New-LocalSecret {
    param (
        [Parameter(Mandatory=$true)] 
        [string]$clientId,
        [ValidateSet("User", "System-Wide")]
        [string]$scope = "System-Wide"
    )

    $RequiredModulesInstalled = Get-RequiredModules -moduleName "CredentialManager"
    $credential = Get-Credential -UserName $clientId -Message "Enter the clientSecret for the App Registration."

    # Check if a non-empty password was provided
    if (-not [string]::IsNullOrEmpty($credential.GetNetworkCredential().Password)) {
        # Get the password (clientSecret)
        $clientSecret = $credential.GetNetworkCredential().Password

        # Add a new generic credential to Windows Credential Store
        $credentialName = "AzureAppRegistration_$clientId"
        $userName = $clientId

        # Determine the persistence type based on the scope parameter
        $persistenceType = $null
        switch ($scope) {
            "User" { $persistenceType = "LocalMachine" }
            "System-Wide" { $persistenceType = "Enterprise" }
        }

        
        New-StoredCredential -Target $credentialName -UserName $userName -Password $clientSecret -Type Generic -Persist LocalMachine | Out-Null
        
        return "Credential saved successfully to Windows Credential Store."
    }
    else {
        Write-Host "ERROR: There was no clientSecret provided! Run the function again and provide a valid clientSecret!" -ForegroundColor Red
    }
}

function Get-LocalSecret {
    param (
        [string]$clientId
    )

    $RequiredModulesInstalled = Get-RequiredModules -moduleName "CredentialManager"

    try {
        $credentialName = "AzureAppRegistration_$clientId"
        $storedCredential = Get-StoredCredential -Target $credentialName

        # Check if a credential was returned
        if ($storedCredential -ne $null) {
            return $storedCredential.Password
        } else {
            Write-Host "ERROR: No credential found for the given clientId. Check if the clientId is correct or if the credential exists in the Windows Credential Store." -ForegroundColor Red
        }
    }
    catch {
        Write-Host "ERROR: Could not retrieve locally stored clientSecret" -ForegroundColor Red
    }
}

function Read-Calendar {
    param (
        [Parameter(Mandatory=$true)]
        [string]$accessToken,
        [Parameter(Mandatory=$true)]
        [string]$fromUser,    # User's email
        [DateTime]$startDate = ($startDate = Get-Date -Day 30 -Hour 0 -Minute 0 -Second 0),
        [DateTime]$endDate = ($startDate.AddMonths(1).AddSeconds(-1))
    )

    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/users/$fromUser/calendarView?startDateTime=$($startDate.ToString('yyyy-MM-ddTHH:mm:ss'))&endDateTime=$($endDate.ToString('yyyy-MM-ddTHH:mm:ss'))&`$filter=showAs eq 'Oof'&`$select=subject,start,end,showAs"

    $headers = @{
        Authorization = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }

    $response = Invoke-RestMethod -Uri $graphApiEndpoint -Method Get -Headers $headers
    return $response.value
}

function Send-Email {
    param (
        [Parameter(Mandatory=$true)]
        [string]$accessToken,
        [Parameter(Mandatory=$true)]
        [string]$recipientEmail,
        [Parameter(Mandatory=$true)]
        [string]$subject,
        [Parameter(Mandatory=$true)]
        [string]$body,
        [Parameter(Mandatory=$true)]
        [string]$fromUserIdOrUpn
    )

    $graphApiEndpoint = "https://graph.microsoft.com/v1.0/users/$fromUserIdOrUpn/sendMail"
    $headers = @{
        Authorization = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }

    $emailData = @{
        message = @{
            subject = $subject
            body = @{
                contentType = "Text"
                content = $body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $recipientEmail
                    }
                }
            )
            from = @{
                emailAddress = @{
                    address = $fromUserIdOrUpn
                }
            }
        }
    }

    $emailJson = $emailData | ConvertTo-Json -Depth 100
    Invoke-RestMethod -Uri $graphApiEndpoint -Method Post -Headers $headers -Body $emailJson -ContentType "application/json"
}


