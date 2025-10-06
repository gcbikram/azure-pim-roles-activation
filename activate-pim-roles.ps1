# Function to initialize connections to Microsoft Graph and Azure
function Initialize-Connections {
    # Check if already connected to Microsoft Graph
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
        Connect-MgGraph -NoWelcome
        $context = Get-MgContext
    }
    
    if (-not $context) {
        throw "Failed to establish Microsoft Graph connection."
    }
    
    Write-Host "✓ Connected to Microsoft Graph - Tenant: $($context.TenantId)" -ForegroundColor Green
    Write-Host "✓ Account: $($context.Account)" -ForegroundColor Green

    # Check Azure PowerShell connection
    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $azContext) {
        Write-Host "Connecting to Azure PowerShell..." -ForegroundColor Cyan
        Connect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        $azContext = Get-AzContext -ErrorAction SilentlyContinue
    }
    
    if ($azContext) {
        Write-Host "✓ Connected to Azure - Subscription: $($azContext.Subscription.Name)" -ForegroundColor Green
    } else {
        Write-Host "Note: Azure PowerShell not connected. Will only check Entra ID roles." -ForegroundColor Yellow
    }

    return @{
        MgContext = $context
        AzContext = $azContext
    }
}

# Function to get Entra ID eligible roles
function Get-EntraIDEligibleRoles {
    param([string]$CurrentUser)
    
    Write-Host "Retrieving eligible Entra ID PIM roles..." -ForegroundColor Cyan
    $entraRoles = @()
    
    try {
        $myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$CurrentUser'"
        
        foreach ($role in $myRoles) {
            $scopeDisplayName = "/"
            $scopeType = "Directory"
            
            if ($role.DirectoryScopeId -ne "/" -and $role.DirectoryScopeId) {
                $scopeDisplayName = $role.DirectoryScopeId
                $scopeType = "Custom"
            }
            
            $entraRoles += [PSCustomObject]@{
                DisplayName = $role.RoleDefinition.DisplayName
                RoleDefinitionId = $role.RoleDefinitionId
                PrincipalId = $role.PrincipalId
                DirectoryScopeId = $role.DirectoryScopeId
                ScopeDisplayName = $scopeDisplayName
                ScopeType = $scopeType
                RoleType = "Entra ID"
            }
        }
    }
    catch {
        Write-Host "Warning: Failed to retrieve Entra ID PIM roles: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return $entraRoles
}

# Function to display roles in a formatted table
function Show-AvailableRoles {
    param([array]$Roles)
    
    Write-Host "`nAvailable PIM Roles:" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    
    for ($i = 0; $i -lt $Roles.Count; $i++) {
        $role = $Roles[$i]
        Write-Host "$($i + 1). $($role.DisplayName)" -ForegroundColor Yellow
        Write-Host "   Type: $($role.RoleType)" -ForegroundColor Cyan
        Write-Host "   Scope: $($role.ScopeDisplayName) ($($role.ScopeType))" -ForegroundColor Gray
        Write-Host ""
    }
}

# Function to get Azure ARM headers (reusable)
function Get-AzureARMHeaders {
    try {
        $accessToken = Get-AzAccessToken -ResourceUrl "https://management.azure.com"
        
        # Handle Az module 14.* breaking change: Token is now SecureString instead of String
        if ($accessToken.Token -is [System.Security.SecureString]) {
            # Az 14.0.0+ returns SecureString
            $token = ConvertFrom-SecureString -SecureString $accessToken.Token -AsPlainText
        } else {
            # Az <14.0.0 returns plain string
            $token = $accessToken.Token
        }
        
        return @{ 
            Authorization = "Bearer $token"
            'Content-Type' = 'application/json'
        }
    }
    catch {
        throw "Failed to get Azure ARM access token: $($_.Exception.Message)"
    }
}

# Function to get Azure resource eligible roles
function Get-AzureResourceEligibleRoles {
    try {
        Write-Host "Retrieving Azure resource PIM eligible roles..." -ForegroundColor Gray
        
        # Use reusable headers function
        $headers = Get-AzureARMHeaders
        
        # Use the correct ARM API endpoint
        $url = "https://management.azure.com/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?api-version=2020-10-01&`$filter=asTarget()"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        
        # Cache for role definitions to avoid multiple API calls for the same role
        $roleDefinitionCache = @{}
        $failedRoleDefinitionIds = @{}

        $azureRoles = @()
        foreach ($assignment in $response.value) {
            $scope = $assignment.properties.scope
            $roleDefinitionId = $assignment.properties.roleDefinitionId

            # Get the role display name - first try from the response
            $roleDisplayName = $assignment.properties.roleDefinitionDisplayName

            # If display name is missing, look it up from the role definition
            if ([string]::IsNullOrWhiteSpace($roleDisplayName)) {
                if ($roleDefinitionCache.ContainsKey($roleDefinitionId)) {
                    $roleDisplayName = $roleDefinitionCache[$roleDefinitionId]
                } elseif ($failedRoleDefinitionIds.ContainsKey($roleDefinitionId)) {
                    $roleDisplayName = "Unknown Role ($($roleDefinitionId.Split('/')[-1]))"
                } else {
                    try {
                        # Look up the role definition to get the display name
                        $roleDefUrl = "https://management.azure.com$roleDefinitionId" + "?api-version=2022-04-01"
                        $roleDefResponse = Invoke-RestMethod -Uri $roleDefUrl -Headers $headers -Method Get
                        $roleDisplayName = $roleDefResponse.properties.roleName

                        # Cache the result
                        $roleDefinitionCache[$roleDefinitionId] = $roleDisplayName

                        Write-Host "   Retrieved role name: $roleDisplayName" -ForegroundColor Gray
                    }
                    catch {
                        Write-Host "Warning: Could not retrieve role definition for $roleDefinitionId : $($_.Exception.Message)" -ForegroundColor Yellow
                        $roleDisplayName = "Unknown Role ($($roleDefinitionId.Split('/')[-1]))"
                        $failedRoleDefinitionIds[$roleDefinitionId] = $true
                    }
                }
            }
            
            # Use helper function for scope parsing
            $scopeInfo = Get-ScopeDisplayInfo -Scope $scope
            
            $azureRoles += [PSCustomObject]@{
                DisplayName = $roleDisplayName
                RoleDefinitionId = $roleDefinitionId
                PrincipalId = $assignment.properties.principalId
                Scope = $scope
                ScopeDisplayName = $scopeInfo.DisplayName
                ScopeType = $scopeInfo.Type
                RoleType = "Azure Resource"
            }
        }
        
        Write-Host "   Found $($azureRoles.Count) Azure resource eligible roles" -ForegroundColor Gray
        return $azureRoles
    }
    catch {
        Write-Host "Warning: Could not retrieve Azure resource roles: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}

# Function to parse scope information (reusable)
function Get-ScopeDisplayInfo {
    param([string]$Scope)
    
    $scopeDisplayName = $Scope
    $scopeType = "Resource"
    
    if ($Scope -match '/subscriptions/([^/]+)') {
        $subId = $matches[1]
        if ($Scope -match '/resourceGroups/([^/]+)') {
            $rgName = $matches[1]
            $scopeType = "Resource Group"
            $scopeDisplayName = "RG: $rgName"
        } elseif ($Scope -match '/providers/([^/]+)/([^/]+)/([^/]+)') {
            $resourceType = $matches[2]
            $resourceName = $matches[3]
            $scopeType = "Resource"
            $scopeDisplayName = "$resourceType`: $resourceName"
        } else {
            $scopeType = "Subscription"
            try {
                $sub = Get-AzSubscription -SubscriptionId $subId -ErrorAction SilentlyContinue
                $scopeDisplayName = "Sub: $($sub.Name)"
            } catch {
                $scopeDisplayName = "Sub: $subId"
            }
        }
    }
    
    return @{
        DisplayName = $scopeDisplayName
        Type = $scopeType
    }
}

# Function to activate a single role
function Invoke-RoleActivation {
    param(
        [PSCustomObject]$Role,
        [string]$Justification = "Administrative work requirement",
        [string]$CurrentUserObjectId
    )
    
    try {
        Write-Host "Activating $($Role.RoleType) role: $($Role.DisplayName) at scope: $($Role.ScopeDisplayName)..." -ForegroundColor Yellow
        
        if ($Role.RoleType -eq "Entra ID") {
            # Entra ID role activation
            $params = @{
                Action = "selfActivate"
                PrincipalId = $Role.PrincipalId
                RoleDefinitionId = $Role.RoleDefinitionId
                DirectoryScopeId = $Role.DirectoryScopeId
                Justification = $Justification
                ScheduleInfo = @{
                    StartDateTime = Get-Date
                    Expiration = @{
                        Type = "AfterDuration"
                        Duration = "PT8H"  # 8 hours duration
                    }
                }
            }
            
            New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params
        }
        else {
            # Azure resource role activation
            $guid = [guid]::NewGuid().ToString()
            $startTime = (Get-Date).ToString("o")
            
            # Try activation with error suppression and proper handling
            try {
                $result = New-AzRoleAssignmentScheduleRequest `
                    -Name $guid `
                    -Scope $Role.Scope `
                    -ExpirationDuration "PT8H" `
                    -ExpirationType "AfterDuration" `
                    -PrincipalId $CurrentUserObjectId `
                    -RequestType "SelfActivate" `
                    -RoleDefinitionId $Role.RoleDefinitionId `
                    -ScheduleInfoStartDateTime $startTime `
                    -Justification $Justification `
                    -ErrorAction Stop 2>$null
            }
            catch {
                $errorMessage = $_.Exception.Message
                if ($errorMessage -like "*already exists*" -or $errorMessage -like "*Role assignment already exists*") {
                    Write-Host "⚠ Role already active (detected during activation): $($Role.DisplayName) ($($Role.ScopeDisplayName))" -ForegroundColor Yellow
                    return $true
                }
                else {
                    throw  # Re-throw if it's a different error
                }
            }
        }
        
        Write-Host "✓ Successfully activated: $($Role.DisplayName) at scope: $($Role.ScopeDisplayName) [Active for 8 hours]" -ForegroundColor Green
        return $true
    }
    catch {
        # Handle specific error cases more gracefully
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*already exists*" -or $errorMessage -like "*Role assignment already exists*") {
            Write-Host "⚠ Role already active: $($Role.DisplayName) ($($Role.ScopeDisplayName))" -ForegroundColor Yellow
            return $true
        }
        else {
            Write-Host "✗ Failed to activate $($Role.DisplayName) at scope: $($Role.ScopeDisplayName): $errorMessage" -ForegroundColor Red
            return $false
        }
    }
}

# Main script execution
try {
    # Initialize connections
    $connections = Initialize-Connections

    # Get current user context
    $currentUserObjectId = $null
    try {
        $currentUser = (Get-MgUser -UserId $connections.MgContext.Account).Id
        $currentUserObjectId = $currentUser
        Write-Host "Current user ID: $currentUser" -ForegroundColor Gray
    }
    catch {
        Write-Host "Warning: Could not retrieve user details. Using account from context." -ForegroundColor Yellow
        $currentUser = $connections.MgContext.Account
        # Try to get object ID from Azure context if available
        if ($connections.AzContext) {
            $currentUserObjectId = $connections.AzContext.Account.ExtendedProperties.HomeAccountId.Split('.')[0]
        }
    }

    # Get all eligible roles
    $entraRoles = Get-EntraIDEligibleRoles -CurrentUser $currentUser
    $azureRoles = @()
    if ($connections.AzContext) {
        $azureRoles = Get-AzureResourceEligibleRoles
    }

    # Combine all roles
    $allRoles = @()
    $allRoles += $entraRoles
    $allRoles += $azureRoles

    if ($allRoles.Count -eq 0) {
        Write-Host "No eligible PIM roles found for current user." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "`nFound $($entraRoles.Count) Entra ID roles and $($azureRoles.Count) Azure resource roles" -ForegroundColor Cyan

    # Display available roles
    Show-AvailableRoles -Roles $allRoles

    # Prompt user for selection
    Write-Host "Options:" -ForegroundColor Cyan
    Write-Host "  Enter role number(s) (e.g., 1,3,5 for multiple roles)" -ForegroundColor White
    Write-Host "  Enter 'ALL' to activate all roles" -ForegroundColor White
    Write-Host "  Enter 'Q' to quit" -ForegroundColor White
    
    $userChoice = Read-Host "`nPlease select your option"
    
    if ($userChoice.ToUpper() -eq 'Q') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }

    # Get justification from user
    $justification = Read-Host "Enter justification for role activation (press Enter for default)"
    if ([string]::IsNullOrWhiteSpace($justification)) {
        $justification = "Administrative work requirement"
    }

    $activatedCount = 0
    $failedCount = 0

    if ($userChoice.ToUpper() -eq 'ALL') {
        # Activate all roles
        Write-Host "`nActivating all eligible roles..." -ForegroundColor Cyan
        
        foreach ($role in $allRoles) {
            if (Invoke-RoleActivation -Role $role -Justification $justification -CurrentUserObjectId $currentUserObjectId) {
                $activatedCount++
            } else {
                $failedCount++
            }
            Start-Sleep -Seconds 1  # Brief pause between activations
        }
    }
    else {
        # Activate selected roles
        $selectedNumbers = $userChoice -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($number in $selectedNumbers) {
            if ($number -match '^\d+$' -and [int]$number -ge 1 -and [int]$number -le $allRoles.Count) {
                $roleIndex = [int]$number - 1
                if (Invoke-RoleActivation -Role $allRoles[$roleIndex] -Justification $justification -CurrentUserObjectId $currentUserObjectId) {
                    $activatedCount++
                } else {
                    $failedCount++
                }
                Start-Sleep -Seconds 1  # Brief pause between activations
            }
            else {
                Write-Host "Invalid selection: $number" -ForegroundColor Red
                $failedCount++
            }
        }
    }

    # Summary
    Write-Host "`n"
    Write-Host ("***") -ForegroundColor Green
    Write-Host "ACTIVATION SUMMARY" -ForegroundColor Green
    Write-Host ("***") -ForegroundColor Green
    Write-Host "Successfully activated: $activatedCount roles" -ForegroundColor Green
    if ($failedCount -gt 0) {
        Write-Host "Failed activations: $failedCount" -ForegroundColor Red
    }
    Write-Host "Roles will be active for 8 hours from activation time." -ForegroundColor Cyan
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure you have the required permissions and Microsoft Graph PowerShell module is installed." -ForegroundColor Yellow
    exit 1
}
finally {
    Write-Host "`nScript execution completed." -ForegroundColor Gray
}