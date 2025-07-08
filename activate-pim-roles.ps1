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

# Function to get Azure resource eligible roles
function Get-AzureResourceEligibleRoles {
    try {
        Write-Host "Retrieving Azure resource PIM eligible roles..." -ForegroundColor Gray
        
        # Get access token for ARM
        $token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token
        $headers = @{ Authorization = "Bearer $token" }
        
        # Use the correct ARM API endpoint
        $url = "https://management.azure.com/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?api-version=2020-10-01&`$filter=asTarget()"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        
        $azureRoles = @()
        foreach ($assignment in $response.value) {
            $scope = $assignment.properties.scope
            $scopeParts = $scope -split '/'
            
            # Determine scope display name and type
            $scopeDisplayName = $scope
            $scopeType = "Resource"
            
            if ($scope -match '/subscriptions/([^/]+)') {
                $subId = $matches[1]
                if ($scope -match '/resourceGroups/([^/]+)') {
                    $rgName = $matches[1]
                    $scopeType = "Resource Group"
                    $scopeDisplayName = "RG: $rgName"
                } elseif ($scope -match '/providers/([^/]+)/([^/]+)/([^/]+)') {
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
            
            $azureRoles += [PSCustomObject]@{
                DisplayName = $assignment.properties.roleDefinitionDisplayName
                RoleDefinitionId = $assignment.properties.roleDefinitionId
                PrincipalId = $assignment.properties.principalId
                Scope = $scope
                ScopeDisplayName = $scopeDisplayName
                ScopeType = $scopeType
                RoleType = "Azure Resource"
            }
        }
        
        return $azureRoles
    }
    catch {
        Write-Host "Warning: Could not retrieve Azure resource roles: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}

# Function to check if a role is already active
function Test-RoleActiveStatus {
    param(
        [PSCustomObject]$Role,
        [string]$CurrentUserObjectId
    )
    
    try {
        if ($Role.RoleType -eq "Entra ID") {
            # Check for active Entra ID role assignments
            $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "principalId eq '$($Role.PrincipalId)' and roleDefinitionId eq '$($Role.RoleDefinitionId)' and directoryScopeId eq '$($Role.DirectoryScopeId)'" -ErrorAction SilentlyContinue
            return $activeAssignments.Count -gt 0
        }
        else {
            # Check for active Azure resource role assignments
            $token = (Get-AzAccessToken -ResourceUrl "https://management.azure.com").Token
            $headers = @{ Authorization = "Bearer $token" }
            
            # Check for active role assignment schedule instances
            $encodedScope = [System.Web.HttpUtility]::UrlEncode($Role.Scope)
            $url = "https://management.azure.com/providers/Microsoft.Authorization/roleAssignmentScheduleInstances?api-version=2020-10-01&`$filter=asTarget() and principalId eq '$CurrentUserObjectId' and roleDefinitionId eq '$($Role.RoleDefinitionId)' and scope eq '$($Role.Scope)'"
            
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get -ErrorAction SilentlyContinue
            
            # Check if there are any active assignments
            $activeAssignments = $response.value | Where-Object { 
                $_.properties.status -eq "Accepted" -and 
                $_.properties.scope -eq $Role.Scope -and
                $_.properties.roleDefinitionId -eq $Role.RoleDefinitionId
            }
            
            return $activeAssignments.Count -gt 0
        }
    }
    catch {
        Write-Host "Warning: Could not check active status for $($Role.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
        return $false  # Assume not active if we can't check
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
        # Check if role is already active
        if (Test-RoleActiveStatus -Role $Role -CurrentUserObjectId $CurrentUserObjectId) {
            Write-Host "⚠ Role already active: $($Role.DisplayName)" -ForegroundColor Yellow
            return $true
        }
        
        Write-Host "Activating $($Role.RoleType) role: $($Role.DisplayName)..." -ForegroundColor Yellow
        
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
            # Azure resource role activation - use the passed CurrentUserObjectId
            $guid = [guid]::NewGuid().ToString()
            $startTime = (Get-Date).ToString("o")
            
            # Use the current user's object ID as the principal ID for self-activation
            New-AzRoleAssignmentScheduleRequest `
                -Name $guid `
                -Scope $Role.Scope `
                -ExpirationDuration "PT8H" `
                -ExpirationType "AfterDuration" `
                -PrincipalId $CurrentUserObjectId `
                -RequestType "SelfActivate" `
                -RoleDefinitionId $Role.RoleDefinitionId `
                -ScheduleInfoStartDateTime $startTime `
                -Justification $Justification
        }
        
        Write-Host "✓ Successfully activated: $($Role.DisplayName)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to activate $($Role.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   Scope: $($Role.Scope)" -ForegroundColor Gray
        return $false
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
    Write-Host "`n" + "="*50 -ForegroundColor Green
    Write-Host "ACTIVATION SUMMARY" -ForegroundColor Green
    Write-Host "="*50 -ForegroundColor Green
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