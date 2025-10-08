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
                IsActive = $false
                ActiveAssignmentId = $null
            }
        }
    }
    catch {
        Write-Host "Warning: Failed to retrieve Entra ID PIM roles: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return $entraRoles
}

# Function to get active Entra ID role assignments
function Get-EntraIDActiveRoles {
    param([string]$CurrentUser)
    
    Write-Host "Retrieving active Entra ID PIM roles..." -ForegroundColor Gray
    $activeRoles = @{
    }
    
    try {
        $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -ExpandProperty RoleDefinition -All -Filter "principalId eq '$CurrentUser'"
        
        foreach ($assignment in $activeAssignments) {
            $key = "$($assignment.RoleDefinitionId)|$($assignment.DirectoryScopeId)"
            $activeRoles[$key] = @{
                AssignmentId = $assignment.Id
                ActivatedDateTime = $assignment.StartDateTime
                ExpirationDateTime = $assignment.EndDateTime
            }
        }
    }
    catch {
        Write-Host "Warning: Failed to retrieve active Entra ID roles: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return $activeRoles
}

# Function to get active Azure resource role assignments
function Get-AzureResourceActiveRoles {
    param([string]$CurrentUserObjectId)
    
    Write-Host "Retrieving active Azure resource PIM roles..." -ForegroundColor Gray
    $activeRoles = @{
    }
    
    try {
        $headers = Get-AzureARMHeaders
        $url = "https://management.azure.com/providers/Microsoft.Authorization/roleAssignmentScheduleInstances?api-version=2020-10-01&`$filter=asTarget()"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get
        
        foreach ($assignment in $response.value) {
            if ($assignment.properties.principalId -eq $CurrentUserObjectId) {
                $key = "$($assignment.properties.roleDefinitionId)|$($assignment.properties.scope)"
                $activeRoles[$key] = @{
                    AssignmentId = $assignment.name
                    ActivatedDateTime = $assignment.properties.startDateTime
                    ExpirationDateTime = $assignment.properties.endDateTime
                    Scope = $assignment.properties.scope
                }
            }
        }
    }
    catch {
        Write-Host "Warning: Failed to retrieve active Azure resource roles: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return $activeRoles
}

# Function to display roles in a formatted table
function Show-AvailableRoles {
    param([array]$Roles)
    
    Write-Host "`nAvailable PIM Roles:" -ForegroundColor Green
    Write-Host "===========================================" -ForegroundColor Green
    
    for ($i = 0; $i -lt $Roles.Count; $i++) {
        $role = $Roles[$i]
        $statusIndicator = if ($role.IsActive) { "🟢 ACTIVE" } else { "⚪ Inactive" }
        
        Write-Host "$($i + 1). $($role.DisplayName) [$statusIndicator]" -ForegroundColor Yellow
        Write-Host "   Type: $($role.RoleType)" -ForegroundColor Cyan
        Write-Host "   Scope: $($role.ScopeDisplayName) ($($role.ScopeType))" -ForegroundColor Gray
        
        if ($role.IsActive -and $role.ExpirationDateTime) {
            Write-Host "   Expires: $($role.ExpirationDateTime)" -ForegroundColor Magenta
        }
        Write-Host ""
    }
}

# Function to get Azure ARM headers (reusable)
function Get-AzureARMHeaders {
    try {
        $accessToken = Get-AzAccessToken -ResourceUrl "https://management.azure.com"
        
        # https://github.com/Azure/azure-powershell/issues/25533
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
        $roleDefinitionCache = @{
        }
        $failedRoleDefinitionIds = @{
        }

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
                IsActive = $false
                ActiveAssignmentId = $null
                ExpirationDateTime = $null
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

# Function to deactivate a single role
function Invoke-RoleDeactivation {
    param(
        [PSCustomObject]$Role,
        [string]$CurrentUserObjectId
    )
    
    try {
        Write-Host "Deactivating $($Role.RoleType) role: $($Role.DisplayName) at scope: $($Role.ScopeDisplayName)..." -ForegroundColor Yellow
        
        if ($Role.RoleType -eq "Entra ID") {
            # Entra ID role deactivation
            $params = @{
                Action = "selfDeactivate"
                PrincipalId = $Role.PrincipalId
                RoleDefinitionId = $Role.RoleDefinitionId
                DirectoryScopeId = $Role.DirectoryScopeId
            }
            
            New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params | Out-Null
        }
        else {
            # Azure resource role deactivation
            $guid = [guid]::NewGuid().ToString()
            
            try {
                New-AzRoleAssignmentScheduleRequest `
                    -Name $guid `
                    -Scope $Role.Scope `
                    -PrincipalId $CurrentUserObjectId `
                    -RequestType "SelfDeactivate" `
                    -RoleDefinitionId $Role.RoleDefinitionId `
                    -ErrorAction Stop 2>$null | Out-Null
            }
            catch {
                $errorMessage = $_.Exception.Message
                if ($errorMessage -notlike "*not found*" -and $errorMessage -notlike "*does not exist*") {
                    throw
                }
            }
        }
        
        Write-Host "✓ Successfully deactivated: $($Role.DisplayName) at scope: $($Role.ScopeDisplayName)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ Failed to deactivate $($Role.DisplayName) at scope: $($Role.ScopeDisplayName): $($_.Exception.Message)" -ForegroundColor Red
        return $false
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
            
            New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params | Out-Null
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

    # Get active role assignments
    $activeEntraRoles = Get-EntraIDActiveRoles -CurrentUser $currentUser
    $activeAzureRoles = @()
    if ($connections.AzContext) {
        $activeAzureRoles = Get-AzureResourceActiveRoles -CurrentUserObjectId $currentUserObjectId
    }

    # Mark active roles in the eligible roles list
    foreach ($role in $entraRoles) {
        $key = "$($role.RoleDefinitionId)|$($role.DirectoryScopeId)"
        if ($activeEntraRoles.ContainsKey($key)) {
            $role.IsActive = $true
            $role.ActiveAssignmentId = $activeEntraRoles[$key].AssignmentId
            $role.ExpirationDateTime = $activeEntraRoles[$key].ExpirationDateTime
        }
    }

    foreach ($role in $azureRoles) {
        $key = "$($role.RoleDefinitionId)|$($role.Scope)"
        if ($activeAzureRoles.ContainsKey($key)) {
            $role.IsActive = $true
            $role.ActiveAssignmentId = $activeAzureRoles[$key].AssignmentId
            $role.ExpirationDateTime = $activeAzureRoles[$key].ExpirationDateTime
        }
    }

    # Combine all roles
    $allRoles = @()
    $allRoles += $entraRoles
    $allRoles += $azureRoles

    if ($allRoles.Count -eq 0) {
        Write-Host "No eligible PIM roles found for current user." -ForegroundColor Yellow
        exit 0
    }

    $activeCount = ($allRoles | Where-Object { $_.IsActive }).Count
    Write-Host "`nFound $($entraRoles.Count) Entra ID roles and $($azureRoles.Count) Azure resource roles ($activeCount currently active)" -ForegroundColor Cyan

    # Display available roles
    Show-AvailableRoles -Roles $allRoles

    # Prompt user for action
    Write-Host "Actions:" -ForegroundColor Cyan
    Write-Host "  [A] Activate role(s)" -ForegroundColor White
    Write-Host "  [D] Deactivate role(s)" -ForegroundColor White
    Write-Host "  [R] Reactivate role(s) (deactivate then activate)" -ForegroundColor White
    Write-Host "  [Q] Quit" -ForegroundColor White
    
    $action = Read-Host "`nSelect action (A/D/R/Q)"
    
    if ($action.ToUpper() -eq 'Q') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }

    if ($action.ToUpper() -notin @('A', 'D', 'R')) {
        Write-Host "Invalid action selected." -ForegroundColor Red
        exit 1
    }

    # Prompt user for role selection
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "  Enter role number(s) (e.g., 1,3,5 for multiple roles)" -ForegroundColor White
    Write-Host "  Enter 'ALL' to select all roles" -ForegroundColor White
    Write-Host "  Enter 'ACTIVE' to select all active roles" -ForegroundColor White
    Write-Host "  Enter 'INACTIVE' to select all inactive roles" -ForegroundColor White
    Write-Host "  Enter 'Q' to quit" -ForegroundColor White
    
    $userChoice = Read-Host "`nPlease select roles"
    
    if ($userChoice.ToUpper() -eq 'Q') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }

    # Determine selected roles
    $selectedRoles = @()
    if ($userChoice.ToUpper() -eq 'ALL') {
        $selectedRoles = $allRoles
    }
    elseif ($userChoice.ToUpper() -eq 'ACTIVE') {
        $selectedRoles = $allRoles | Where-Object { $_.IsActive }
        if ($selectedRoles.Count -eq 0) {
            Write-Host "No active roles found." -ForegroundColor Yellow
            exit 0
        }
    }
    elseif ($userChoice.ToUpper() -eq 'INACTIVE') {
        $selectedRoles = $allRoles | Where-Object { -not $_.IsActive }
        if ($selectedRoles.Count -eq 0) {
            Write-Host "No inactive roles found." -ForegroundColor Yellow
            exit 0
        }
    }
    else {
        $selectedNumbers = $userChoice -split ',' | ForEach-Object { $_.Trim() }
        foreach ($number in $selectedNumbers) {
            if ($number -match '^\d+$' -and [int]$number -ge 1 -and [int]$number -le $allRoles.Count) {
                $selectedRoles += $allRoles[[int]$number - 1]
            }
            else {
                Write-Host "Invalid selection: $number" -ForegroundColor Red
            }
        }
    }

    if ($selectedRoles.Count -eq 0) {
        Write-Host "No valid roles selected." -ForegroundColor Yellow
        exit 0
    }

    # Get justification if activating or reactivating
    $justification = "Administrative work requirement"
    if ($action.ToUpper() -in @('A', 'R')) {
        $justificationInput = Read-Host "Enter justification for role activation (press Enter for default)"
        if (-not [string]::IsNullOrWhiteSpace($justificationInput)) {
            $justification = $justificationInput
        }
    }

    $successCount = 0
    $failedCount = 0

    # Perform the requested action
    switch ($action.ToUpper()) {
        'A' {
            # Activate roles
            Write-Host "`nActivating $($selectedRoles.Count) role(s)..." -ForegroundColor Cyan
            foreach ($role in $selectedRoles) {
                if (Invoke-RoleActivation -Role $role -Justification $justification -CurrentUserObjectId $currentUserObjectId) {
                    $successCount++
                } else {
                    $failedCount++
                }
                Start-Sleep -Seconds 1
            }
        }
        'D' {
            # Deactivate roles
            Write-Host "`nDeactivating $($selectedRoles.Count) role(s)..." -ForegroundColor Cyan
            foreach ($role in $selectedRoles) {
                if ($role.IsActive) {
                    if (Invoke-RoleDeactivation -Role $role -CurrentUserObjectId $currentUserObjectId) {
                        $successCount++
                    } else {
                        $failedCount++
                    }
                } else {
                    Write-Host "⚠ Role already inactive: $($role.DisplayName) ($($role.ScopeDisplayName))" -ForegroundColor Yellow
                }
                Start-Sleep -Seconds 1
            }
        }
        'R' {
            # Reactivate roles (deactivate then activate)
            Write-Host "`nReactivating $($selectedRoles.Count) role(s)..." -ForegroundColor Cyan
            
            # First deactivate active roles
            $rolesToDeactivate = $selectedRoles | Where-Object { $_.IsActive }
            if ($rolesToDeactivate.Count -gt 0) {
                Write-Host "Step 1: Deactivating $($rolesToDeactivate.Count) active role(s)..." -ForegroundColor Cyan
                foreach ($role in $rolesToDeactivate) {
                    Invoke-RoleDeactivation -Role $role -CurrentUserObjectId $currentUserObjectId | Out-Null
                    Start-Sleep -Seconds 1
                }
                Write-Host "Waiting 5 seconds before reactivation..." -ForegroundColor Gray
                Start-Sleep -Seconds 5
            }
            
            # Then activate all selected roles
            Write-Host "Step 2: Activating $($selectedRoles.Count) role(s)..." -ForegroundColor Cyan
            foreach ($role in $selectedRoles) {
                if (Invoke-RoleActivation -Role $role -Justification $justification -CurrentUserObjectId $currentUserObjectId) {
                    $successCount++
                } else {
                    $failedCount++
                }
                Start-Sleep -Seconds 1
            }
        }
    }

    # Summary
    Write-Host "`n"
    Write-Host ("***") -ForegroundColor Green
    Write-Host "OPERATION SUMMARY" -ForegroundColor Green
    Write-Host ("***") -ForegroundColor Green
    Write-Host "Action: $($action.ToUpper())" -ForegroundColor Cyan
    Write-Host "Successfully processed: $successCount role(s)" -ForegroundColor Green
    if ($failedCount -gt 0) {
        Write-Host "Failed operations: $failedCount" -ForegroundColor Red
    }
    if ($action.ToUpper() -in @('A', 'R')) {
        Write-Host "Activated roles will be active for 8 hours from activation time." -ForegroundColor Cyan
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please ensure you have the required permissions and Microsoft Graph PowerShell module is installed." -ForegroundColor Yellow
    exit 1
}
finally {
    Write-Host "`nScript execution completed." -ForegroundColor Gray
}