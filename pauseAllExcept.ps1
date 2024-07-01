# DISCLAIMER:
# Copyright (c) Microsoft Corporation. All rights reserved. This 
# script is made available to you without any express, implied or 
# statutory warranty, not even the implied warranty of 
# merchantability or fitness for a particular purpose, or the 
# warranty of title or non-infringement. The entire risk of the 
# use or the results from the use of this script remains with you.
#
#
#
# Usage: powershell.exe .\PauseAllExcept.ps1
# This script allows you to pause all groups with dynamic membership except the ones you provides IDs for.
# It can be helpful when you need to mitigate ongoing issues with your groups that have dynamic membership rules.

# This function checks if you are already connected to Microsoft Graph.
#     If yes,
#         It disconnects to fetch your current information. Then, it prompts you to confirm the your current information.
#             If you confirm, it reconnects to Microsoft Graph with Group.ReadWrite.All permissions.
#             If you don't confirm, it will not reconnect and will prompt you to connect manually using Connect-MgGraph.
#     If not,
#         It attempts to connect to Microsoft Graph with Group.ReadWrite.All permissions.
#         If it fails to connect, it informs you that the Microsoft.Graph module might not be installed and provides the command to install it.
function ConnectToGraph {
    param (
        [string]$environment
    )
    # Check if already connected to Microsoft Graph
    if (Get-MgContext) {
        # Disconnect to fetch your current information
        $accountInfo = Disconnect-MgGraph
        Write-Host "MAKE SURE THE BELOW ACCOUNT/CLIENT APPLICATION HAS THE RIGHT SET OF PERMISSIONS TO PAUSE GROUPS WITH DYNAMIC MEMBERSHIP" -ForegroundColor Yellow
        Write-Host "Confirm the account: $($accountInfo.Account), TenantId: $($accountInfo.TenantId), and ClientId: $($accountInfo.ClientId)" -ForegroundColor Yellow
        $input = Read-Host "Type 'yes' to confirm: "
        if ($input.Trim().ToLower() -eq "yes") {
            # Reconnect with necessary permissions
            Connect-MgGraph -Environment $environment -Scopes "Group.ReadWrite.All"
        } else {
            # Inform you to reconnect manually
            Write-Host "Information not confirmed. Either re-run the script to confirm again or call <Connect-MgGraph> to log in using a different account." -ForegroundColor Yellow
            exit 1
        }
    } else {
        # Attempt to connect with necessary permissions
        Connect-MgGraph -Scopes "Group.ReadWrite.All"
        if (Get-MgContext) {
            # Recursive call to confirm your information
            ConnectToGraph -environment $environment
        } else {
            # Inform you to install Microsoft.Graph module if not connected
            Write-Host "If the Microsoft.Graph module is not installed, you need to install it to run this script." -ForegroundColor Yellow
            Write-Host "Run <Install-Module Microsoft.Graph -Scope CurrentUser> as an administrator." -ForegroundColor Yellow
            exit 1
        }
    }
}

# This function fetches a page of groups from Microsoft Graph.
# It returns the groups and the next page token if available.
function PageableFetchFromGraph {
    param (
        [string] $uri
    )
    try {
        # Make GET request to fetch groups
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri
        $dynamicGroups = $response.value
        $nextPage = $response.'@odata.nextLink'
        Write-Host "Found $($dynamicGroups.count) groups with dynamic membership." -ForegroundColor Green
        return $dynamicGroups, $nextPage
    } catch {
        # Handle any errors during fetch
        $errorMessage = $_.Exception.Message
        Write-Host "Failed to fetch groups with dynamic membership. Error message: $($errorMessage)" -ForegroundColor Red
        return $null
    }
}

# This function pauses a specific group with dynamic membership.
# It handles any errors, including throttling, and retries if necessary.
function PauseGroup {
    param (
        [string] $url, [string] $groupId
    )
    try {
        # Make PATCH request to pause the group
        Invoke-MgGraphRequest -Uri $url/$groupId -Method PATCH -Headers @{ConsistencyLevel = "eventual"} -Body $pauseDGjson
        Write-Host "Successfully paused group with dynamic membership. Id: $groupId" -ForegroundColor Green
        return $true
    } catch {
        # Handle errors and throttling
        $errorMessage = $_.Exception.Message
        $statusCode = $ErrorRecord.Exception.Response.StatusCode
        if ($statusCode -eq 429) {
            try {
                # Handle throttling if status code is 429
                HandleThrottling $_
                Write-Host "Throttling mitigated. Successfully paused group with dynamic membership. Id: $groupId" -ForegroundColor Green
                return $true
            } catch {
                # Handle failure after throttling mitigation
                $errorMessage = $_.Exception.Message
                Write-Host "Failed to pause group with dynamic membership. Id: $groupId. Error message: $errorMessage" -ForegroundColor Red
                return $false
            }
        } else {
            # Handle other errors
            Write-Host "Failed to pause group with dynamic membership. Id: $groupId. Error message: $errorMessage" -ForegroundColor Red
            return $false
        }
    }
}

# Function to validate GUID
function Is-Guid {
    param (
        [string]$Guid
    )
    return [guid]::TryParse($Guid, [ref]([guid]::Empty))
}

# This function prompts you to enter the IDs of the groups you want to skip while pausing all other groups with dynamic membership.
# It reports the number of successfully paused and failed groups.
function PauseAllGroupsExceptSpecified {
    param (
        [string] $graphEndpoint
    )
    # Prompt you to enter the ID of the groups you want to be excluded.
    $groupIdList = @()
    Write-Host "Enter the Group IDs (separated by comma if multiple):" -ForegroundColor Yellow
    $inputIds = Read-Host

    # Extracting individual group IDs
    $groupIdList = $inputIds -split ',' | ForEach-Object { $_.Trim() }

    # Validate each Group ID and remove invalid ones
    $validGroupIdList = @()
    $invalidIds = @()
    foreach ($id in $groupIdList) {
        if (Is-Guid $id) {
            $validGroupIdList += $id
        } else {
            $invalidIds += $id
        }
    }

    if ($invalidIds.Count -gt 0) {
        Write-Host "The following IDs are not valid GUIDs and will be removed:" -ForegroundColor Red
        foreach ($invalidId in $invalidIds) {
            Write-Host $invalidId
        }
    }

    # Use the valid list going forward
    $groupIdList = $validGroupIdList

    if ($groupIdList.Count -eq 0) {
        Write-Host "No valid IDs entered. Please make sure to enter valid IDs in GUID format and re-run the script." -ForegroundColor Red
        exit 1
    }

    Write-Host "Confirm that you entered $($groupIdList.count) valid IDs, as displayed here:" -ForegroundColor Yellow

    # Output each valid Group ID on a new line
    foreach ($id in $groupIdList) {
        Write-Host $id
    }

    $input = Read-Host "Type 'yes' to confirm: "
    if ($input.Trim().ToLower() -ne "yes") {
        exit 1
    }

    Write-Host "You have confirmed the entry of valid Group IDs." -ForegroundColor Green

    # Initialize URI for the first page
    $url = "$graphEndpoint/v1.0/groups"
    $filter = "?`$filter=groupTypes/any(c:c eq 'DynamicMembership')"
    $uri = "$url$filter"

    # Initialize success and failure count.
    $successCount = 0
    $failureCount = 0

    # Pause all groups page by page except specified ones
    do {
        # Fetch a page of groups
        $groupsData = PageableFetchFromGraph -uri $uri
        if ($groupsData -eq $null) {
            # If failed to fetch groups, skip to the next page
            break
        }
        $dynamicGroups = $groupsData[0]
        $nextPage = $groupsData[1]

        # Pause each group in the current page except specified ones
        foreach ($group in $dynamicGroups) {
            if ($excludedGroupIdList -contains $group.id) {
                Write-Host "Group excluded as per your request. Id: $($group.Id)" -ForegroundColor Yellow
                continue
            }
            if ($group.membershipRuleProcessingState -ceq "On") {
                $result = PauseGroup -url $url -groupId $group.id
                if ($result) {
                    $successCount++
                } else {
                    $failureCount++
                }
            } else {
                Write-Host "Group skipped because it was found to be in $($group.membershipRuleProcessingState) state. Id: $($group.Id)" -ForegroundColor Yellow
            }
        }

        # Move to the next page
        $uri = $nextPage
        Write-Host "Checking if more groups are present on next page." -ForegroundColor Yellow
    } while ($uri -ne $null)
    Write-Host "No more groups found." -ForegroundColor Yellow
    # Report the total count
    Write-Host "PauseAllExcept Operation Complete. Successfully Paused: $successCount, Failed: $failureCount" -ForegroundColor Yellow
}


# Internal function to handle throttling and retry requests
# This function handles throttling by waiting for the specified time before retrying the request.
function HandleThrottling {
    param (
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    # Throttling occurred, extract Retry-After header if available
    $retryAfter = $ErrorRecord.Exception.Response.Headers.'Retry-After'
    if ($retryAfter) {
        Write-Host "Throttling detected. Waiting for $retryAfter seconds before retrying..."
        Start-Sleep -Seconds $retryAfter
    } else {
        # If Retry-After header is not available, wait for a default time
        Write-Host "Throttling detected. Waiting for default time i.e. 60 seconds before retrying..."
        Start-Sleep -Seconds 60  # Wait for 60 seconds by default
    }
    # Retry the request
    Invoke-MgGraphRequest @($ErrorRecord.Exception.InvocationInfo.BoundParameters)
}

# Function to prompt you to select the environment for determining the Microsoft Graph endpoint.
# This function helps you choose the correct environment and returns the corresponding endpoint.
function GetEnvironmentAndEndpoint {
    # Prompt you to select the environment
    try {
        $environmentChoice = Read-Host "Please select the environment (default is 'Global'): `nOptions: Global, USGov, USGovDoD, China"
    } catch {
        $environmentChoice = 'Global'
    }

    # Normalize the input to lower case
    $environmentChoice = $environmentChoice.Trim().ToLower()

    # Set the default environment to global
    $selectedEnvironment = "Global"

    # Map your choice to the corresponding environment
    switch ($environmentChoice) {
        "usgov" { $selectedEnvironment = "USGov" }
        "usgovdod" { $selectedEnvironment = "USGovDoD" }
        "china" { $selectedEnvironment = "China" }
        default { $selectedEnvironment = "Global" }
    }

    # Dictionary to map environment names to their endpoints
    $endpoints = @{
        "Global"   = "https://graph.microsoft.com"
        "USGov"    = "https://graph.microsoft.us"
        "USGovDoD" = "https://dod-graph.microsoft.us"
        "China"    = "https://microsoftgraph.chinacloudapi.cn"
    }

    $graphEndpoint = $endpoints[$selectedEnvironment]

    Write-Host "Environment Selected: $($selectedEnvironment). It maps to the graph endpoint: $($graphEndpoint)" -ForegroundColor Green

    # Return the selected environment and graph endpoint
    return @{ "SelectedEnvironment" = $selectedEnvironment; "GraphEndpoint" = $graphEndpoint }
}

# Running the script:
# Prompt you to confirm if you want to run the pause all except flow.
Write-Host "DO YOU WANT TO PAUSE ALL GROUPS WITH DYNAMIC MEMBERSHIP EXCEPT SPECIFIED ONES?" -ForegroundColor Yellow
$input = Read-Host "Type 'yes' to confirm: "

#Global variable for JSON change.
$global:pauseDGjson = '{"membershipRuleProcessingState":"Paused"}'

# Start the pause all except flow if confirmed.
if ($input.Trim().ToLower() -eq "yes") {
    $result = GetEnvironmentAndEndpoint
    $selectedEnvironment = $result.SelectedEnvironment
    $graphEndpoint = $result.GraphEndpoint
    try {
        # Connect to Microsoft Graph
        ConnectToGraph -environment $selectedEnvironment
        # Pause all groups with dynamic membership except ones with user specified IDs
        PauseAllGroupsExceptSpecified -graphEndpoint $graphEndpoint
    } catch {
        # Handle any errors during the process
        $errorMessage = $_.Exception.Message
        Write-Host "PauseAllExcept operation failed. Error message: $($errorMessage)" -ForegroundColor Red
    }
# Inform you that the input was not accepted.
} else {
    Write-Host "PauseAllExcept script terminated. Please re-run the script and input 'yes' to run the PauseAllExcept script." -ForegroundColor Red
}
