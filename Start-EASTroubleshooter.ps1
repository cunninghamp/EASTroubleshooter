<#
.SYNOPSIS
EASTroubleshooter - Exchange ActiveSync Troubleshooting Script

.DESCRIPTION 
EASTroubleshooter is a PowerShell script that helps you to troubleshoot
Exchange ActiveSync device problems by collecting relevant information
about a mailbox's configuration and device associations.

Please refer to the installation and usage instructions at http://bit.ly/eastroubleshooter

.OUTPUTS
Results are output to the PowerShell console.

.PARAMETER Mailbox
Specifies the mailbox you are troubleshooting device issues for.

.EXAMPLE
.\Start-ExchangeAnalyzer.ps1 -Mailbox alan.reid

.LINK
http://bit.ly/eastroubleshooter

.NOTES
Written by: Paul Cunningham

Find me on:

* My Blog:	https://paulcunningham.me
* Twitter:	https://twitter.com/paulcunningham
* LinkedIn:	https://au.linkedin.com/in/cunninghamp/
* Github:	https://github.com/cunninghamp

*** License ***

The MIT License (MIT)

Copyright (c) 2017 Paul Cunningham

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

#...................................
#   Parameters
#...................................

# Identity of the mailbox the troubleshooter is going to look at
param (
    [Parameter(Mandatory=$true)]
    [String]$Mailbox
)
#endregion


#...................................
#   Functions
#...................................

# This function is used to write output to the console. 
# Example: Write-ResultsToConsole -Pretext "Something" -Result $something
function Write-ResultsToConsole() {
    param (
        [Parameter(Mandatory=$true)]
        $Pretext,

        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [AllowNull()]
        $Result,

        [Parameter(Mandatory=$false)]
        $ResultColor
    )

    if ($ResultColor) {
        Write-Host -ForegroundColor White "$($Pretext): " -NoNewLine
        Write-Host -ForeGroundColor $ResultColor $Result
    }
    else {
        Write-Host -ForegroundColor White "$($Pretext): " -NoNewLine
        Write-Host $Result
    }
}

# This function returns the device friendly name, or an "Unknown" result
# if the property is null or doesn't exist.
function Get-MobileDeviceFriendlyName ($DeviceId) {
    $MobileDeviceDetails = $MobileDevices | Where {$_.DeviceId -ieq $DeviceID}
    if (-not($MobileDeviceDetails.FriendlyName)) {
        $FriendlyName = "Unknown"
    }
    else {
        $FriendlyName = $MobileDeviceDetails.FriendlyName
    }
    return $FriendlyName
}

# This function calculates the number of days between now and the specified end date
function Get-NumberOfDaysSince() {
    param (
        [parameter(Mandatory=$true)]
        $Date
    )

    $Result = "{0:N0}" -f (New-TimeSpan -Start $Date -End $Now).TotalDays

    return $Result
}


#...................................
#   Script
#...................................

# If the script is running in a shell without Exchange cmdlets this is a good time to exit.
if (-not (Get-Command Get-CASMailbox)) {
    throw "The Get-CASMailbox cmdlet is not available. Please run this script from the Exchange Management Shell or an Exchange remote session."
}

# Get the current date/time for comparisons
$Now = Get-Date

# Display some summary info to get started
Write-Host ""
Write-Host -ForegroundColor White "============================================"
Write-Host -ForegroundColor White "            EAS Troubleshooter"
Write-Host -ForegroundColor White "============================================"
Write-Host ""

Write-ResultsToConsole -Pretext "Mailbox" -Result $Mailbox

# Retrieve the CAS Mailbox properties for the mailbox
try {
    $CASMailbox = Get-CASMailbox -Identity $Mailbox -ErrorAction STOP
}
catch {
    # Exit with error if cmdlet failed
    throw $_.Exception.Message
}

#Retrieve all the mobile devices associated with the mailbox
$MobileDevices = @(Get-MobileDevice -Mailbox $Mailbox)
Write-ResultsToConsole -Pretext "Mobile devices" -Result $MobileDevices.Count

# Use the Active Directory module if available to check the permissions inheritance flag
if (-not(Get-Module -ListAvailable ActiveDirectory)) {
    Write-ResultsToConsole -Pretext "AD Perms Inheritance" -Result "AD PowerShell module not available, skipping this check"
} 
else {
    #Skip the AD permissions inheritance check for cloud mailbox users
    if (-not($CASMailbox.DistinguishedName -like "*onmicrosoft.com*")) {
        $samAccountName = $CASMailbox.SamAccountName
        try {
            $ADUser = Get-ADUser -Filter 'SamAccountName -eq $samAccountName' -Properties nTSecurityDescriptor -ErrorAction STOP
            if ($ADUser) {
                $ADUserPretext = "AD Perms Inheritance"
                switch ($($ADUser.nTSecurityDescriptor.AreAccessRulesProtected)) {
                    $true { Write-ResultsToConsole -Pretext $ADUserPretext -Result "Disabled" -ResultColor "Red" }
                    $false { Write-ResultsToConsole -Pretext $ADUserPretext -Result "Enabled" -ResultColor "Green" }
                    default { Write-ResultsToConsole -Pretext $ADUserPretext -Result "Unknown" }
                }
            }
            else {
                Write-ResultsToConsole -Pretext "AD Perms Inheritance" -Result "Get-ADUser did not find a user"
            }

        }
        catch {
            Write-ResultsToConsole -Pretext "AD Perms Inheritance" -Result "Unable to find user in local Active Directory"
        }
    }
}

# Output the ActiveSync protocol state
$ActiveSyncProtocolPretext = "ActiveSync Protocol"
switch ($CASMailbox.ActiveSyncEnabled) {
    $true { Write-ResultsToConsole -Pretext $ActiveSyncProtocolPretext -Result "Enabled" -ResultColor "Green" }
    $false { Write-ResultsToConsole -Pretext $ActiveSyncProtocolPretext -Result "Disabled" -ResultColor "Red"}
}

# Output the EWS protocol state, access policy, and block list (if any)
$EWSProtocolPretext = "EWS Protocol"
switch ($CASMailbox.EWSEnabled) {
    $true { Write-ResultsToConsole -Pretext $EWSProtocolPretext -Result "Enabled" -ResultColor "Green" }
    $false { Write-ResultsToConsole -Pretext $EWSProtocolPretext -Result "Disabled" -ResultColor "Red"}
    default { Write-ResultsToConsole -Pretext $EWSProtocolPretext -Result "Not set" }
}

$EWSAccessPolicyPretext = "EWS Access Policy"
switch ($CASMailbox.EWSApplicationAccessPolicy) {
    "EnforceAllowList" { Write-ResultsToConsole -Pretext $EWSAccessPolicyPretext -Result "Enforce Allow List" -ResultColor "Yellow" }
    "EnforceBlockList" { Write-ResultsToConsole -Pretext $EWSAccessPolicyPretext -Result "Enforce Block List" -ResultColor "Yellow"}
    default { Write-ResultsToConsole -Pretext $EWSAccessPolicyPretext -Result "Not set" }
}

# If EWS block list is being used and there are blocked user agents, list those user agents
if ($CASMailbox.EWSApplicationAccessPolicy -eq "EnforceBlockList") {
    $EWSBlockListPretext = "EWS Blocked User Agents"
    if ($CASMailbox.EWSBlockList) {
        $EWSBlockedUserAgents = @($CASMailbox.EWSBlockList)
        Write-ResultsToConsole -Pretext $EWSBlockListPretext -Result $EWSBlockedUserAgents.Count
        foreach ($UserAgent in $EWSBlockedUserAgents) {
            Write-Host " - Blocked user agent: $UserAgent"
        }
    }    
}


# Retrieve the mobile device statistics (this part can be slow)
$MobileDeviceStats = @(foreach ($mobiledevice in $MobileDevices) {Get-MobileDeviceStatistics $mobiledevice.Identity})

Write-Host ""
Write-Host -ForegroundColor White " *** Mailbox Allow/Block Device ID List ***"
Write-Host ""

$AllowedDeviceIds = @($CASMailbox.ActiveSyncAllowedDeviceIds)
$BlockedDeviceIds = @($CASMailbox.ActiveSyncBlockedDeviceIds)

# Output the list of allowed device IDs
if ($AllowedDeviceIds.Count -gt 0) {
    Write-ResultsToConsole -Pretext "Allowed Device IDs" -Result "$($AllowedDeviceIds.Count) devices"
    #List each device ID and the device's friendly name
    foreach ($DeviceId in $AllowedDeviceIds) {
        Write-Host "ID: $($DeviceId) ($(Get-MobileDeviceFriendlyName $DeviceId))"
    }
}
else {
    #If there are no allowed device IDs output a result of None
    Write-ResultsToConsole -Pretext "Allowed Device IDs" -Result "None"
}

# Output the list of blocked device IDs
if ($BlockedDeviceIds.count -gt 0) {
    Write-ResultsToConsole -Pretext "Blocked Device IDs" -Result "$($BlockedDeviceIds.count) devices"
    #List each device ID and the device's friendly name
    foreach ($DeviceId in $BlockedDeviceIds) {
        Write-Host "ID: $DeviceId ($(Get-MobileDeviceFriendlyName $DeviceId))"
    }
}
else {
    #If there are no blocked device IDs output a result of None
    Write-ResultsToConsole -Pretext "Blocked Device IDs" -Result "None"
}

# Output information about each associated device
Write-Host ""
Write-Host -ForegroundColor White " ***        Mobile Device Details       ***"

foreach ($Device in $MobileDeviceStats) {
    Write-Host ""
    Write-ResultsToConsole -Pretext "*** Device ID" -Result "$($Device.DeviceId) ($($Device.DeviceFriendlyName))"
    Write-ResultsToConsole -Pretext "Client Type" -Result $Device.ClientType

    # Check if device ID is in block list
    $BlockedDevicePretext = "In ActiveSync Block List"
    if ($BlockedDeviceIds -icontains $Device.DeviceID) {
        Write-ResultsToConsole -Pretext $BlockedDevicePretext -Result "Yes" -ResultColor "Red"
    }
    else {
        Write-ResultsToConsole -Pretext $BlockedDevicePretext -Result "No"
    }

    # Output device access state, reason, and rule (if applicable)
    $AccessStatePretext = "ActiveSync Access State"
    switch ($Device.DeviceAccessState) {
        "Allowed" {Write-ResultsToConsole -Pretext $AccessStatePretext -Result $Device.DeviceAccessState -ResultColor "Green"}
        "Quarantined" {Write-ResultsToConsole -Pretext $AccessStatePretext -Result $Device.DeviceAccessState -ResultColor "Yellow"}
        "Blocked" {Write-ResultsToConsole -Pretext $AccessStatePretext -Result $Device.DeviceAccessState -ResultColor "Red"}
        default {Write-ResultsToConsole -Pretext $AccessStatePretext -Result $Device.DeviceAccessState}
    }

    Write-ResultsToConsole -Pretext "Access State Reason" -Result $Device.DeviceAccessStateReason

    if ($Device.DeviceAccessControlRule) {
        Write-ResultsToConsole -Pretext "Device Rule" -Result $Device.DeviceAccessControlRule
    }

    # Output the first/last sync times and compare last sync attempt and success
    $FirstSyncLocalTime = $Device.FirstSyncTime.ToLocalTime()
    $FirstSyncTimeSpan = Get-NumberOfDaysSince -Date $FirstSyncLocalTime
    Write-ResultsToConsole -Pretext "First Sync Time (UTC)" -Result $Device.FirstSyncTime
    Write-ResultsToConsole -Pretext "First Sync Time (local)" -Result "$($FirstSyncLocalTime) ($($FirstSyncTimeSpan) days ago)"
    
    $LastSyncAttemptLocalTime = $Device.LastSyncAttemptTime.ToLocalTime()
    $LastSyncAttemptSpan = Get-NumberOfDaysSince -Date $LastSyncAttemptLocalTime
    $LastSuccessSyncLocalTime = $Device.LastSuccessSync.ToLocalTime()
    $LastSuccessSyncSpan = Get-NumberOfDaysSince -Date $LastSyncAttemptLocalTime
    Write-ResultsToConsole -Pretext "Last Sync Attempt (local)" -Result "$($LastSyncAttemptLocalTime) ($($LastSyncAttemptSpan) days ago)"
    Write-ResultsToConsole -Pretext "Last Successful Sync (local)" -Result "$($LastSuccessSyncLocalTime) ($($LastSuccessSyncSpan) days ago)"

    $LastSyncMatch = $Device.LastSyncAttemptTime.ToLongDateString() -eq $Device.LastSuccessSync.ToLongDateString()
    switch ($LastSyncMatch) {
        $true {Write-ResultsToConsole -Pretext "Last Sync Times Match" -Result "Yes" -ResultColor "Green"}
        $false {
            $SyncDiff = "{0:N2}" -f (New-TimeSpan -Start $LastSyncAttemptLocalTime -End $LastSuccessSyncLocalTime).TotalMinutes
            if ($SyncDiff -gt 30) {
                $SyncDiffColor = "Red"
            }
            else {
                $SyncDiffColor = "Yellow"
            }
            Write-ResultsToConsole -Pretext "Last Sync Times Match" -Result "No ($($SyncDiff) minutes difference)" -ResultColor $SyncDiffColor
        }
    }

}

# All done
Write-Host ""
Write-Host -ForegroundColor White "============================================"
Write-Host -ForegroundColor White "                 Finished"
Write-Host ""
Write-Host -ForeGroundColor White " If you need help interpreting the output"
Write-Host -ForeGroundColor White " from this script please read the FAQ"
Write-Host -ForeGroundColor White " at: http://bit.ly/eastroubleshooter"
Write-Host ""
Write-Host -ForegroundColor White "============================================"
Write-Host ""

#...................................
#   Finished
#...................................
