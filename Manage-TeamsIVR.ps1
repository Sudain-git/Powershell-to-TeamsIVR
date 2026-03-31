#Requires -Modules MicrosoftTeams

<#
.SYNOPSIS
    Reads and modifies IVR (Auto Attendant / Call Queue) configurations in Microsoft Teams.

.DESCRIPTION
    Connects to a Microsoft Teams tenant and provides functions to read
    and update IVR configurations (Auto Attendants and Call Queues).

.NOTES
    Required module: MicrosoftTeams
    Install with: Install-Module -Name MicrosoftTeams -Force -AllowClobber
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId
)

#region Connection

function Connect-TeamsEnvironment {
    [CmdletBinding()]
    param(
        [string]$TenantId
    )

    Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan

    $connectParams = @{}
    if ($TenantId) {
        $connectParams['TenantId'] = $TenantId
    }

    Connect-MicrosoftTeams @connectParams
    Write-Host "Connected." -ForegroundColor Green
}

#endregion

#region IVR - Auto Attendants

function Get-TeamsAutoAttendants {
    [CmdletBinding()]
    param()

    Write-Host "Retrieving Auto Attendants..." -ForegroundColor Cyan
    Get-CsAutoAttendant
}

#endregion

#region IVR - Call Queues

function Get-TeamsCallQueues {
    [CmdletBinding()]
    param()

    Write-Host "Retrieving Call Queues..." -ForegroundColor Cyan
    Get-CsCallQueue
}

#endregion

# --- Main ---

Connect-TeamsEnvironment -TenantId $TenantId
