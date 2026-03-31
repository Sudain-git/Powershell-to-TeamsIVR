#Requires -Modules MicrosoftTeams

<#
.SYNOPSIS
    Reads and modifies IVR (Auto Attendant / Call Queue) configurations in Microsoft Teams.

.DESCRIPTION
    Connects to a Microsoft Teams tenant and provides functions to read
    and update IVR configurations (Auto Attendants and Call Queues).

    This script supports two execution modes:

    USER INTERACTION MODE (default)
    --------------------------------
    Intended for direct human use. On startup, displays:
      - Script version
      - Hostname of the server running the script
      - The process/user that invoked the script
    Long-running operations display progress bars.
    Authentication uses interactive login (prompts for credentials).

    HEADLESS / CRON MODE
    --------------------------------
    Intended for scheduled execution (e.g., Task Scheduler, cron).
    Startup banner and progress bars are suppressed.
    Authentication must be non-interactive — likely via a service principal
    or certificate-based auth rather than interactive OAuth.
    TODO: Confirm auth method for headless mode (service principal vs. cert).

    PLANNED FUNCTIONS
    --------------------------------
    - Get-TeamsIVRList          : Retrieve all IVRs (Auto Attendants + Call Queues)
    - Get-TeamsIVRSchedule      : Retrieve current business hours / after-hours schedule for an IVR
    - Set-TeamsIVRSchedule      : Set/update an IVR's hours based on provided input
    - Test-TeamsIVRUserAccess   : Check whether the calling user is in the authorized users list
                                  for a given IVR before allowing changes

.PARAMETER TenantId
    The Azure AD Tenant ID to connect to. Optional — if omitted, uses the default tenant.

.PARAMETER Headless
    Switch to enable headless/cron mode. Suppresses banner and progress bars,
    and triggers non-interactive authentication.

.NOTES
    Script Version : 0.1
    Required module: MicrosoftTeams
    Install with   : Install-Module -Name MicrosoftTeams -Force -AllowClobber
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    # Run in headless/cron mode: suppresses banner/progress bars and uses non-interactive auth
    [Parameter(Mandatory = $false)]
    [switch]$Headless
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
