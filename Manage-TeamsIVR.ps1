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

    FUNCTIONS
    --------------------------------
    - Get-TeamsIVRList          : Retrieve all IVRs (Auto Attendants + Call Queues)
    - Get-TeamsIVRSchedule      : Retrieve current schedule for an IVR.
                                  Returns AfterHours (the schedule Teams stores as closed/after-hours
                                  time ranges) and BusinessHours (all other schedules, typically empty).
    - Format-TeamsIVRSchedule   : Decode a raw CsOnlineSchedule into human-readable day/time lines.
                                  Pass -InvertToOpenHours when displaying an AfterHours schedule
                                  to show open/business hours instead of the raw closed periods.
    - Build-TeamsIVRSchedule    : Convert human-readable open-hours input (e.g. Monday = '9:00am - 5:00pm')
                                  into a CsOnlineSchedule ready for Set-TeamsIVRSchedule.
                                  Automatically inverts open hours to closed/after-hours ranges, because
                                  Teams AfterHours schedules store WHEN THE AA IS CLOSED, not when it is open.
    - Set-TeamsIVRSchedule      : Apply a CsOnlineSchedule to an IVR's AfterHours association.
                                  The schedule must contain closed/after-hours time ranges
                                  (use Build-TeamsIVRSchedule to produce the correctly formatted object).
    - Get-TeamsIVRsByUser       : Return all AAs where a given user appears in the AuthorizedUsers list.
    - Test-TeamsIVRUserAccess   : Check whether a given user is in the authorized users list
                                  for a specific IVR.

.PARAMETER TenantId
    The Azure AD Tenant ID to connect to. Optional — if omitted, uses the default tenant.

.PARAMETER Headless
    Switch to enable headless/cron mode. Suppresses banner and progress bars,
    and triggers non-interactive authentication.
    Non-interactive auth is designed but not yet activated -- three options are
    available in Connect-TeamsEnvironment (certificate, client secret, managed identity)
    that MSops must review and enable before headless auth is fully functional.

.NOTES
    Script Version : 0.2
    Required module: MicrosoftTeams
    Install with   : Install-Module -Name MicrosoftTeams -Force -AllowClobber
#>

[CmdletBinding(DefaultParameterSetName = 'ShowUsage')]
param(
    # ---- Shared parameters (available in all modes) ----

    [string]$TenantId,

    # Run in headless/cron mode: suppresses banner/progress bars and uses non-interactive auth
    [switch]$Headless,

    # =============================================================================
    # NON-INTERACTIVE AUTH PARAMETERS -- FOR MSOPS REVIEW
    # Uncomment the parameters for whichever auth method is approved below.
    # All three blocks correspond to Options A, B, and C in Connect-TeamsEnvironment.
    # NOTE: $ApplicationId must not be declared twice. Uncomment ONLY the block for
    #       the chosen option. MSops should remove the other commented block entirely.
    # =============================================================================

    # -- Option A: Service Principal + Certificate (RECOMMENDED) --
    # [Parameter(ParameterSetName = 'ShowUsage')]
    # [Parameter(ParameterSetName = 'GetAllIVRs')]
    # [Parameter(ParameterSetName = 'GetIVR')]
    # [Parameter(ParameterSetName = 'GetSchedule')]
    # [Parameter(ParameterSetName = 'DecodeSchedule')]
    # [Parameter(ParameterSetName = 'SetSchedule')]
    # [Parameter(ParameterSetName = 'GetAuthorizedUsers')]
    # [Parameter(ParameterSetName = 'CheckUserAccess')]
    # [Parameter(ParameterSetName = 'GetUserIVRs')]
    # [string]$ApplicationId,          # Azure AD App Registration client ID (GUID)
    #
    # [Parameter(ParameterSetName = 'ShowUsage')]
    # [Parameter(ParameterSetName = 'GetAllIVRs')]
    # [Parameter(ParameterSetName = 'GetIVR')]
    # [Parameter(ParameterSetName = 'GetSchedule')]
    # [Parameter(ParameterSetName = 'DecodeSchedule')]
    # [Parameter(ParameterSetName = 'SetSchedule')]
    # [Parameter(ParameterSetName = 'GetAuthorizedUsers')]
    # [Parameter(ParameterSetName = 'CheckUserAccess')]
    # [Parameter(ParameterSetName = 'GetUserIVRs')]
    # [string]$CertificateThumbprint,  # Thumbprint of cert installed in Windows cert store

    # -- Option B: Service Principal + Client Secret (alternative; uncomment instead of Option A) --
    # [Parameter(ParameterSetName = 'ShowUsage')]
    # [Parameter(ParameterSetName = 'GetAllIVRs')]
    # [Parameter(ParameterSetName = 'GetIVR')]
    # [Parameter(ParameterSetName = 'GetSchedule')]
    # [Parameter(ParameterSetName = 'DecodeSchedule')]
    # [Parameter(ParameterSetName = 'SetSchedule')]
    # [Parameter(ParameterSetName = 'GetAuthorizedUsers')]
    # [Parameter(ParameterSetName = 'CheckUserAccess')]
    # [Parameter(ParameterSetName = 'GetUserIVRs')]
    # [string]$ApplicationId,          # Azure AD App Registration client ID (GUID)
    #
    # [Parameter(ParameterSetName = 'ShowUsage')]
    # [Parameter(ParameterSetName = 'GetAllIVRs')]
    # [Parameter(ParameterSetName = 'GetIVR')]
    # [Parameter(ParameterSetName = 'GetSchedule')]
    # [Parameter(ParameterSetName = 'DecodeSchedule')]
    # [Parameter(ParameterSetName = 'SetSchedule')]
    # [Parameter(ParameterSetName = 'GetAuthorizedUsers')]
    # [Parameter(ParameterSetName = 'CheckUserAccess')]
    # [Parameter(ParameterSetName = 'GetUserIVRs')]
    # [SecureString]$ClientSecret,     # Client secret as SecureString

    # -- Option C: Managed Identity (Azure-hosted only, no credential params needed) --
    # No additional parameters required. Set -Headless and ensure the hosting environment
    # (Azure VM, App Service, Automation Account) has a managed identity with Teams Admin role.

    # ---- Action: Show usage reference (default when no args given) ----
    # Usage: .\Manage-TeamsIVR.ps1  OR  .\Manage-TeamsIVR.ps1 -Help

    [Parameter(ParameterSetName = 'ShowUsage', Mandatory = $false)]
    [switch]$Help,

    # ---- Action: Get all IVRs ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetAllIVRs

    [Parameter(ParameterSetName = 'GetAllIVRs', Mandatory = $true)]
    [switch]$GetAllIVRs,

    # ---- Action: Get a specific IVR by name ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetIVR -Name "My Auto Attendant"

    [Parameter(ParameterSetName = 'GetIVR', Mandatory = $true)]
    [switch]$GetIVR,

    [Parameter(ParameterSetName = 'GetIVR', Mandatory = $true)]
    [string]$Name,

    # ---- Action: Get the schedule for an IVR (raw object output) ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetSchedule -Identity "aa_guid_or_name"

    [Parameter(ParameterSetName = 'GetSchedule', Mandatory = $true)]
    [switch]$GetSchedule,

    # ---- Action: Decode the schedule into human-readable day/time format ----
    # Usage: .\Manage-TeamsIVR.ps1 -DecodeSchedule -Identity "aa_guid_or_name"

    [Parameter(ParameterSetName = 'DecodeSchedule', Mandatory = $true)]
    [switch]$DecodeSchedule,

    # ---- Action: Set the schedule for an IVR ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetSchedule -Identity "aa_guid_or_name" -Schedule $scheduleObject

    [Parameter(ParameterSetName = 'SetSchedule', Mandatory = $true)]
    [switch]$SetSchedule,

    [Parameter(ParameterSetName = 'SetSchedule', Mandatory = $true)]
    [ValidateNotNull()]
    [object]$Schedule,

    # ---- Action: List authorized users for an IVR ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetAuthorizedUsers -Identity "aa_guid_or_name"

    [Parameter(ParameterSetName = 'GetAuthorizedUsers', Mandatory = $true)]
    [switch]$GetAuthorizedUsers,

    # ---- Action: Check whether a user is authorized for an IVR ----
    # Usage: .\Manage-TeamsIVR.ps1 -CheckUserAccess -Identity "aa_guid_or_name" -User "user@contoso.com"

    [Parameter(ParameterSetName = 'CheckUserAccess', Mandatory = $true)]
    [switch]$CheckUserAccess,

    # ---- Action: List all IVRs a user is authorized to manage ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetUserIVRs -User "user@contoso.com"

    [Parameter(ParameterSetName = 'GetUserIVRs', Mandatory = $true)]
    [switch]$GetUserIVRs,

    # User identity — accepts UPN (user@domain.com), display name, or object ID (GUID)
    [Parameter(ParameterSetName = 'CheckUserAccess', Mandatory = $true)]
    [Parameter(ParameterSetName = 'GetUserIVRs',     Mandatory = $true)]
    [string]$User,

    # ---- Shared: Identity used by schedule and user-access actions ----

    [Parameter(ParameterSetName = 'GetSchedule', Mandatory = $true)]
    [Parameter(ParameterSetName = 'DecodeSchedule', Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetSchedule', Mandatory = $true)]
    [Parameter(ParameterSetName = 'GetAuthorizedUsers', Mandatory = $true)]
    [Parameter(ParameterSetName = 'CheckUserAccess', Mandatory = $true)]
    [string]$Identity
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Constants

$Script:Version = '0.2'

#endregion

#region Banner

function Show-Banner {
    <#
    .SYNOPSIS
        Displays startup information. Only called in user interaction mode.
    #>
    [CmdletBinding()]
    param()

    $invoker = try { "$env:USERDOMAIN\$env:USERNAME" } catch { $env:USERNAME }
    $hostname = try { [System.Net.Dns]::GetHostName() } catch { 'Unknown' }
    $process  = try { (Get-Process -Id $PID).Name } catch { 'Unknown' }

    Write-Host ""
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "  Teams IVR Manager  v$($Script:Version)" -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host "  Host    : $hostname"
    Write-Host "  User    : $invoker"
    Write-Host "  Process : $process (PID $PID)"
    Write-Host "=======================================" -ForegroundColor Cyan
    Write-Host ""
}

#endregion

#region Connection

function Connect-TeamsEnvironment {
    <#
    .SYNOPSIS
        Connects to Microsoft Teams. Uses interactive auth in user mode,
        non-interactive in headless mode.
    .PARAMETER TenantId
        Optional Azure AD tenant ID.
    .PARAMETER Headless
        When set, uses non-interactive authentication. Three options are available
        (certificate, client secret, managed identity) -- see inline comments in the
        function body. MSops must review and activate one before headless auth is functional.
    #>
    [CmdletBinding()]
    param(
        [string]$TenantId,
        [switch]$Headless
        # -- NON-INTERACTIVE AUTH PARAMS -- uncomment params for the chosen option --
        # [string]$ApplicationId,
        # [string]$CertificateThumbprint,   # Option A
        # [SecureString]$ClientSecret       # Option B (mutually exclusive with CertificateThumbprint)
    )

    $connectParams = @{}
    if ($TenantId) {
        $connectParams['TenantId'] = $TenantId
    }

    try {
        if ($Headless) {
            # =========================================================================
            # NON-INTERACTIVE AUTH OPTIONS -- MSOPS REVIEW REQUIRED
            # Activate exactly ONE option by uncommenting its block.
            # Delete or comment out the FALLBACK line at the bottom when done.
            # =========================================================================

            # -- OPTION A: Service Principal + Certificate (RECOMMENDED) -------------
            # Why recommended: certificate credentials do not expire on a fixed schedule
            # and the private key never leaves the certificate store.
            #
            # Prerequisites:
            #   1. Create an App Registration in Entra ID (Azure AD).
            #   2. Assign the "Teams Administrator" role to the app in Entra ID roles.
            #      (API permissions are NOT used -- Teams CS cmdlets use role assignment.)
            #   3. Generate a certificate (self-signed is acceptable for internal use).
            #      Upload the public key (.cer) to the App Registration under Certificates.
            #   4. Install the certificate WITH the private key on the machine that runs
            #      this script (Cert:\LocalMachine\My recommended for service accounts;
            #      Cert:\CurrentUser\My for per-user installs).
            #   5. Pass -ApplicationId, -TenantId, and -CertificateThumbprint at runtime,
            #      or store them in a config file that is dot-sourced before calling.
            #
            # if ($ApplicationId -and $TenantId -and $CertificateThumbprint) {
            #     $connectParams['ApplicationId']         = $ApplicationId
            #     $connectParams['CertificateThumbprint'] = $CertificateThumbprint
            #     Connect-MicrosoftTeams @connectParams
            # } else {
            #     throw "Option A requires -ApplicationId, -TenantId, and -CertificateThumbprint."
            # }

            # -- OPTION B: Service Principal + Client Secret -------------------------
            # Use when a certificate cannot be deployed to the target machine.
            #
            # Prerequisites:
            #   1. Create an App Registration in Entra ID (Azure AD).
            #   2. Assign the "Teams Administrator" role to the app.
            #   3. Generate a client secret in the App Registration; note the expiry date.
            #   4. Store the secret securely -- do NOT embed it in this script or in
            #      source control. Recommended storage:
            #        - Azure Key Vault  (retrieve via Get-AzKeyVaultSecret at runtime)
            #        - DPAPI-encrypted file  (ConvertFrom-SecureString / Export-Clixml)
            #        - Windows Credential Manager
            #   5. Load the secret as a SecureString before calling, e.g.:
            #        $secret = Get-AzKeyVaultSecret -VaultName '...' -Name '...' -AsPlainText |
            #                  ConvertTo-SecureString -AsPlainText -Force
            #
            # if ($ApplicationId -and $TenantId -and $ClientSecret) {
            #     $connectParams['ApplicationId'] = $ApplicationId
            #     $appCred = New-Object System.Management.Automation.PSCredential(
            #         $ApplicationId, $ClientSecret
            #     )
            #     $connectParams['ApplicationCredential'] = $appCred
            #     Connect-MicrosoftTeams @connectParams
            # } else {
            #     throw "Option B requires -ApplicationId, -TenantId, and -ClientSecret (SecureString)."
            # }

            # -- OPTION C: Managed Identity (Azure-hosted environments only) ---------
            # Use when the script runs on an Azure resource (VM, App Service, Automation
            # Account, Azure Function). No credentials are stored anywhere.
            #
            # Prerequisites:
            #   1. Enable system-assigned (or user-assigned) managed identity on the
            #      Azure resource that runs this script.
            #   2. In Entra ID, assign the "Teams Administrator" role to the managed identity.
            #   3. No credential parameters needed -- Azure handles token acquisition.
            #      Remove $ApplicationId/$CertificateThumbprint/$ClientSecret from both
            #      param() blocks when using this option.
            #
            # Connect-MicrosoftTeams -Identity

            # -- FALLBACK (remove this block after activating an option above) -------
            # Currently falls back to interactive Connect-MicrosoftTeams.
            # -Headless will NOT suppress the interactive prompt until an option is activated.
            Write-Verbose "Headless mode: non-interactive auth not yet configured -- falling back to interactive."
            Connect-MicrosoftTeams @connectParams
        } else {
            Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
            Connect-MicrosoftTeams @connectParams
            Write-Host "Connected." -ForegroundColor Green
        }
    } catch {
        Write-Error "Failed to connect to Microsoft Teams: $_"
        throw
    }
}

#endregion

#region Get-TeamsIVRList

function Get-TeamsIVRList {
    <#
    .SYNOPSIS
        Retrieves all Auto Attendants. (Call Queue support is stubbed out — see comments in body.)
    .OUTPUTS
        PSCustomObject with properties: Name, Type, Identity
    #>
    [CmdletBinding()]
    param(
        [switch]$Headless
    )

    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    # --- Auto Attendants ---
    try {
        if (-not $Headless) {
            Write-Host "Retrieving Auto Attendants..." -ForegroundColor Cyan
        }

        $autoAttendants = Get-CsAutoAttendant
        $aaCount = @($autoAttendants).Count
        $i = 0

        foreach ($aa in $autoAttendants) {
            $i++
            if (-not $Headless) {
                Write-Progress -Activity "Loading Auto Attendants" `
                               -Status "$i of $aaCount : $($aa.Name)" `
                               -PercentComplete (($i / $aaCount) * 100)
            }
            $results.Add([PSCustomObject]@{
                Name     = $aa.Name
                Type     = 'AutoAttendant'
                Identity = $aa.Identity
            })
        }

        if (-not $Headless) { Write-Progress -Activity "Loading Auto Attendants" -Completed }
    } catch {
        Write-Error "Failed to retrieve Auto Attendants: $_"
        throw
    }

    # --- Call Queues ---
    # Call Queue support is stubbed out -- Get-CsCallQueue requires additional permissions
    # not always available in the tenant. Uncomment when CQ support is needed.
    <#
    try {
        if (-not $Headless) {
            Write-Host "Retrieving Call Queues..." -ForegroundColor Cyan
        }

        $callQueues = Get-CsCallQueue
        $cqCount = @($callQueues).Count
        $i = 0

        foreach ($cq in $callQueues) {
            $i++
            if (-not $Headless) {
                Write-Progress -Activity "Loading Call Queues" `
                               -Status "$i of $cqCount : $($cq.Name)" `
                               -PercentComplete (($i / $cqCount) * 100)
            }
            $results.Add([PSCustomObject]@{
                Name     = $cq.Name
                Type     = 'CallQueue'
                Identity = $cq.Identity
            })
        }

        if (-not $Headless) { Write-Progress -Activity "Loading Call Queues" -Completed }
    } catch {
        Write-Error "Failed to retrieve Call Queues: $_"
        throw
    }
    #>
    
    return $results
}

#endregion

#region Get-TeamsIVRSchedule

function Get-TeamsIVRSchedule {
    <#
    .SYNOPSIS
        Retrieves the business hours and after-hours schedule for a given Auto Attendant.
    .PARAMETER Identity
        The Identity of the Auto Attendant to query.
    .OUTPUTS
        PSCustomObject with properties: Name, Identity, BusinessHours, AfterHours
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    # Identify the AfterHours schedule via the CallHandlingAssociations collection (authoritative)
    # rather than guessing by schedule name. The AfterHours association points to the schedule
    # that defines CLOSED/after-hours time ranges; all other schedules are business-hours related.
    $ahAssoc       = @($aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq 'AfterHours' }) | Select-Object -First 1
    $afterHours    = if ($ahAssoc) { @($aa.Schedules | Where-Object { $_.Id -eq $ahAssoc.ScheduleId }) | Select-Object -First 1 } else { $null }
    $businessHours = if ($ahAssoc) { @($aa.Schedules | Where-Object { $_.Id -ne $ahAssoc.ScheduleId }) } else { @($aa.Schedules) }

    return [PSCustomObject]@{
        Name          = $aa.Name
        Identity      = $aa.Identity
        BusinessHours = $businessHours
        AfterHours    = $afterHours
    }
}

#endregion

#region Format-TeamsIVRSchedule

function Format-TeamsIVRSchedule {
    <#
    .SYNOPSIS
        Decodes a CsOnlineSchedule object into human-readable day/time strings.
    .DESCRIPTION
        Converts a WeeklyRecurrentSchedule's TimeRange entries (stored as TimeSpan)
        into lines like "Monday   : 9:00am - 5:00pm".
        Days with no hours configured are shown as "Closed".

        Use -InvertToOpenHours when decoding an AfterHours schedule (which stores CLOSED
        periods) so the output shows the open/business hours instead of the closed windows.
        This is required for the round-trip workflow: DecodeSchedule output should match
        the input format that Build-TeamsIVRSchedule accepts.
    .PARAMETER Schedule
        A single CsOnlineSchedule object (e.g. one entry from $aa.Schedules).
    .PARAMETER InvertToOpenHours
        When set, treats the schedule's time ranges as closed/after-hours periods and
        displays the complement (open hours) instead. Use this for AfterHours schedules.
    .OUTPUTS
        String[] — one line per day of the week.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Schedule,

        [switch]$InvertToOpenHours
    )

    # Converts a TimeSpan (e.g. 09:00:00) to 12-hour time (e.g. "9:00am")
    function ConvertTo-12Hour {
        param([TimeSpan]$TimeSpan)
        try {
            return [DateTime]::Today.Add($TimeSpan).ToString('h:mmtt').ToLower()
        } catch {
            return $TimeSpan.ToString()
        }
    }

    # Given a list of CLOSED time ranges for one day, returns the complement (open periods).
    # Mirrors the logic in Get-AfterHoursRanges inside Build-TeamsIVRSchedule.
    #   Closed [12am-9am, 5pm-12am] -> Open [9am-5pm]
    #   Closed [12am-12am]          -> Open []  (displayed as "Closed")
    #   Closed []                   -> Open [12am-12am]  (open all day)
    function Get-OpenHoursFromClosed {
        param([object[]]$ClosedRanges)
        $dayStart = [TimeSpan]::Zero
        $dayEnd   = [TimeSpan]::FromHours(24)
        if ($ClosedRanges.Count -eq 0) {
            return @([PSCustomObject]@{ Start = $dayStart; End = $dayEnd })
        }
        $result = [System.Collections.Generic.List[object]]::new()
        $cursor = $dayStart
        foreach ($range in ($ClosedRanges | Sort-Object { $_.Start })) {
            if ($cursor -lt $range.Start) {
                $result.Add([PSCustomObject]@{ Start = $cursor; End = $range.Start })
            }
            $cursor = $range.End
        }
        if ($cursor -ne $dayStart -and $cursor -lt $dayEnd) {
            $result.Add([PSCustomObject]@{ Start = $cursor; End = $dayEnd })
        }
        return @($result)
    }

    if ($null -eq $Schedule.WeeklyRecurrentSchedule) {
        return @('  (No weekly recurring schedule — may be a fixed/date-range schedule)')
    }

    $wrs = $Schedule.WeeklyRecurrentSchedule

    $days = @(
        [PSCustomObject]@{ Label = 'Monday';    Hours = $wrs.MondayHours    }
        [PSCustomObject]@{ Label = 'Tuesday';   Hours = $wrs.TuesdayHours   }
        [PSCustomObject]@{ Label = 'Wednesday'; Hours = $wrs.WednesdayHours }
        [PSCustomObject]@{ Label = 'Thursday';  Hours = $wrs.ThursdayHours  }
        [PSCustomObject]@{ Label = 'Friday';    Hours = $wrs.FridayHours    }
        [PSCustomObject]@{ Label = 'Saturday';  Hours = $wrs.SaturdayHours  }
        [PSCustomObject]@{ Label = 'Sunday';    Hours = $wrs.SundayHours    }
    )

    $lines = foreach ($day in $days) {
        $hoursList = @($day.Hours)
        if ($InvertToOpenHours) {
            $hoursList = @(Get-OpenHoursFromClosed $hoursList)
        }
        if ($hoursList.Count -eq 0) {
            "  $($day.Label.PadRight(11)): Closed"
        } else {
            $ranges = $hoursList | ForEach-Object {
                "$(ConvertTo-12Hour $_.Start) - $(ConvertTo-12Hour $_.End)"
            }
            "  $($day.Label.PadRight(11)): $($ranges -join ', ')"
        }
    }

    return $lines
}

#endregion

#region Build-TeamsIVRSchedule

function Build-TeamsIVRSchedule {
    <#
    .SYNOPSIS
        Assembles a CsOnlineSchedule object from human-readable day/time input.
    .DESCRIPTION
        Accepts a hashtable of day names to time-range strings — matching the output
        format of Format-TeamsIVRSchedule — and returns a CsOnlineSchedule object
        ready to pass directly to Set-TeamsIVRSchedule -Schedule.

        Input format per day:
          Single range : '9:00am - 5:00pm'
          Multi-range  : '9:00am - 12:00pm, 1:00pm - 5:00pm'   (e.g. lunch break)
          Closed       : 'Closed'  (or omit the day key entirely)

    .PARAMETER Name
        Name label for the schedule object (e.g. 'BusinessHours').

    .PARAMETER Days
        Hashtable mapping day names to time-range strings.
        Valid keys: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday

        Example:
            @{
                Monday    = '9:00am - 5:00pm'
                Tuesday   = '9:00am - 5:00pm'
                Wednesday = '9:00am - 12:00pm, 1:00pm - 5:00pm'
                Thursday  = '9:00am - 5:00pm'
                Friday    = '9:00am - 5:00pm'
                Saturday  = 'Closed'
                Sunday    = 'Closed'
            }

    .OUTPUTS
        CsOnlineSchedule — pass directly to Set-TeamsIVRSchedule -Schedule.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [hashtable]$Days
    )

    $validDays = @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday')

    # Converts a 12-hour time string like "9:00am" or "5:00pm" to a TimeSpan
    function ConvertFrom-12Hour {
        param([string]$TimeStr)
        try {
            $formats = [string[]]@('h:mmtt', 'hh:mmtt', 'h:mm tt', 'hh:mm tt')
            $dt = [DateTime]::ParseExact(
                $TimeStr.Trim().ToUpper(),
                $formats,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None
            )
            return $dt.TimeOfDay
        } catch {
            throw "Could not parse time '$TimeStr'. Expected format: '9:00am' or '5:00pm'"
        }
    }

    # Parses a single range string like "9:00am - 5:00pm" into a CsOnlineTimeRange
    function ConvertTo-TimeRange {
        param([string]$RangeStr)
        $parts = $RangeStr -split '\s*-\s*', 2
        if ($parts.Count -ne 2) {
            throw "Invalid time range '$RangeStr'. Expected format: '9:00am - 5:00pm'"
        }
        try {
            $start = ConvertFrom-12Hour $parts[0]
            $end   = ConvertFrom-12Hour $parts[1]
            if ($start -ge $end) {
                throw "Start time '$($parts[0].Trim())' must be earlier than end time '$($parts[1].Trim())'"
            }
            return New-CsOnlineTimeRange -Start $start -End $end
        } catch {
            throw "Failed to build time range from '$RangeStr': $_"
        }
    }

    # Parses a full day value -- "Closed", a single range, or comma-separated ranges
    function Get-DayRanges {
        param([string]$DayValue)
        if ([string]::IsNullOrWhiteSpace($DayValue) -or $DayValue.Trim() -ieq 'Closed') {
            return @()
        }
        try {
            return @($DayValue -split ',' | ForEach-Object { ConvertTo-TimeRange $_.Trim() })
        } catch {
            throw $_
        }
    }

    # Inverts open-hours ranges into closed/after-hours ranges.
    # Teams stores the schedule on the AfterHours association as WHEN THE AA IS CLOSED,
    # so we must submit the complement of the business hours the user provides.
    #   Open 9am-5pm  ->  Closed 12am-9am + 5pm-12am
    #   Closed all day -> Closed 12am-12am (full-day sentinel Teams recognises)
    function Get-AfterHoursRanges {
        param([object[]]$OpenRanges)

        $dayStart = [TimeSpan]::Zero           # 00:00:00 -- midnight, start of day
        $dayEnd   = [TimeSpan]::FromHours(24)  # 24:00:00 -- midnight, end of day
                                               # Teams requires Start < End, so end-of-day
                                               # must be 24h, not 00:00 (which is < any start)

        if ($OpenRanges.Count -eq 0) {
            # Closed all day: full 24-hour closed range
            return @(New-CsOnlineTimeRange -Start $dayStart -End $dayEnd)
        }

        $result = [System.Collections.Generic.List[object]]::new()
        $cursor = $dayStart

        foreach ($range in ($OpenRanges | Sort-Object { $_.Start })) {
            if ($cursor -lt $range.Start) {
                $result.Add((New-CsOnlineTimeRange -Start $cursor -End $range.Start))
            }
            $cursor = $range.End
        }

        # Final segment from last open period's end to end of day.
        # Skip if cursor is back at dayStart (00:00) -- means the last open range ended
        # at midnight (Teams stores that as 00:00), so the full day is already covered.
        if ($cursor -ne $dayStart -and $cursor -lt $dayEnd) {
            $result.Add((New-CsOnlineTimeRange -Start $cursor -End $dayEnd))
        }

        return @($result)
    }

    # Validate all provided keys before building anything
    foreach ($key in $Days.Keys) {
        if ($key -notin $validDays) {
            throw "Invalid day key '$key'. Valid keys: $($validDays -join ', ')"
        }
    }

    $scheduleParams = @{
        Name                    = $Name
        WeeklyRecurrentSchedule = $true
    }

    foreach ($day in $validDays) {
        # Omitted days are treated the same as 'Closed' (after-hours all day)
        $dayValue = if ($Days.ContainsKey($day)) { $Days[$day] } else { 'Closed' }
        try {
            $openRanges = @(Get-DayRanges $dayValue)
        } catch {
            throw "Error processing $day : $_"
        }
        try {
            $closedRanges = @(Get-AfterHoursRanges $openRanges)
        } catch {
            throw "Error computing after-hours for $day : $_"
        }
        if ($closedRanges.Count -gt 0) {
            $scheduleParams["${day}Hours"] = $closedRanges
        }
    }

    try {
        $schedule = New-CsOnlineSchedule @scheduleParams
        Write-Verbose "Schedule '$Name' assembled successfully."
        return $schedule
    } catch {
        throw "Failed to create CsOnlineSchedule '$Name': $_"
    }
}

#endregion

#region Set-TeamsIVRSchedule

function Set-TeamsIVRSchedule {
    <#
    .SYNOPSIS
        Updates the AfterHours schedule for a given Auto Attendant, which controls business hours.
    .PARAMETER Identity
        The Identity of the Auto Attendant to update.
    .PARAMETER Schedule
        A CsOnlineSchedule object containing after-hours (closed period) time ranges.
        Build one with Build-TeamsIVRSchedule, which accepts open hours and inverts them automatically.
    .NOTES
        Teams AfterHours schedule semantics: the schedule passed to this function must
        represent CLOSED/after-hours time ranges (not open hours). Use Build-TeamsIVRSchedule
        to convert human-readable open-hours input into the correct closed-period format.
        The Schedules and CallHandlingAssociations collections on the AA object are fixed-size
        arrays -- .Add()/.Remove() will throw. This function rebuilds them via filtering + reassign.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [object]$Schedule
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Update IVR schedule")) {

        # In Teams, the AfterHours schedule defines WHEN the AA is CLOSED (after-hours periods).
        # Calls that arrive during those time ranges are routed to the AfterHours call flow.
        # All other times use the DefaultCallFlow (business hours). Updating this schedule
        # changes which hours are treated as after-hours.
        $ahAssoc = @($aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq 'AfterHours' }) | Select-Object -First 1

        if (-not $ahAssoc) {
            throw "No AfterHours CallHandlingAssociation found on '$($aa.Name)'. Cannot determine which schedule to update."
        }

        $existingSchedule = @($aa.Schedules | Where-Object { $_.Id -eq $ahAssoc.ScheduleId }) | Select-Object -First 1

        if (-not $existingSchedule) {
            throw "AfterHours association references schedule ID '$($ahAssoc.ScheduleId)' but that schedule was not found on '$($aa.Name)'."
        }

        # The Schedules/CallHandlingAssociations collections are fixed-size arrays -- .Remove()/.Add()
        # throw. Instead, build new arrays via filtering and reassign the properties.
        $callFlowId = $ahAssoc.CallFlowId

        $newAssoc = New-CsAutoAttendantCallHandlingAssociation `
            -Type AfterHours `
            -ScheduleId $Schedule.Id `
            -CallFlowId $callFlowId

        $aa.Schedules = @($aa.Schedules | Where-Object { $_.Id -ne $existingSchedule.Id }) + @($Schedule)
        $aa.CallHandlingAssociations = @($aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ahAssoc.ScheduleId }) + @($newAssoc)

        try {
            Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference
            Write-Verbose "Schedule updated for '$($aa.Name)'"
        } catch {
            Write-Error "Failed to update schedule for '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Get-TeamsIVRAuthorizedUsers

function Get-TeamsIVRAuthorizedUsers {
    <#
    .SYNOPSIS
        Returns the list of users authorized to manage a given Auto Attendant.
    .DESCRIPTION
        Fetches the Auto Attendant's AuthorizedUsers property (stored as object ID GUIDs)
        and resolves each to a friendly user object via Get-CsOnlineUser.
        If a user cannot be resolved (e.g. deleted or unlicensed), their raw object ID
        is still included in the output so the list is always complete.
    .PARAMETER Identity
        Identity of the Auto Attendant to query.
    .OUTPUTS
        PSCustomObject[] with properties: DisplayName, UserPrincipalName, ObjectId
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $authorizedIds = @($aa.AuthorizedUsers)

    if ($authorizedIds.Count -eq 0) {
        Write-Verbose "No authorized users configured for '$($aa.Name)'"
        return @()
    }

    $users = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($objectId in $authorizedIds) {
        try {
            $resolved = Get-CsOnlineUser -Identity $objectId
            $users.Add([PSCustomObject]@{
                DisplayName       = $resolved.DisplayName
                UserPrincipalName = $resolved.UserPrincipalName
                ObjectId          = $objectId
            })
        } catch {
            # User may be deleted or unlicensed — include raw ID so list stays complete
            Write-Warning "Could not resolve user object ID '$objectId': $_"
            $users.Add([PSCustomObject]@{
                DisplayName       = '(unresolved)'
                UserPrincipalName = '(unresolved)'
                ObjectId          = $objectId
            })
        }
    }

    return $users
}

#endregion

#region Resolve-TeamsUser (private helper)

function Resolve-TeamsUser {
    <#
    .SYNOPSIS
        Resolves a user identity string to a CsOnlineUser object.
    .DESCRIPTION
        Tries a direct Get-CsOnlineUser lookup first (works for UPN or object ID GUID).
        Falls back to a DisplayName filter if the direct lookup fails.
        Throws a descriptive error if the user cannot be found or if a display name
        matches more than one account.
    .PARAMETER User
        UPN, object ID GUID, or display name (must be unique in the tenant).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$User
    )

    try {
        return Get-CsOnlineUser -Identity $User
    } catch {
        try {
            $nameMatches = @(Get-CsOnlineUser -Filter "DisplayName -eq '$User'")
            if ($nameMatches.Count -eq 0) {
                throw "No user found matching '$User'."
            }
            if ($nameMatches.Count -gt 1) {
                throw "Multiple users matched display name '$User'. Use UPN or object ID to be specific."
            }
            return $nameMatches[0]
        } catch {
            throw "Failed to resolve user '$User': $_"
        }
    }
}

#endregion

#region Test-TeamsIVRUserAccess

function Test-TeamsIVRUserAccess {
    <#
    .SYNOPSIS
        Checks whether a given user is authorized to manage a specific Auto Attendant.
    .DESCRIPTION
        Resolves the provided user identity against Teams (accepts UPN, display name,
        or object ID GUID), then checks their object ID against the IVR's AuthorizedUsers list.
        Display name lookups must be unique — if multiple users share a name, use UPN or object ID.
    .PARAMETER Identity
        Identity of the Auto Attendant to check against.
    .PARAMETER User
        The user to check. Accepts:
          - UPN           : user@contoso.com
          - Object ID     : GUID (e.g. from Get-TeamsIVRAuthorizedUsers output)
          - Display name  : "Jane Smith"  (must be unique in the tenant)
    .OUTPUTS
        Boolean -- $true if the user is authorized, $false if not.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [string]$User
    )

    $resolvedUser = Resolve-TeamsUser -User $User

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $isAuthorized = @($aa.AuthorizedUsers) -contains $resolvedUser.Identity

    Write-Verbose "Access check: '$($resolvedUser.UserPrincipalName)' authorized for '$Identity': $isAuthorized"
    return $isAuthorized
}

#endregion

#region Get-TeamsIVRsByUser

function Get-TeamsIVRsByUser {
    <#
    .SYNOPSIS
        Returns all Auto Attendants where the given user is in the authorized users list.
    .DESCRIPTION
        Resolves the provided user identity against Teams (accepts UPN, display name,
        or object ID GUID), then scans every Auto Attendant and returns the ones where
        the resolved user appears in AuthorizedUsers.
    .PARAMETER User
        The user to look up. Accepts:
          - UPN           : user@contoso.com
          - Object ID     : GUID (e.g. from Get-TeamsIVRAuthorizedUsers output)
          - Display name  : "Jane Smith"  (must be unique in the tenant)
    .PARAMETER Headless
        Suppress progress bar output.
    .OUTPUTS
        PSCustomObject[] with properties: Name, Identity
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$User,

        [switch]$Headless
    )

    $resolvedUser = Resolve-TeamsUser -User $User

    if (-not $Headless) {
        Write-Host "Scanning Auto Attendants for user '$($resolvedUser.UserPrincipalName)'..." -ForegroundColor Cyan
    }

    try {
        $allAAs = @(Get-CsAutoAttendant)
    } catch {
        Write-Error "Failed to retrieve Auto Attendants: $_"
        throw
    }

    $aaCount = $allAAs.Count
    $i = 0
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($aa in $allAAs) {
        $i++
        if (-not $Headless) {
            Write-Progress -Activity "Scanning Auto Attendants" `
                           -Status "$i of $aaCount : $($aa.Name)" `
                           -PercentComplete (($i / $aaCount) * 100)
        }
        if (@($aa.AuthorizedUsers) -contains $resolvedUser.Identity) {
            $results.Add([PSCustomObject]@{
                Name     = $aa.Name
                Identity = $aa.Identity
            })
        }
    }

    if (-not $Headless) { Write-Progress -Activity "Scanning Auto Attendants" -Completed }

    return $results
}

#endregion

#region Show-Usage

function Show-Usage {
    <#
    .SYNOPSIS
        Prints a usage reference for all CLI actions and helper functions.
        Called automatically when the script is run with no arguments.
    #>

    Write-Host ""
    Write-Host "  Teams IVR Manager  v$($Script:Version)" -ForegroundColor Cyan
    Write-Host "  Usage: .\Manage-TeamsIVR.ps1 [action] [options]" -ForegroundColor Cyan

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  CLI ACTIONS" -ForegroundColor Yellow
    Write-Host "  -----------"
    Write-Host "    (none) | -Help"
    Write-Host "        Show this reference. No Teams connection required."
    Write-Host ""
    Write-Host "    -GetAllIVRs"
    Write-Host "        List every Auto Attendant and Call Queue in the tenant."
    Write-Host ""
    Write-Host "    -GetIVR -Name <string>"
    Write-Host "        Find a specific IVR by name (partial/wildcard match)."
    Write-Host ""
    Write-Host "    -GetSchedule -Identity <string>"
    Write-Host "        Retrieve the raw schedule object for an Auto Attendant."
    Write-Host "        Use -DecodeSchedule for a human-readable version."
    Write-Host ""
    Write-Host "    -DecodeSchedule -Identity <string>"
    Write-Host "        Retrieve and display an IVR schedule in human-readable format:"
    Write-Host "          Monday     : 9:00am - 5:00pm"
    Write-Host "          Saturday   : Closed"
    Write-Host "        The output format matches the input format for Build-TeamsIVRSchedule."
    Write-Host ""
    Write-Host "    -SetSchedule -Identity <string> -Schedule <CsOnlineSchedule>"
    Write-Host "        Apply a schedule object to an Auto Attendant."
    Write-Host "        Build the schedule object first with Build-TeamsIVRSchedule (see below)."
    Write-Host "        NOTE: Changes are accepted by the Teams backend immediately, but the"
    Write-Host "        Teams Admin Center portal may take 1-15 minutes to reflect the update."
    Write-Host "        Use -DecodeSchedule to verify the change right away instead of the portal."
    Write-Host ""
    Write-Host "    -GetAuthorizedUsers -Identity <string>"
    Write-Host "        List all users authorized to manage an Auto Attendant."
    Write-Host "        Output: DisplayName, UserPrincipalName, ObjectId per user."
    Write-Host ""
    Write-Host "    -CheckUserAccess -Identity <string> -User <string>"
    Write-Host "        Check whether a specific user is authorized for an Auto Attendant."
    Write-Host "        Outputs AUTHORIZED or NOT AUTHORIZED."
    Write-Host "        -User accepts UPN, display name (unique), or object ID GUID."
    Write-Host ""
    Write-Host "    -GetUserIVRs -User <string>"
    Write-Host "        List all Auto Attendants where the given user is an authorized user."
    Write-Host "        Output: Name and Identity of each matching Auto Attendant."
    Write-Host "        -User accepts UPN, display name (unique), or object ID GUID."

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  OPTIONS" -ForegroundColor Yellow
    Write-Host "  -------"
    Write-Host "    -TenantId <string>   Azure AD tenant ID (optional - uses default if omitted)"
    Write-Host "    -Headless            Suppress banner/progress bars; use non-interactive auth"

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  HELPER FUNCTIONS  (available after dot-sourcing this script)" -ForegroundColor Yellow
    Write-Host "  ---------------------------------------------------------------"
    Write-Host "  Dot-source to load all functions into your session without running the script:"
    Write-Host "    . .\Manage-TeamsIVR.ps1"
    Write-Host ""
    Write-Host "    Get-TeamsIVRList"
    Write-Host "        Returns all Auto Attendants as objects (Name, Type, Identity)."
    Write-Host "        Use this for programmatic filtering -- unlike -GetAllIVRs, it does not pipe through Format-Table."
    Write-Host ""
    Write-Host "    Get-TeamsIVRSchedule -Identity <string>"
    Write-Host "        Returns the raw schedule object for an AA (AfterHours + any other schedules)."
    Write-Host "        Use Format-TeamsIVRSchedule -InvertToOpenHours on the AfterHours property to display open hours."
    Write-Host ""
    Write-Host "    Build-TeamsIVRSchedule -Name <string> -Days <hashtable>"
    Write-Host "        Converts human-readable day/time input into a CsOnlineSchedule object"
    Write-Host "        ready to pass to -SetSchedule. Day keys: Monday..Sunday."
    Write-Host "        Values: '9:00am - 5:00pm'  |  'Closed'  |  '9:00am - 12:00pm, 1:00pm - 5:00pm'"
    Write-Host ""
    Write-Host "    Format-TeamsIVRSchedule -Schedule <CsOnlineSchedule> [-InvertToOpenHours]"
    Write-Host "        Decodes a raw CsOnlineSchedule object into human-readable day/time lines."
    Write-Host "        Pass -InvertToOpenHours when decoding an AfterHours schedule to show"
    Write-Host "        open/business hours instead of the raw closed periods."
    Write-Host "        Called internally by -DecodeSchedule; also usable directly."
    Write-Host ""
    Write-Host "    Get-TeamsIVRAuthorizedUsers -Identity <string>"
    Write-Host "        Returns the authorized users list for an IVR as objects (DisplayName, UPN, ObjectId)."
    Write-Host "        Called internally by -GetAuthorizedUsers and -CheckUserAccess; also usable directly."
    Write-Host ""
    Write-Host "    Test-TeamsIVRUserAccess -Identity <string> -User <string>"
    Write-Host "        Resolves the user from Teams and returns `$true if authorized, `$false if not."
    Write-Host "        -User accepts UPN, display name (must be unique), or object ID GUID."
    Write-Host "        Called internally by -CheckUserAccess; also usable directly in scripts."
    Write-Host ""
    Write-Host "    Get-TeamsIVRsByUser -User <string>"
    Write-Host "        Returns all Auto Attendants where the given user is authorized, as objects (Name, Identity)."
    Write-Host "        Called internally by -GetUserIVRs; also usable directly in scripts."

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  CHAINING OUTPUTS  (PowerShell equivalent of bash pipes)" -ForegroundColor Yellow
    Write-Host "  ----------------------------------------------------------"
    Write-Host "  In bash you pass output between commands with |  e.g.:"
    Write-Host "    cmd1 | grep Sales | awk '{print `$2}'"
    Write-Host ""
    Write-Host "  PowerShell's | pipe works the same way but passes objects, not text."
    Write-Host "  The cleanest pattern: store results in a variable, then access properties."
    Write-Host ""
    Write-Host "  Extract one value:    `$id = (SomeCommand | Where-Object { ... }).Property"
    Write-Host "  Chain several steps:  `$a = cmd1 ; `$b = `$a | Where-Object { ... } ; cmd2 `$b.Identity"
    Write-Host ""
    Write-Host "  IMPORTANT: -GetAllIVRs and -GetIVR pipe through Format-Table for display,"
    Write-Host "  so capturing their output gives formatted strings, not objects."
    Write-Host "  For programmatic chaining, dot-source the script and call Get-TeamsIVRList"
    Write-Host "  directly -- it returns plain objects you can filter with Where-Object."
    Write-Host "    . .\Manage-TeamsIVR.ps1    # loads all functions into your session"
    Write-Host "    `$id = (Get-TeamsIVRList | Where-Object { `$_.Name -like `"*Sales*`" } | Select-Object -First 1).Identity"
    Write-Host ""

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  TYPICAL WORKFLOW: FROM 'WHAT IVRs EXIST?' TO SETTING A SCHEDULE" -ForegroundColor Yellow
    Write-Host "  -------------------------------------------------------------------"
    Write-Host ""
    Write-Host "  ---- STEP 0 : Discover IVRs and capture the Identity ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  List every IVR in the tenant:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -GetAllIVRs"
    Write-Host "    # Output:"
    Write-Host "    #   Name                  Type           Identity"
    Write-Host "    #   ----                  ----           --------"
    Write-Host "    #   Sales Support AA      AutoAttendant  xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    Write-Host "    #   IT Help Desk          AutoAttendant  yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy"
    Write-Host ""
    Write-Host "  Filter by partial name if the list is long:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -GetIVR -Name `"Sales`""
    Write-Host ""
    Write-Host "  Copy the Identity GUID from the output and store it -- every step below reuses it:"
    Write-Host "    `$id = `"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`"   # paste your GUID here"
    Write-Host ""
    Write-Host "  Or capture it programmatically (no copy+paste needed):"
    Write-Host "    . .\Manage-TeamsIVR.ps1"
    Write-Host "    `$id = (Get-TeamsIVRList | Where-Object { `$_.Name -like `"*Sales*`" } | Select-Object -First 1).Identity"
    Write-Host "    `$id   # verify the GUID printed correctly before continuing"
    Write-Host ""
    Write-Host "  ---- STEP 1 : (Optional) Confirm you are authorized to change this IVR ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -CheckUserAccess -Identity `$id -User `"you@contoso.com`""
    Write-Host "    # Output: AUTHORIZED  or  NOT AUTHORIZED"
    Write-Host ""
    Write-Host "  ---- STEP 2 : View the current schedule ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -DecodeSchedule -Identity `$id"
    Write-Host "    # Output:"
    Write-Host "    #   Monday     : 9:00am - 5:00pm"
    Write-Host "    #   Tuesday    : 9:00am - 5:00pm"
    Write-Host "    #   Wednesday  : 9:00am - 5:00pm"
    Write-Host "    #   Thursday   : 9:00am - 5:00pm"
    Write-Host "    #   Friday     : 9:00am - 5:00pm"
    Write-Host "    #   Saturday   : Closed"
    Write-Host "    #   Sunday     : Closed"
    Write-Host ""
    Write-Host "  ---- STEP 3 : Dot-source, then build the new schedule ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Dot-source loads all functions into your session (required for Build-TeamsIVRSchedule):"
    Write-Host "    . .\Manage-TeamsIVR.ps1"
    Write-Host ""
    Write-Host "  Build the schedule object -- copy+paste and edit only the days/times you want to change."
    Write-Host "  Entries are separated by semicolons so this works pasted as one line OR as a multi-line block:"
    Write-Host "    `$newSchedule = Build-TeamsIVRSchedule -Name 'BusinessHours' -Days @{"
    Write-Host "        Monday    = '9:00am - 5:00pm';"
    Write-Host "        Tuesday   = '9:00am - 5:00pm';"
    Write-Host "        Wednesday = '9:00am - 12:00pm, 1:00pm - 5:00pm';"
    Write-Host "        Thursday  = '9:00am - 5:00pm';"
    Write-Host "        Friday    = '9:00am - 3:00pm';"
    Write-Host "        Saturday  = 'Closed';"
    Write-Host "        Sunday    = 'Closed'"
    Write-Host "    }"
    Write-Host ""
    Write-Host "  ---- STEP 4 : Apply the new schedule ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  `$id must still hold your GUID (see Step 0). If you opened a new terminal, re-assign it."
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetSchedule -Identity `$id -Schedule `$newSchedule"
    Write-Host ""
    Write-Host "  ---- STEP 5 : Verify the change ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -DecodeSchedule -Identity `$id"
    Write-Host ""
}

#endregion

#region Main

$Script:IsDotSourced = $MyInvocation.InvocationName -eq '.'

if (-not $Headless -and -not $Script:IsDotSourced) {
    Show-Banner
}

switch ($PSCmdlet.ParameterSetName) {

    'ShowUsage' {
        # When dot-sourced: load functions silently; caller uses Get-Help for documentation
        # When run directly with no args: print the usage reference
        # NOTE: exit would terminate the shell session if dot-sourced, so use return
        if (-not $Script:IsDotSourced) {
            Show-Usage
        }
        return
    }

    default {
        try {
            Connect-TeamsEnvironment -TenantId $TenantId -Headless:$Headless
            # When Option A or B is activated, replace the call above with:
            # Connect-TeamsEnvironment -TenantId $TenantId -Headless:$Headless `
            #     -ApplicationId $ApplicationId `
            #     -CertificateThumbprint $CertificateThumbprint   # Option A
            #
            # Connect-TeamsEnvironment -TenantId $TenantId -Headless:$Headless `
            #     -ApplicationId $ApplicationId `
            #     -ClientSecret $ClientSecret                     # Option B
            #
            # Option C requires no extra params:
            # Connect-TeamsEnvironment -TenantId $TenantId -Headless:$Headless
        } catch {
            exit 1
        }

        switch ($PSCmdlet.ParameterSetName) {

            'GetAllIVRs' {
                # Retrieve every Auto Attendant and Call Queue
                try {
                    $ivrs = Get-TeamsIVRList -Headless:$Headless
                    $ivrs | Format-Table -AutoSize
                } catch {
                    Write-Error "Failed to retrieve IVR list: $_"
                    exit 1
                }
            }

            'GetIVR' {
                # Retrieve a specific IVR by name (partial match supported)
                try {
                    $ivrs = Get-TeamsIVRList -Headless:$Headless
                    $match = $ivrs | Where-Object { $_.Name -like "*$Name*" }

                    if (-not $match) {
                        Write-Warning "No IVR found matching name: '$Name'"
                    } else {
                        $match | Format-Table -AutoSize
                    }
                } catch {
                    Write-Error "Failed to retrieve IVR '$Name': $_"
                    exit 1
                }
            }

            'GetSchedule' {
                # Retrieve raw schedule objects for an Auto Attendant
                try {
                    $scheduleData = Get-TeamsIVRSchedule -Identity $Identity
                    Write-Host "`nSchedule for: $($scheduleData.Name)" -ForegroundColor Cyan
                    Write-Host "--- After-Hours Schedule (closed periods stored by Teams) ---" -ForegroundColor Yellow
                    $scheduleData.AfterHours | Format-List
                } catch {
                    Write-Error "Failed to retrieve schedule for '$Identity': $_"
                    exit 1
                }
            }

            'DecodeSchedule' {
                # Decode the schedule into human-readable open/business hours.
                # The AfterHours schedule stores closed periods — -InvertToOpenHours flips
                # them back so the output matches the Build-TeamsIVRSchedule input format.
                try {
                    $scheduleData = Get-TeamsIVRSchedule -Identity $Identity
                    Write-Host "`n  Schedule for: $($scheduleData.Name)" -ForegroundColor Cyan

                    if ($scheduleData.AfterHours) {
                        Write-Host "`n  Business Hours:" -ForegroundColor Yellow
                        Write-Host "  [$($scheduleData.AfterHours.Name)]" -ForegroundColor Gray
                        Format-TeamsIVRSchedule -Schedule $scheduleData.AfterHours -InvertToOpenHours |
                            ForEach-Object { Write-Host $_ }
                    } else {
                        Write-Host "`n  No schedule configured for this Auto Attendant." -ForegroundColor Yellow
                    }

                    # Show any non-AfterHours schedules (uncommon; present on some AAs)
                    if (@($scheduleData.BusinessHours).Count -gt 0) {
                        Write-Host "`n  Other Schedules (raw):" -ForegroundColor Yellow
                        foreach ($s in @($scheduleData.BusinessHours)) {
                            Write-Host "  [$($s.Name)]" -ForegroundColor Gray
                            Format-TeamsIVRSchedule -Schedule $s | ForEach-Object { Write-Host $_ }
                        }
                    }

                    Write-Host ""
                } catch {
                    Write-Error "Failed to decode schedule for '$Identity': $_"
                    exit 1
                }
            }

            'SetSchedule' {
                # Update the schedule for an Auto Attendant
                try {
                    Set-TeamsIVRSchedule -Identity $Identity -Schedule $Schedule
                    Write-Host "Schedule updated successfully for Identity: $Identity" -ForegroundColor Green
                } catch {
                    Write-Error "Failed to set schedule for '$Identity': $_"
                    exit 1
                }
            }

            'GetAuthorizedUsers' {
                # List all authorized users for an Auto Attendant
                try {
                    $authUsers = Get-TeamsIVRAuthorizedUsers -Identity $Identity
                    if (@($authUsers).Count -eq 0) {
                        Write-Host "No authorized users configured for Identity: $Identity" -ForegroundColor Yellow
                    } else {
                        Write-Host "`n  Authorized users for Identity: $Identity" -ForegroundColor Cyan
                        $authUsers | Format-Table DisplayName, UserPrincipalName, ObjectId -AutoSize
                    }
                } catch {
                    Write-Error "Failed to retrieve authorized users for '$Identity': $_"
                    exit 1
                }
            }

            'CheckUserAccess' {
                # Check whether a specific user is authorized for an Auto Attendant
                try {
                    $isAuthorized = Test-TeamsIVRUserAccess -Identity $Identity -User $User
                    if ($isAuthorized) {
                        Write-Host "AUTHORIZED   - '$User' is in the authorized users list for Identity: $Identity" -ForegroundColor Green
                    } else {
                        Write-Host "NOT AUTHORIZED - '$User' is not in the authorized users list for Identity: $Identity" -ForegroundColor Red
                    }
                    # Return the boolean so callers can capture it (e.g. $result = .\script.ps1 -CheckUserAccess ...)
                    return $isAuthorized
                } catch {
                    Write-Error "Failed to check user access for '$User' on '$Identity': $_"
                    exit 1
                }
            }

            'GetUserIVRs' {
                # List all Auto Attendants where the given user is an authorized user
                try {
                    $ivrs = Get-TeamsIVRsByUser -User $User -Headless:$Headless
                    if (@($ivrs).Count -eq 0) {
                        Write-Host "`n  '$User' is not an authorized user for any Auto Attendant." -ForegroundColor Yellow
                    } else {
                        Write-Host "`n  Auto Attendants where '$User' is authorized:" -ForegroundColor Cyan
                        $ivrs | Format-Table Name, Identity -AutoSize
                    }
                } catch {
                    Write-Error "Failed to retrieve IVRs for user '$User': $_"
                    exit 1
                }
            }
        }
    }
}

#endregion
