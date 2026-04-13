#Requires -Modules MicrosoftTeams
# CODING NOTE: Do NOT use em dashes or any multi-byte Unicode characters inside double-quoted
# strings or here-strings. Windows PowerShell 5.x reads files as Windows-1252 when no BOM is
# present, and byte 0x94 (the 3rd byte of the UTF-8 em dash E2 80 94) is RIGHT DOUBLE QUOTATION
# MARK in Windows-1252, which terminates string literals early and causes cascading parse errors.
# Use double-hyphen (--) instead of em dash in strings and Write-Host calls.

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
    - Get-TeamsIVRByUser            : Return all AAs where a given user appears in the AuthorizedUsers list.
    - Test-TeamsIVRUserAccess       : Check whether a given user is in the authorized users list
                                      for a specific IVR.
    - Get-TeamsIVRHolidayList       : Retrieve all holiday schedules defined in the tenant.
    - Get-TeamsIVRHoliday           : Retrieve a specific holiday schedule by name.
    - New-TeamsIVRHoliday           : Create a new (empty) holiday schedule in the tenant.
    - Set-TeamsIVRHolidayTime       : Add or replace the date/time range on a holiday schedule.
    - Add-TeamsIVRHolidayToIVR      : Attach an existing holiday schedule to an Auto Attendant.
    - Set-TeamsIVRHolidayGreeting   : Set the greeting message played when an IVR is in holiday mode.
    - Set-TeamsIVRHolidayRoutingMessage : Set the routing-options prompt for an IVR/holiday pair.
    - Set-TeamsIVRHolidayRouting    : Set the call routing action (redirect/disconnect/etc.) for
                                      an IVR/holiday pair.
    - Get-TeamsIVRHolidayMenuOptions : List keypress menu options on a holiday call flow.
    - Add-TeamsIVRHolidayMenuOption : Add or replace a single keypress binding on a holiday call flow.
    - Remove-TeamsIVRHolidayMenuOption : Remove a keypress binding from a holiday call flow.
    - Set-TeamsIVRHolidayMenu       : Replace all keypress bindings on a holiday call flow from JSON.
    - Set-TeamsIVRGreeting          : Set the greeting on an IVR's default (business-hours) call flow.
    - Set-TeamsIVRRoutingMessage    : Set the routing-options prompt on an IVR's default call flow.
    - Set-TeamsIVRRouting           : Set the call routing action on an IVR's default call flow.
    - Get-TeamsIVRMenuOptions       : List keypress menu options on an IVR's default call flow.
    - Add-TeamsIVRMenuOption        : Add or replace a single keypress binding on an IVR's default call flow.
    - Remove-TeamsIVRMenuOption     : Remove a keypress binding from an IVR's default call flow.
    - Set-TeamsIVRMenu              : Replace all keypress bindings on an IVR's default call flow from JSON.

.PARAMETER TenantId
    The Azure AD Tenant ID to connect to. Optional — if omitted, uses the default tenant.

.PARAMETER Mode
    Selects the operating mode. Default: Interactive.
      Interactive  Human operator in a terminal. Shows banner, progress bars, and
                   colored output. Uses interactive OAuth authentication.
      Headless     Cron job or calling script. Suppresses banner and progress bars;
                   uses non-interactive auth. Output is plain text.
      Api          Web-app caller. Suppresses all decorative output. Every response
                   is a single JSON line on stdout:
                     Success: {"ok":true,"data":[...]}  or  {"ok":true,"message":"..."}
                     Failure: {"ok":false,"error":"..."}  (exit code 1)
                   Non-interactive auth is designed but not yet activated -- three options
                   are available in Connect-TeamsEnvironment (certificate, client secret,
                   managed identity) that MSops must review and enable.

.NOTES
    Script Version : 0.2
    Required module: MicrosoftTeams
    Install with   : Install-Module -Name MicrosoftTeams -Force -AllowClobber
#>

[CmdletBinding(DefaultParameterSetName = 'ShowUsage')]
param(
    # ---- Shared parameters (available in all modes) ----

    [string]$TenantId,

    # Operating mode: Interactive (default), Headless (cron/script), or Api (web-app JSON output)
    [ValidateSet('Interactive', 'Headless', 'Api')]
    [string]$Mode = 'Interactive',

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
    # No additional parameters required. Set -Mode Headless (or Api) and ensure the hosting environment
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
    [ValidateNotNullOrEmpty()]
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
    [ValidateNotNullOrEmpty()]
    [string]$User,

    # ---- Shared: Identity used by schedule and user-access actions ----

    [Parameter(ParameterSetName = 'GetSchedule',              Mandatory = $true)]
    [Parameter(ParameterSetName = 'DecodeSchedule',           Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetSchedule',              Mandatory = $true)]
    [Parameter(ParameterSetName = 'GetAuthorizedUsers',       Mandatory = $true)]
    [Parameter(ParameterSetName = 'CheckUserAccess',          Mandatory = $true)]
    [Parameter(ParameterSetName = 'NewHoliday',               Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayTime',           Mandatory = $true)]
    [Parameter(ParameterSetName = 'AttachHolidayToIVR',       Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayGreeting',       Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayRoutingMessage', Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayRouting',        Mandatory = $true)]
    [Parameter(ParameterSetName = 'GetHolidayMenuOptions',    Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddHolidayMenuOption',     Mandatory = $true)]
    [Parameter(ParameterSetName = 'RemoveHolidayMenuOption',  Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayMenu',           Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRGreeting',           Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRRoutingMessage',     Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRRouting',            Mandatory = $true)]
    [Parameter(ParameterSetName = 'GetIVRMenuOptions',        Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddIVRMenuOption',         Mandatory = $true)]
    [Parameter(ParameterSetName = 'RemoveIVRMenuOption',      Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRMenu',               Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Identity,

    # ---- Action: List all tenant holiday schedules ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetHolidayList

    [Parameter(ParameterSetName = 'GetHolidayList', Mandatory = $true)]
    [switch]$GetHolidayList,

    # ---- Action: Get a specific holiday schedule by name ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetHoliday -HolidayName "Christmas"

    [Parameter(ParameterSetName = 'GetHoliday', Mandatory = $true)]
    [switch]$GetHoliday,

    # ---- Action: Create a new holiday on an Auto Attendant ----
    # Usage: .\Manage-TeamsIVR.ps1 -NewHoliday -Identity "aa_guid" -HolidayName "Christmas"
    # Creates an empty schedule + default disconnect call flow.
    # Use -SetHolidayTime to add dates, -SetHolidayGreeting / -SetHolidayRouting to configure.
    # Note: Teams stores holidays inside AA objects -- -Identity is required to know which AA to update.

    [Parameter(ParameterSetName = 'NewHoliday', Mandatory = $true)]
    [switch]$NewHoliday,

    # ---- Action: Set the date/time range on a holiday schedule ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetHolidayTime -Identity "aa_guid" -HolidayName "Christmas" -StartDateTime "2025-12-25 00:00" -EndDateTime "2025-12-26 00:00"
    # NOTE: Both times must fall on 15-minute boundaries (minutes 0, 15, 30, or 45). Use next-day midnight to cover a full day.
    # Replaces any existing date ranges on the named holiday with the single provided range.

    [Parameter(ParameterSetName = 'SetHolidayTime', Mandatory = $true)]
    [switch]$SetHolidayTime,

    [Parameter(ParameterSetName = 'SetHolidayTime', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$StartDateTime,

    [Parameter(ParameterSetName = 'SetHolidayTime', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$EndDateTime,

    # ---- Action: Attach an existing holiday schedule to an IVR ----
    # Usage: .\Manage-TeamsIVR.ps1 -AttachHolidayToIVR -Identity "aa_guid" -HolidayName "Christmas"

    [Parameter(ParameterSetName = 'AttachHolidayToIVR', Mandatory = $true)]
    [switch]$AttachHolidayToIVR,

    # ---- Action: Set the greeting message for an IVR/holiday pair ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetHolidayGreeting -Identity "aa_guid" -HolidayName "Christmas" -GreetingText "We are closed for Christmas."

    [Parameter(ParameterSetName = 'SetHolidayGreeting', Mandatory = $true)]
    [switch]$SetHolidayGreeting,

    # ---- Action: Set the greeting message for an IVR's default (business-hours) call flow ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetIVRGreeting -Identity "aa_guid" -GreetingText "Thank you for calling Contoso."

    [Parameter(ParameterSetName = 'SetIVRGreeting', Mandatory = $true)]
    [switch]$SetIVRGreeting,

    # Shared: greeting text used by SetHolidayGreeting and SetIVRGreeting
    [Parameter(ParameterSetName = 'SetHolidayGreeting', Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRGreeting',     Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateLength(1, 1000)]
    [string]$GreetingText,

    # ---- Action: Set the routing-options prompt for an IVR/holiday pair ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetHolidayRoutingMessage -Identity "aa_guid" -HolidayName "Christmas" -RoutingMessage "Press 1 for voicemail."

    [Parameter(ParameterSetName = 'SetHolidayRoutingMessage', Mandatory = $true)]
    [switch]$SetHolidayRoutingMessage,

    # ---- Action: Set the routing-options prompt for an IVR's default call flow ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetIVRRoutingMessage -Identity "aa_guid" -RoutingMessage "Press 1 for Sales, press 2 for Support."

    [Parameter(ParameterSetName = 'SetIVRRoutingMessage', Mandatory = $true)]
    [switch]$SetIVRRoutingMessage,

    # Shared: routing message used by SetHolidayRoutingMessage and SetIVRRoutingMessage
    [Parameter(ParameterSetName = 'SetHolidayRoutingMessage', Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRRoutingMessage',     Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateLength(1, 1000)]
    [string]$RoutingMessage,

    # ---- Action: Set the call routing action for an IVR/holiday pair ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetHolidayRouting -Identity "aa_guid" -HolidayName "Christmas" -RoutingAction Disconnect
    # Usage: .\Manage-TeamsIVR.ps1 -SetHolidayRouting -Identity "aa_guid" -HolidayName "Christmas" -RoutingAction RedirectToUser -RedirectTarget "user@contoso.com"
    # Valid RoutingAction values: Disconnect | RedirectToUser | RedirectToPSTN | RedirectToVoicemail | RedirectToAA
    # -RedirectTarget is required for all actions except Disconnect:
    #   RedirectToUser/RedirectToVoicemail : UPN or object ID of the target user
    #   RedirectToPSTN                     : E.164 phone number (e.g. +15551234567)
    #   RedirectToAA                       : Identity (GUID or name) of the target Auto Attendant

    [Parameter(ParameterSetName = 'SetHolidayRouting', Mandatory = $true)]
    [switch]$SetHolidayRouting,

    # ---- Action: Set the call routing action for an IVR's default call flow ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetIVRRouting -Identity "aa_guid" -RoutingAction Disconnect
    # Usage: .\Manage-TeamsIVR.ps1 -SetIVRRouting -Identity "aa_guid" -RoutingAction RedirectToUser -RedirectTarget "user@contoso.com"
    # Same RoutingAction values and -RedirectTarget semantics as -SetHolidayRouting.

    [Parameter(ParameterSetName = 'SetIVRRouting', Mandatory = $true)]
    [switch]$SetIVRRouting,

    # ---- Action: View all keypress menu options on an IVR/holiday call flow ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetHolidayMenuOptions -Identity "aa_guid" -HolidayName "Christmas"
    # Prints a table of key -> action -> target and a JSON blob for use with -SetHolidayMenu.

    [Parameter(ParameterSetName = 'GetHolidayMenuOptions', Mandatory = $true)]
    [switch]$GetHolidayMenuOptions,

    # ---- Action: Add or replace a single keypress option on an IVR/holiday call flow menu ----
    # Usage: .\Manage-TeamsIVR.ps1 -AddHolidayMenuOption -Identity "aa_guid" -HolidayName "Christmas" -Key 1 -RoutingAction RedirectToUser -RedirectTarget "user@contoso.com"

    [Parameter(ParameterSetName = 'AddHolidayMenuOption', Mandatory = $true)]
    [switch]$AddHolidayMenuOption,

    # ---- Action: Remove a single keypress option from an IVR/holiday call flow menu ----
    # Usage: .\Manage-TeamsIVR.ps1 -RemoveHolidayMenuOption -Identity "aa_guid" -HolidayName "Christmas" -Key 1

    [Parameter(ParameterSetName = 'RemoveHolidayMenuOption', Mandatory = $true)]
    [switch]$RemoveHolidayMenuOption,

    # ---- Action: Replace all keypress options on an IVR/holiday call flow menu (bulk JSON) ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetHolidayMenu -Identity "aa_guid" -HolidayName "Christmas" -MenuOptionsJson '[{"Key":"1","Action":"RedirectToUser","Target":"user@contoso.com"}]'
    # JSON format: array of {Key, Action, Target} -- same Key/Action values as -AddHolidayMenuOption.
    # Use -GetHolidayMenuOptions to retrieve the current options as copy-paste JSON.

    [Parameter(ParameterSetName = 'SetHolidayMenu', Mandatory = $true)]
    [switch]$SetHolidayMenu,

    # ---- Action: View all keypress menu options on an IVR's default call flow ----
    # Usage: .\Manage-TeamsIVR.ps1 -GetIVRMenuOptions -Identity "aa_guid"

    [Parameter(ParameterSetName = 'GetIVRMenuOptions', Mandatory = $true)]
    [switch]$GetIVRMenuOptions,

    # ---- Action: Add or replace a single keypress option on an IVR's default call flow menu ----
    # Usage: .\Manage-TeamsIVR.ps1 -AddIVRMenuOption -Identity "aa_guid" -Key 1 -RoutingAction RedirectToUser -RedirectTarget "user@contoso.com"

    [Parameter(ParameterSetName = 'AddIVRMenuOption', Mandatory = $true)]
    [switch]$AddIVRMenuOption,

    # ---- Action: Remove a single keypress option from an IVR's default call flow menu ----
    # Usage: .\Manage-TeamsIVR.ps1 -RemoveIVRMenuOption -Identity "aa_guid" -Key 1

    [Parameter(ParameterSetName = 'RemoveIVRMenuOption', Mandatory = $true)]
    [switch]$RemoveIVRMenuOption,

    # ---- Action: Replace all keypress options on an IVR's default call flow menu (bulk JSON) ----
    # Usage: .\Manage-TeamsIVR.ps1 -SetIVRMenu -Identity "aa_guid" -MenuOptionsJson '[{"Key":"1","Action":"RedirectToUser","Target":"user@contoso.com"}]'

    [Parameter(ParameterSetName = 'SetIVRMenu', Mandatory = $true)]
    [switch]$SetIVRMenu,

    # Shared: keypress key used by AddHolidayMenuOption, RemoveHolidayMenuOption, AddIVRMenuOption, RemoveIVRMenuOption
    # Valid values: 0-9 and * #
    [Parameter(ParameterSetName = 'AddHolidayMenuOption',    Mandatory = $true)]
    [Parameter(ParameterSetName = 'RemoveHolidayMenuOption', Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddIVRMenuOption',        Mandatory = $true)]
    [Parameter(ParameterSetName = 'RemoveIVRMenuOption',     Mandatory = $true)]
    [ValidateSet('0','1','2','3','4','5','6','7','8','9','*','#')]
    [string]$Key,

    # Shared: JSON menu definition used by SetHolidayMenu and SetIVRMenu
    [Parameter(ParameterSetName = 'SetHolidayMenu', Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRMenu',     Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$MenuOptionsJson,

    # Shared: routing action and redirect target used by SetHolidayRouting, SetIVRRouting, AddHolidayMenuOption, AddIVRMenuOption
    [Parameter(ParameterSetName = 'SetHolidayRouting',    Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetIVRRouting',        Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddHolidayMenuOption', Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddIVRMenuOption',     Mandatory = $true)]
    [ValidateSet('Disconnect', 'RedirectToUser', 'RedirectToPSTN', 'RedirectToVoicemail', 'RedirectToAA')]
    [string]$RoutingAction,

    # Required for all RoutingAction values except 'Disconnect'
    [Parameter(ParameterSetName = 'SetHolidayRouting',    Mandatory = $false)]
    [Parameter(ParameterSetName = 'SetIVRRouting',        Mandatory = $false)]
    [Parameter(ParameterSetName = 'AddHolidayMenuOption', Mandatory = $false)]
    [Parameter(ParameterSetName = 'AddIVRMenuOption',     Mandatory = $false)]
    [string]$RedirectTarget,

    # ---- Shared: HolidayName used across all holiday actions ----

    [Parameter(ParameterSetName = 'GetHoliday',              Mandatory = $true)]
    [Parameter(ParameterSetName = 'NewHoliday',              Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayTime',          Mandatory = $true)]
    [Parameter(ParameterSetName = 'AttachHolidayToIVR',      Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayGreeting',      Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayRoutingMessage',Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayRouting',       Mandatory = $true)]
    [Parameter(ParameterSetName = 'GetHolidayMenuOptions',   Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddHolidayMenuOption',    Mandatory = $true)]
    [Parameter(ParameterSetName = 'RemoveHolidayMenuOption', Mandatory = $true)]
    [Parameter(ParameterSetName = 'SetHolidayMenu',          Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateLength(1, 256)]
    [string]$HolidayName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Constants

$Script:Version = '0.2'

# Unique 8-char ID for this script run -- included in all API responses and Verbose output
# so log lines from concurrent invocations can be correlated.
$Script:RunId = [guid]::NewGuid().ToString('N').Substring(0, 8)

# In-process AA object cache: keyed by Identity string.
# Prevents redundant Get-CsAutoAttendant calls within one invocation (e.g. Build-MenuOptionFromSpec
# resolving multiple RedirectToAA targets). NOT automatically invalidated -- callers that mutate
# an AA must call $Script:AACache.Remove($Identity) afterward to avoid stale reads in the same session.
$Script:AACache = @{}

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
            # Non-interactive auth will NOT work until MSops activates an option above.
            Write-Verbose "Non-interactive mode: auth not yet configured -- falling back to interactive."
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
    [OutputType([PSCustomObject[]])]
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
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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
    [OutputType([string[]])]
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

    .NOTES
        New-CsOnlineSchedule creates a persistent schedule object in Teams when called.
        If Set-TeamsIVRSchedule is never called with the returned object (or fails),
        the schedule will be orphaned in Teams and must be removed manually via
        Get-CsOnlineSchedule / Remove-CsOnlineSchedule.
    .OUTPUTS
        CsOnlineSchedule -- pass directly to Set-TeamsIVRSchedule -Schedule.
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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
        [ValidateNotNullOrEmpty()]
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
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Schedule updated for '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to update schedule for '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Get-TeamsIVRAuthorizedUser

function Get-TeamsIVRAuthorizedUser {
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
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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
        [ValidateNotNullOrEmpty()]
        [string]$User
    )

    try {
        return Get-CsOnlineUser -Identity $User
    } catch {
        try {
            # Use Where-Object instead of -Filter string to avoid OData injection via
            # special characters in display names (e.g. single quotes).
            $nameMatches = @(Get-CsOnlineUser | Where-Object { $_.DisplayName -eq $User })
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
          - Object ID     : GUID (e.g. from Get-TeamsIVRAuthorizedUser output)
          - Display name  : "Jane Smith"  (must be unique in the tenant)
    .OUTPUTS
        Boolean -- $true if the user is authorized, $false if not.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
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

#region Get-TeamsIVRByUser

function Get-TeamsIVRByUser {
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
          - Object ID     : GUID (e.g. from Get-TeamsIVRAuthorizedUser output)
          - Display name  : "Jane Smith"  (must be unique in the tenant)
    .PARAMETER Headless
        Suppress progress bar output.
    .OUTPUTS
        PSCustomObject[] with properties: Name, Identity
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$User,

        [switch]$Headless
    )

    $resolvedUser = Resolve-TeamsUser -User $User

    Write-Verbose "Scanning Auto Attendants for user '$($resolvedUser.UserPrincipalName)'..."

    try {
        $allAAs = @(Invoke-TeamsRetry { Get-CsAutoAttendant } 'Get-CsAutoAttendant (tenant scan)')
    } catch {
        Write-Error "Failed to retrieve Auto Attendants: $_"
        throw
    }

    $aaCount = @($allAAs).Count
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

#region Find-TeamsIVRHolidayContext (private helper)

function Find-TeamsIVRHolidayContext {
    <#
    .SYNOPSIS
        Finds the schedule, call flow, and association for a named holiday on a given AA.
    .DESCRIPTION
        Private helper used by all holiday write-functions. Scans the AA's
        CallHandlingAssociations for Holiday-type entries, matches the associated
        schedule name against $HolidayName, and returns a context object with all
        three pieces needed to modify or replace the holiday.
    .PARAMETER AA
        A CsAutoAttendant object (already retrieved with Get-CsAutoAttendant).
    .PARAMETER HolidayName
        Name of the holiday schedule to locate.
    .OUTPUTS
        PSCustomObject with properties: Assoc, Schedule, CallFlow
        Returns $null if the named holiday is not found on the AA.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$AA,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName
    )

    $holidayAssocs = @($AA.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq 'Holiday' })

    foreach ($assoc in $holidayAssocs) {
        $schedule = @($AA.Schedules | Where-Object { $_.Id -eq $assoc.ScheduleId }) | Select-Object -First 1
        if (-not $schedule) { continue }
        if ($schedule.Name -ne $HolidayName) { continue }

        $callFlow = @($AA.CallFlows | Where-Object { $_.Id -eq $assoc.CallFlowId }) | Select-Object -First 1

        return [PSCustomObject]@{
            Assoc    = $assoc
            Schedule = $schedule
            CallFlow = $callFlow
        }
    }

    return $null
}

#endregion

#region Menu-option private helpers

function Convert-DtmfKeyToTone {
    # Maps user-facing key label (0-9, *, #) to Teams DtmfResponse enum string.
    param([Parameter(Mandatory = $true)][string]$Key)
    switch ($Key) {
        '0' { return 'Tone0'     }
        '1' { return 'Tone1'     }
        '2' { return 'Tone2'     }
        '3' { return 'Tone3'     }
        '4' { return 'Tone4'     }
        '5' { return 'Tone5'     }
        '6' { return 'Tone6'     }
        '7' { return 'Tone7'     }
        '8' { return 'Tone8'     }
        '9' { return 'Tone9'     }
        '*' { return 'ToneStar'  }
        '#' { return 'TonePound' }
        default { throw "Unknown key '$Key'. Valid: 0-9, *, #" }
    }
}

function Convert-ToneToDisplayKey {
    # Maps Teams DtmfResponse enum string back to a human-readable key label.
    param([Parameter(Mandatory = $true)][string]$Tone)
    switch ($Tone) {
        'Tone0'     { return '0'      }
        'Tone1'     { return '1'      }
        'Tone2'     { return '2'      }
        'Tone3'     { return '3'      }
        'Tone4'     { return '4'      }
        'Tone5'     { return '5'      }
        'Tone6'     { return '6'      }
        'Tone7'     { return '7'      }
        'Tone8'     { return '8'      }
        'Tone9'     { return '9'      }
        'ToneStar'  { return '*'      }
        'TonePound' { return '#'      }
        'Automatic' { return '(auto)' }
        default     { return $Tone    }
    }
}

function Convert-MenuOptionToActionSpec {
    # Converts a CsAutoAttendantMenuOption to a PSCustomObject {Key, Action, Target}
    # using the same Action names accepted by -RoutingAction / -SetHolidayMenu JSON.
    param([Parameter(Mandatory = $true)][object]$MenuOption)

    $key    = Convert-ToneToDisplayKey -Tone $MenuOption.DtmfResponse.ToString()
    $action = $null
    $target = $null

    switch ($MenuOption.Action.ToString()) {
        'DisconnectCall' {
            $action = 'Disconnect'
        }
        'TransferCallToTarget' {
            if ($MenuOption.CallTarget) {
                switch ($MenuOption.CallTarget.Type.ToString()) {
                    'User'                { $action = 'RedirectToUser'      }
                    'ExternalPstn'        { $action = 'RedirectToPSTN'      }
                    'SharedVoicemail'     { $action = 'RedirectToVoicemail' }
                    'ApplicationEndpoint' { $action = 'RedirectToAA'        }
                    default               { $action = "TransferTo($($MenuOption.CallTarget.Type))" }
                }
                $target = $MenuOption.CallTarget.Identity
            } else {
                $action = 'TransferCallToTarget'
            }
        }
        default {
            $action = $MenuOption.Action.ToString()
            if ($MenuOption.CallTarget) { $target = $MenuOption.CallTarget.Identity }
        }
    }

    return [PSCustomObject]@{ Key = $key; Action = $action; Target = $target }
}

function Build-MenuOptionFromSpec {
    # Creates a CsAutoAttendantMenuOption from a PSCustomObject {Key, Action, Target}.
    # Key must be 0-9, *, or # (or '(auto)' for DtmfResponse Automatic).
    # Action must be one of the ValidateSet values used by -RoutingAction.
    param([Parameter(Mandatory = $true)][PSCustomObject]$Spec)

    $tone = if ($Spec.Key -eq '(auto)') { 'Automatic' } else { Convert-DtmfKeyToTone -Key $Spec.Key }

    switch ($Spec.Action) {
        'Disconnect' {
            return New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse $tone
        }
        'RedirectToUser' {
            $resolvedUser = Resolve-TeamsUser -User $Spec.Target
            $entity = New-CsAutoAttendantCallableEntity -Identity $resolvedUser.Identity -Type User
            return New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse $tone -CallTarget $entity
        }
        'RedirectToPSTN' {
            Assert-E164Format -PhoneNumber $Spec.Target -ParamName "Key '$($Spec.Key)' Target"
            $entity = New-CsAutoAttendantCallableEntity -Identity $Spec.Target -Type ExternalPstn
            return New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse $tone -CallTarget $entity
        }
        'RedirectToVoicemail' {
            $resolvedUser = Resolve-TeamsUser -User $Spec.Target
            $entity = New-CsAutoAttendantCallableEntity -Identity $resolvedUser.Identity -Type SharedVoicemail
            return New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse $tone -CallTarget $entity
        }
        'RedirectToAA' {
            try {
                $targetAA = Get-CsAutoAttendantCached -Identity $Spec.Target
            } catch {
                throw "Could not find target Auto Attendant '$($Spec.Target)': $_"
            }
            $entity = New-CsAutoAttendantCallableEntity -Identity $targetAA.Identity -Type ApplicationEndpoint
            return New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse $tone -CallTarget $entity
        }
        default {
            throw "Unknown action '$($Spec.Action)'. Valid: Disconnect, RedirectToUser, RedirectToPSTN, RedirectToVoicemail, RedirectToAA"
        }
    }
}

function Assert-E164Format {
    # Throws a descriptive error if the supplied string is not E.164 format (+<country><number>).
    param([Parameter(Mandatory)][string]$PhoneNumber, [string]$ParamName = '-RedirectTarget')
    if ($PhoneNumber -notmatch '^\+[1-9]\d{6,14}$') {
        throw "$ParamName '$PhoneNumber' is not valid E.164 format. Use +<country-code><number>, e.g. +15551234567."
    }
}

function Format-MenuOptionsAsJson {
    # Converts an array of {Key, Action, Target} specs to a compact JSON string
    # suitable for copy-paste back into -SetHolidayMenu / -SetIVRMenu.
    param([Parameter(Mandatory = $true)][object[]]$Specs)
    $lines = $Specs | ForEach-Object {
        $obj = [ordered]@{ Key = $_.Key; Action = $_.Action }
        if (-not [string]::IsNullOrWhiteSpace($_.Target)) { $obj['Target'] = $_.Target }
        [PSCustomObject]$obj | ConvertTo-Json -Compress -Depth 3
    }
    return '[' + ($lines -join ',') + ']'
}

#endregion

#region Get-TeamsIVRHolidayList

function Get-TeamsIVRHolidayList {
    <#
    .SYNOPSIS
        Retrieves all holiday schedules across all Auto Attendants in the tenant.
    .DESCRIPTION
        Teams does not expose a tenant-level holiday schedule list cmdlet. Holiday
        schedules are embedded in Auto Attendant objects, identified by
        CallHandlingAssociations of type 'Holiday'. This function scans every AA,
        collects holiday schedules, deduplicates by schedule ID (a schedule can be
        reused across AAs), and returns a flat list.

        Each output object includes the IVR(s) the holiday is linked to so you can
        see which AAs are affected without a separate lookup.
    .PARAMETER Headless
        Suppress progress bar output.
    .OUTPUTS
        PSCustomObject[] with properties:
          Name           - Holiday schedule display name
          ScheduleId     - GUID of the CsOnlineSchedule object
          DateTimeRanges - Array of date/time range objects (Start, End)
          LinkedIVRs     - Array of IVR names that reference this schedule
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [switch]$Headless
    )

    Write-Verbose "Retrieving holiday schedules from all Auto Attendants..."

    try {
        $allAAs = @(Invoke-TeamsRetry { Get-CsAutoAttendant } 'Get-CsAutoAttendant (tenant scan)')
    } catch {
        Write-Error "Failed to retrieve Auto Attendants: $_"
        throw
    }

    $aaCount = @($allAAs).Count
    $i = 0

    # Key: schedule ID, Value: [Name, DateTimeRanges, [linked IVR names]]
    $seen = [System.Collections.Generic.Dictionary[string, PSCustomObject]]::new()

    foreach ($aa in $allAAs) {
        $i++
        if (-not $Headless) {
            Write-Progress -Activity "Scanning Auto Attendants for holidays" `
                           -Status "$i of $aaCount : $($aa.Name)" `
                           -PercentComplete (($i / $aaCount) * 100)
        }

        $holidayAssocs = @($aa.CallHandlingAssociations | Where-Object { $_.Type.ToString() -eq 'Holiday' })

        foreach ($assoc in $holidayAssocs) {
            $schedule = @($aa.Schedules | Where-Object { $_.Id -eq $assoc.ScheduleId }) | Select-Object -First 1
            if (-not $schedule) { continue }

            $id = $schedule.Id.ToString()

            if ($seen.ContainsKey($id)) {
                # Already recorded -- just add this IVR to the linked list if not already there
                if ($seen[$id].LinkedIVRs -notcontains $aa.Name) {
                    $seen[$id].LinkedIVRs += $aa.Name
                }
            } else {
                # Extract date/time ranges from FixedSchedule
                $ranges = @()
                if ($schedule.FixedSchedule -and $schedule.FixedSchedule.DateTimeRanges) {
                    $ranges = @($schedule.FixedSchedule.DateTimeRanges | ForEach-Object {
                        [PSCustomObject]@{
                            Start = $_.Start
                            End   = $_.End
                        }
                    })
                }

                $seen[$id] = [PSCustomObject]@{
                    Name           = $schedule.Name
                    ScheduleId     = $id
                    DateTimeRanges = $ranges
                    LinkedIVRs     = @($aa.Name)
                }
            }
        }
    }

    if (-not $Headless) { Write-Progress -Activity "Scanning Auto Attendants for holidays" -Completed }

    return @($seen.Values)
}

#endregion

#region Get-TeamsIVRHoliday

function Get-TeamsIVRHoliday {
    <#
    .SYNOPSIS
        Retrieves every occurrence of a named holiday schedule across all Auto Attendants.
    .DESCRIPTION
        Teams stores holiday schedules inside AA objects, not at the tenant level.
        This function scans all AAs and returns one result per AA that has a holiday
        matching $HolidayName. If the same schedule name exists on multiple AAs you
        get multiple results, each with its IVR context.
    .PARAMETER HolidayName
        Name of the holiday schedule to find (exact match, case-sensitive).
    .PARAMETER Headless
        Suppress progress bar output.
    .OUTPUTS
        PSCustomObject[] with properties: HolidayName, ScheduleId, DateTimeRanges,
        IVRName, IVRIdentity, CallFlowId
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [switch]$Headless
    )

    Write-Verbose "Searching for holiday '$HolidayName' across all Auto Attendants..."

    try {
        $allAAs = @(Invoke-TeamsRetry { Get-CsAutoAttendant } 'Get-CsAutoAttendant (tenant scan)')
    } catch {
        Write-Error "Failed to retrieve Auto Attendants: $_"
        throw
    }

    $aaCount = @($allAAs).Count
    $i       = 0
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($aa in $allAAs) {
        $i++
        if (-not $Headless) {
            Write-Progress -Activity "Searching for holiday '$HolidayName'" `
                           -Status "$i of $aaCount : $($aa.Name)" `
                           -PercentComplete (($i / $aaCount) * 100)
        }

        $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
        if (-not $ctx) { continue }

        $ranges = @()
        if ($ctx.Schedule.FixedSchedule -and $ctx.Schedule.FixedSchedule.DateTimeRanges) {
            $ranges = @($ctx.Schedule.FixedSchedule.DateTimeRanges | ForEach-Object {
                [PSCustomObject]@{ Start = $_.Start; End = $_.End }
            })
        }

        $results.Add([PSCustomObject]@{
            HolidayName    = $ctx.Schedule.Name
            ScheduleId     = $ctx.Schedule.Id
            DateTimeRanges = $ranges
            IVRName        = $aa.Name
            IVRIdentity    = $aa.Identity
            CallFlowId     = if ($ctx.CallFlow) { $ctx.CallFlow.Id } else { $null }
        })
    }

    if (-not $Headless) { Write-Progress -Activity "Searching for holiday '$HolidayName'" -Completed }

    return $results
}

#endregion

#region New-TeamsIVRHoliday

function New-TeamsIVRHoliday {
    <#
    .SYNOPSIS
        Creates a new holiday on an Auto Attendant with a default disconnect call flow.
    .DESCRIPTION
        Teams stores holiday schedules inside AA objects rather than as tenant-level
        entities, so an IVR identity is required. This function:
          1. Verifies the holiday does not already exist on the AA.
          2. Creates a FixedSchedule with no date ranges (add dates with Set-TeamsIVRHolidayTime).
          3. Creates a minimal call flow: no greeting, menu with automatic disconnect.
          4. Creates a CallHandlingAssociation of type Holiday linking the two.
          5. Rebuilds the AA's Schedules, CallFlows, and CallHandlingAssociations arrays
             (they are fixed-size in Teams and cannot be mutated in place) and saves.

        After creation, configure the holiday with:
          Set-TeamsIVRHolidayTime       -- add date ranges
          Set-TeamsIVRHolidayGreeting   -- set the greeting callers hear
          Set-TeamsIVRHolidayRouting    -- change from disconnect to a redirect if needed
    .NOTES
        WhatIf safety: all side-effecting operations (including New-CsOnlineSchedule) are
        guarded inside the ShouldProcess block and are fully skipped when -WhatIf is used.
        If a runtime error occurs after the schedule is created (step 1) but before the
        final Set-CsAutoAttendant commit succeeds, an orphaned schedule object will remain
        in Teams and must be removed manually (Get-CsOnlineSchedule / Remove-CsOnlineSchedule).
    .PARAMETER Identity
        Identity of the Auto Attendant to add the holiday to.
    .PARAMETER HolidayName
        Display name for the new holiday (e.g. "Christmas 2025").
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    # Guard: prevent duplicate holiday names on the same AA
    $existing = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if ($existing) {
        throw "Holiday '$HolidayName' already exists on '$($aa.Name)'. Use Set-TeamsIVRHolidayTime / Set-TeamsIVRHolidayGreeting to modify it."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Add holiday '$HolidayName'")) {

        # 1. Create the holiday schedule (FixedSchedule, no date ranges yet)
        try {
            $schedule = New-CsOnlineSchedule -Name $HolidayName -FixedSchedule
        } catch {
            throw "Failed to create holiday schedule '$HolidayName': $_"
        }

        # 2. Create a minimal call flow: automatic disconnect, no greeting
        try {
            $disconnectOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
            $menu             = New-CsAutoAttendantMenu -Name "$HolidayName Menu" -MenuOptions @($disconnectOption)
            $callFlow         = New-CsAutoAttendantCallFlow -Name "$HolidayName Flow" -Menu $menu
        } catch {
            throw "Failed to create holiday call flow for '$HolidayName': $_"
        }

        # 3. Create the association linking schedule -> call flow
        try {
            $assoc = New-CsAutoAttendantCallHandlingAssociation `
                -Type Holiday `
                -ScheduleId $schedule.Id `
                -CallFlowId $callFlow.Id
        } catch {
            throw "Failed to create CallHandlingAssociation for '$HolidayName': $_"
        }

        # 4. Rebuild the fixed-size arrays with the new objects appended
        $aa.Schedules              = @($aa.Schedules)              + @($schedule)
        $aa.CallFlows              = @($aa.CallFlows)              + @($callFlow)
        $aa.CallHandlingAssociations = @($aa.CallHandlingAssociations) + @($assoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Holiday '$HolidayName' created on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRHolidayTime

function Set-TeamsIVRHolidayTime {
    <#
    .SYNOPSIS
        Sets (replaces) the date/time range on a holiday schedule attached to an AA.
    .DESCRIPTION
        Locates the named holiday on the specified Auto Attendant, builds a new
        CsOnlineDateTimeRange from the provided start/end strings, and replaces the
        holiday's entire date range list with the single provided range.

        Because Teams schedule objects are immutable once assigned, this function:
          1. Creates a brand-new CsOnlineSchedule with the new date range.
          2. Creates a new CallHandlingAssociation pointing to the new schedule
             (keeping the existing call flow ID unchanged).
          3. Replaces the old schedule and association in the AA and saves.

        If you need to set multiple date ranges for a multi-day holiday, call this
        function and extend it, or modify the AA object directly.
    .NOTES
        WhatIf safety: all side-effecting operations (including the replacement
        New-CsOnlineSchedule) are guarded inside the ShouldProcess block and are fully
        skipped when -WhatIf is used. If a runtime error occurs after the new schedule is
        created but before the final Set-CsAutoAttendant commit succeeds, the new schedule
        object will be orphaned in Teams and must be removed manually
        (Get-CsOnlineSchedule / Remove-CsOnlineSchedule).
    .PARAMETER Identity
        Identity of the Auto Attendant that owns the holiday.
    .PARAMETER HolidayName
        Name of the holiday schedule to update.
    .PARAMETER StartDateTime
        Start of the holiday period. Accepts any .NET-parseable date/time string
        (e.g. '2025-12-25 00:00', '12/25/2025', '2025-12-25T00:00:00').
        IMPORTANT: Must fall on a 15-minute boundary -- minutes must be 0, 15, 30, or 45
        and seconds must be 0. Teams rejects any other value.
    .PARAMETER EndDateTime
        End of the holiday period. Must be later than StartDateTime.
        Same 15-minute boundary restriction applies. To cover an entire day, use midnight
        of the following day (e.g. '2025-12-26 00:00' to cover all of December 25).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$StartDateTime,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$EndDateTime
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'. Create it first with New-TeamsIVRHoliday."
    }

    # Parse and validate the date strings.
    # Use ParseExact with InvariantCulture to avoid locale-specific date interpretation
    # (e.g. 01/02/2025 parsing as Jan 2 on US servers but Feb 1 on EU servers).
    $dateFormats = [string[]]@(
        'yyyy-MM-dd HH:mm',
        'yyyy-MM-ddTHH:mm:ss',
        'yyyy-MM-ddTHH:mm',
        'M/d/yyyy H:mm',
        'M/d/yyyy'
    )
    try {
        $start = [DateTime]::ParseExact($StartDateTime, $dateFormats,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None)
    } catch {
        throw "Could not parse -StartDateTime '$StartDateTime'. Accepted formats: 'yyyy-MM-dd HH:mm', 'yyyy-MM-ddTHH:mm:ss', 'M/d/yyyy H:mm'."
    }
    try {
        $end = [DateTime]::ParseExact($EndDateTime, $dateFormats,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None)
    } catch {
        throw "Could not parse -EndDateTime '$EndDateTime'. Accepted formats: 'yyyy-MM-dd HH:mm', 'yyyy-MM-ddTHH:mm:ss', 'M/d/yyyy H:mm'."
    }
    # Teams requires date-time boundaries to align with 15-minute intervals (0, 15, 30, or 45 minutes).
    if ($start.Minute % 15 -ne 0 -or $start.Second -ne 0) {
        throw "StartDateTime must fall on a 15-minute boundary (minutes must be 0, 15, 30, or 45; seconds must be 0). Got: $($start.ToString('yyyy-MM-ddTHH:mm:ss'))"
    }
    if ($end.Minute % 15 -ne 0 -or $end.Second -ne 0) {
        throw "EndDateTime must fall on a 15-minute boundary (minutes must be 0, 15, 30, or 45; seconds must be 0). Got: $($end.ToString('yyyy-MM-ddTHH:mm:ss'))"
    }
    if ($start -ge $end) {
        throw "StartDateTime ('$StartDateTime') must be earlier than EndDateTime ('$EndDateTime')."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set date range on holiday '$HolidayName' to $start -- $end")) {

        # Build the new date range object Teams expects.
        # New-CsOnlineDateTimeRange only accepts specific string formats; passing a DateTime
        # object causes PowerShell to serialize it using the system locale format, which Teams
        # rejects. Format explicitly as yyyy-MM-ddTHH:mm:ss (one of three supported formats).
        try {
            $dateRange = New-CsOnlineDateTimeRange `
                -Start $start.ToString('yyyy-MM-ddTHH:mm:ss') `
                -End   $end.ToString('yyyy-MM-ddTHH:mm:ss')
        } catch {
            throw "Failed to create date/time range: $_"
        }

        # Create a replacement schedule with the same name but new date range
        try {
            $newSchedule = New-CsOnlineSchedule -Name $ctx.Schedule.Name -FixedSchedule -DateTimeRanges @($dateRange)
        } catch {
            throw "Failed to create replacement schedule for '$HolidayName': $_"
        }

        # New association keeping the same call flow, pointing to the new schedule ID
        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation `
                -Type Holiday `
                -ScheduleId $newSchedule.Id `
                -CallFlowId $ctx.Assoc.CallFlowId
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        # Rebuild fixed-size arrays: swap old schedule/assoc for new ones
        $aa.Schedules = @($aa.Schedules | Where-Object { $_.Id -ne $ctx.Schedule.Id }) + @($newSchedule)
        $aa.CallHandlingAssociations = @(
            $aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }
        ) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Date range updated for holiday '$HolidayName' on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Add-TeamsIVRHolidayToIVR

function Add-TeamsIVRHolidayToIVR {
    <#
    .SYNOPSIS
        Attaches an existing holiday schedule to an Auto Attendant.
    .DESCRIPTION
        Locates the first Auto Attendant that has a holiday named $HolidayName,
        copies its date ranges, and creates a matching holiday (schedule + default
        disconnect call flow) on the target AA. Teams does not support sharing a single
        schedule object across multiple AAs, so the dates are duplicated.

        The greeting, routing message, and routing action are NOT copied -- configure
        those separately with Set-TeamsIVRHolidayGreeting, Set-TeamsIVRHolidayRoutingMessage,
        and Set-TeamsIVRHolidayRouting.

        If the target AA already has a holiday with the same name, the function throws
        rather than silently overwriting.
    .NOTES
        WhatIf safety: all side-effecting operations (New-CsOnlineSchedule, Set-CsAutoAttendant)
        are guarded inside the ShouldProcess block and are fully skipped when -WhatIf is used.
        If a runtime error occurs after the schedule (step 2) or call flow (step 3) is created
        but before the final Set-CsAutoAttendant commit succeeds, those objects are orphaned.
        The error message includes the orphaned object IDs for manual cleanup via
        Get-CsOnlineSchedule / Remove-CsOnlineSchedule.
    .PARAMETER Identity
        Identity of the target Auto Attendant to add the holiday to.
    .PARAMETER HolidayName
        Name of the holiday to find and copy.
    .PARAMETER Headless
        Suppress progress bar while scanning source AAs.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [switch]$Headless
    )

    # -- Find the source holiday on any AA --
    if (-not $Headless) {
        Write-Host "Scanning Auto Attendants for source holiday '$HolidayName'..." -ForegroundColor Cyan
    }
    try {
        $allAAs = @(Invoke-TeamsRetry { Get-CsAutoAttendant } 'Get-CsAutoAttendant (tenant scan)')
    } catch {
        Write-Error "Failed to retrieve Auto Attendants: $_"
        throw
    }

    $sourceCtx  = $null
    $aaCount    = $allAAs.Count
    $i          = 0
    foreach ($aa in $allAAs) {
        $i++
        if (-not $Headless) {
            Write-Progress -Activity "Scanning for holiday '$HolidayName'" `
                           -Status "$i of $aaCount : $($aa.Name)" `
                           -PercentComplete (($i / $aaCount) * 100)
        }
        # Skip the target AA -- we are adding to it, not reading from it
        if ($aa.Identity -eq $Identity) { continue }
        $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
        if ($ctx) { $sourceCtx = $ctx; break }
    }
    if (-not $Headless) { Write-Progress -Activity "Scanning for holiday '$HolidayName'" -Completed }

    if (-not $sourceCtx) {
        throw "Holiday '$HolidayName' was not found on any Auto Attendant. Create it first with New-TeamsIVRHoliday."
    }

    # -- Load target AA --
    try {
        $targetAA = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve target Auto Attendant '$Identity': $_"
        throw
    }

    $existingOnTarget = Find-TeamsIVRHolidayContext -AA $targetAA -HolidayName $HolidayName
    if ($existingOnTarget) {
        throw "Holiday '$HolidayName' already exists on '$($targetAA.Name)'. Nothing was changed."
    }

    if ($PSCmdlet.ShouldProcess($targetAA.Name, "Add holiday '$HolidayName' (copied from source AA)")) {

        # Track object IDs so we can include them in the error message if the final
        # Set-CsAutoAttendant commit fails -- the operator needs them to manually
        # remove any orphaned schedule/call-flow objects.
        $createdScheduleId = $null
        $createdCallFlowId = $null

        # -- Step 1: Build date ranges --
        Write-Verbose "[$Script:RunId] Step 1: Copying date ranges for '$HolidayName'..."
        $dateRanges = @()
        if ($sourceCtx.Schedule.FixedSchedule -and $sourceCtx.Schedule.FixedSchedule.DateTimeRanges) {
            foreach ($r in $sourceCtx.Schedule.FixedSchedule.DateTimeRanges) {
                try {
                    $dateRanges += New-CsOnlineDateTimeRange -Start $r.Start -End $r.End
                } catch {
                    Write-Warning "Could not copy date range ($($r.Start) -- $($r.End)): $_"
                }
            }
        }
        Write-Verbose "[$Script:RunId] Step 1 complete: $($dateRanges.Count) date range(s) prepared."

        # -- Step 2: Create schedule --
        Write-Verbose "[$Script:RunId] Step 2: Creating schedule object for '$HolidayName'..."
        try {
            if ($dateRanges.Count -gt 0) {
                $newSchedule = New-CsOnlineSchedule -Name $HolidayName -FixedSchedule -DateTimeRanges $dateRanges
            } else {
                $newSchedule = New-CsOnlineSchedule -Name $HolidayName -FixedSchedule
            }
            $createdScheduleId = $newSchedule.Id
        } catch {
            throw "Step 2 failed -- could not create schedule for '$HolidayName': $_"
        }
        Write-Verbose "[$Script:RunId] Step 2 complete: schedule ID $createdScheduleId."

        # -- Step 3: Create call flow --
        Write-Verbose "[$Script:RunId] Step 3: Creating default disconnect call flow..."
        try {
            $disconnectOption = New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
            $menu             = New-CsAutoAttendantMenu -Name "$HolidayName Menu" -MenuOptions @($disconnectOption)
            $newCallFlow      = New-CsAutoAttendantCallFlow -Name "$HolidayName Flow" -Menu $menu
            $createdCallFlowId = $newCallFlow.Id
        } catch {
            throw "Step 3 failed -- could not create call flow for '$HolidayName'. Orphaned schedule ID: $createdScheduleId. Error: $_"
        }
        Write-Verbose "[$Script:RunId] Step 3 complete: call flow ID $createdCallFlowId."

        # -- Step 4: Create association --
        Write-Verbose "[$Script:RunId] Step 4: Creating CallHandlingAssociation..."
        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation `
                -Type Holiday `
                -ScheduleId $newSchedule.Id `
                -CallFlowId $newCallFlow.Id
        } catch {
            throw "Step 4 failed -- could not create association for '$HolidayName'. Orphaned schedule ID: $createdScheduleId, call flow ID: $createdCallFlowId. Error: $_"
        }
        Write-Verbose "[$Script:RunId] Step 4 complete."

        # -- Step 5: Commit to Teams --
        Write-Verbose "[$Script:RunId] Step 5: Committing changes to '$($targetAA.Name)'..."
        $targetAA.Schedules              = @($targetAA.Schedules)              + @($newSchedule)
        $targetAA.CallFlows              = @($targetAA.CallFlows)              + @($newCallFlow)
        $targetAA.CallHandlingAssociations = @($targetAA.CallHandlingAssociations) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $targetAA -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($targetAA.Name)'"
            Write-Verbose "[$Script:RunId] Step 5 complete: holiday '$HolidayName' committed to '$($targetAA.Name)'. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            # The AA was not saved. The schedule and call flow objects created in steps 2-3
            # are now orphaned. Include their IDs so the operator can clean them up.
            Write-Error "Step 5 failed -- could not save '$($targetAA.Name)'. Orphaned objects were created and must be manually removed: schedule ID=$createdScheduleId, call flow ID=$createdCallFlowId. Error: $_"
            throw
        }
        Write-Verbose "[$Script:RunId] Holiday '$HolidayName' successfully added to '$($targetAA.Name)'."
    }
}

#endregion

#region Set-TeamsIVRHolidayGreeting

function Set-TeamsIVRHolidayGreeting {
    <#
    .SYNOPSIS
        Sets the greeting message played when an IVR enters holiday mode.
    .DESCRIPTION
        Locates the holiday call flow on the specified Auto Attendant for the named
        holiday and replaces its greeting prompt with a new text-to-speech string.
        The existing routing-options message and menu options are preserved.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER HolidayName
        The holiday whose call flow greeting should be updated.
    .PARAMETER GreetingText
        Text-to-speech string to use as the greeting. Maximum 1000 characters
        (Teams TTS limit). Use plain ASCII; avoid em dashes and special Unicode.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 1000)]
        [string]$GreetingText
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'. Create it first with New-TeamsIVRHoliday."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set greeting on holiday '$HolidayName'")) {

        try {
            $newPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $GreetingText
        } catch {
            throw "Failed to create TTS prompt: $_"
        }

        # Rebuild the call flow with the new greeting, preserving the existing menu
        try {
            $newCallFlow = New-CsAutoAttendantCallFlow `
                -Name $ctx.CallFlow.Name `
                -Greetings @($newPrompt) `
                -Menu $ctx.CallFlow.Menu
        } catch {
            throw "Failed to create replacement call flow for '$HolidayName': $_"
        }

        # New association: same schedule ID, new call flow ID
        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation `
                -Type Holiday `
                -ScheduleId $ctx.Assoc.ScheduleId `
                -CallFlowId $newCallFlow.Id
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        $aa.CallFlows = @($aa.CallFlows | Where-Object { $_.Id -ne $ctx.CallFlow.Id }) + @($newCallFlow)
        $aa.CallHandlingAssociations = @(
            $aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }
        ) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Greeting updated for holiday '$HolidayName' on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRHolidayRoutingMessage

function Set-TeamsIVRHolidayRoutingMessage {
    <#
    .SYNOPSIS
        Sets the routing-options prompt for an IVR/holiday pair.
    .DESCRIPTION
        Updates the menu prompt (the message callers hear before key-press options are
        read) on the holiday call flow's menu for the given Auto Attendant and holiday.
        The existing greeting and menu options are preserved.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER HolidayName
        The holiday whose routing-options prompt should be updated.
    .PARAMETER RoutingMessage
        Text-to-speech string to use as the routing-options prompt. Maximum 1000 characters
        (Teams TTS limit). Use plain ASCII; avoid em dashes and special Unicode.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 1000)]
        [string]$RoutingMessage
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'. Create it first with New-TeamsIVRHoliday."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set routing-options message on holiday '$HolidayName'")) {

        try {
            $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $RoutingMessage
        } catch {
            throw "Failed to create TTS prompt: $_"
        }

        # Rebuild the menu with the new prompt, preserving the existing menu options
        try {
            $existingMenu = $ctx.CallFlow.Menu
            $newMenu = New-CsAutoAttendantMenu `
                -Name $existingMenu.Name `
                -Prompts @($menuPrompt) `
                -MenuOptions @($existingMenu.MenuOptions)
        } catch {
            throw "Failed to create replacement menu for '$HolidayName': $_"
        }

        # Rebuild the call flow with the new menu, preserving any existing greeting
        try {
            $existingGreetings = @($ctx.CallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $ctx.CallFlow.Name `
                    -Greetings $existingGreetings `
                    -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $ctx.CallFlow.Name `
                    -Menu $newMenu
            }
        } catch {
            throw "Failed to create replacement call flow for '$HolidayName': $_"
        }

        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation `
                -Type Holiday `
                -ScheduleId $ctx.Assoc.ScheduleId `
                -CallFlowId $newCallFlow.Id
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        $aa.CallFlows = @($aa.CallFlows | Where-Object { $_.Id -ne $ctx.CallFlow.Id }) + @($newCallFlow)
        $aa.CallHandlingAssociations = @(
            $aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }
        ) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Routing-options message updated for holiday '$HolidayName' on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRHolidayRouting

function Set-TeamsIVRHolidayRouting {
    <#
    .SYNOPSIS
        Sets the call routing action for an IVR/holiday pair.
    .DESCRIPTION
        Replaces the routing action on the holiday call flow for the specified AA and holiday.
        Supported RoutingAction values:
          Disconnect          -- Hang up the call. No -RedirectTarget needed.
          RedirectToUser      -- Transfer to a Teams user. -RedirectTarget: UPN or object ID.
          RedirectToPSTN      -- Transfer to a PSTN number. -RedirectTarget: E.164 (e.g. +15551234567).
          RedirectToVoicemail -- Send to a shared voicemail. -RedirectTarget: UPN or object ID.
          RedirectToAA        -- Transfer to another Auto Attendant. -RedirectTarget: AA identity (GUID or name).

        The existing greeting and routing-options message are preserved -- only the
        menu option (the routing action) is replaced.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER HolidayName
        The holiday whose routing action should be updated.
    .PARAMETER RoutingAction
        The action to take. See the Description for valid values.
    .PARAMETER RedirectTarget
        Required for all RoutingAction values except 'Disconnect'.
        UPN / object ID (User, Voicemail), E.164 number (PSTN), or AA identity (AA).
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Disconnect', 'RedirectToUser', 'RedirectToPSTN', 'RedirectToVoicemail', 'RedirectToAA')]
        [string]$RoutingAction,

        [Parameter(Mandatory = $false)]
        [string]$RedirectTarget
    )

    # Validate that RedirectTarget is supplied for all non-Disconnect actions
    if ($RoutingAction -ne 'Disconnect' -and [string]::IsNullOrWhiteSpace($RedirectTarget)) {
        throw "-RedirectTarget is required when -RoutingAction is '$RoutingAction'."
    }

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'. Create it first with New-TeamsIVRHoliday."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set routing action '$RoutingAction' on holiday '$HolidayName'")) {

        # Build the new menu option
        try {
            $menuOption = switch ($RoutingAction) {

                'Disconnect' {
                    New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                }

                'RedirectToUser' {
                    $resolvedUser = Resolve-TeamsUser -User $RedirectTarget
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $resolvedUser.Identity `
                        -Type User
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }

                'RedirectToPSTN' {
                    Assert-E164Format -PhoneNumber $RedirectTarget
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $RedirectTarget `
                        -Type ExternalPstn
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }

                'RedirectToVoicemail' {
                    $resolvedUser = Resolve-TeamsUser -User $RedirectTarget
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $resolvedUser.Identity `
                        -Type SharedVoicemail
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }

                'RedirectToAA' {
                    try {
                        $targetAA = Get-CsAutoAttendant -Identity $RedirectTarget
                    } catch {
                        throw "Could not find target Auto Attendant '$RedirectTarget': $_"
                    }
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $targetAA.Identity `
                        -Type ApplicationEndpoint
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }
            }
        } catch {
            throw "Failed to build menu option for RoutingAction '$RoutingAction': $_"
        }

        # Rebuild menu: new option, preserve any existing prompt
        try {
            $existingMenu    = $ctx.CallFlow.Menu
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu `
                    -Name $existingMenu.Name `
                    -Prompts $existingPrompts `
                    -MenuOptions @($menuOption)
            } else {
                $newMenu = New-CsAutoAttendantMenu `
                    -Name $existingMenu.Name `
                    -MenuOptions @($menuOption)
            }
        } catch {
            throw "Failed to rebuild menu for '$HolidayName': $_"
        }

        # Rebuild call flow: new menu, preserve any existing greeting
        try {
            $existingGreetings = @($ctx.CallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $ctx.CallFlow.Name `
                    -Greetings $existingGreetings `
                    -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $ctx.CallFlow.Name `
                    -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild call flow for '$HolidayName': $_"
        }

        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation `
                -Type Holiday `
                -ScheduleId $ctx.Assoc.ScheduleId `
                -CallFlowId $newCallFlow.Id
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        $aa.CallFlows = @($aa.CallFlows | Where-Object { $_.Id -ne $ctx.CallFlow.Id }) + @($newCallFlow)
        $aa.CallHandlingAssociations = @(
            $aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }
        ) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Routing action '$RoutingAction' set for holiday '$HolidayName' on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRGreeting

function Set-TeamsIVRGreeting {
    <#
    .SYNOPSIS
        Sets the greeting message on an Auto Attendant's default (business-hours) call flow.
    .DESCRIPTION
        Locates the DefaultCallFlow on the specified Auto Attendant and replaces its
        greeting prompt with a new text-to-speech string. The DefaultCallFlow is the
        call flow used during business hours (i.e. outside AfterHours and Holiday
        windows). The existing menu options are preserved.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER GreetingText
        Text-to-speech string to use as the greeting. Maximum 1000 characters
        (Teams TTS limit). Use plain ASCII; avoid em dashes and special Unicode.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 1000)]
        [string]$GreetingText
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set greeting on default call flow")) {

        try {
            $newPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $GreetingText
        } catch {
            throw "Failed to create TTS prompt: $_"
        }

        # Rebuild the default call flow with the new greeting, preserving the existing menu
        try {
            $newCallFlow = New-CsAutoAttendantCallFlow `
                -Name $aa.DefaultCallFlow.Name `
                -Greetings @($newPrompt) `
                -Menu $aa.DefaultCallFlow.Menu
        } catch {
            throw "Failed to create replacement default call flow: $_"
        }

        $aa.DefaultCallFlow = $newCallFlow

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Greeting updated on default call flow for '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRRoutingMessage

function Set-TeamsIVRRoutingMessage {
    <#
    .SYNOPSIS
        Sets the routing-options prompt on an Auto Attendant's default call flow menu.
    .DESCRIPTION
        Locates the DefaultCallFlow on the specified Auto Attendant, replaces the
        menu's TTS prompt with a new string, and preserves existing menu options and
        the call flow's greeting. This is the message callers hear before key-press
        options are read (e.g. "Press 1 for Sales, press 2 for Support").
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER RoutingMessage
        Text-to-speech string to use as the routing-options prompt. Maximum 1000 characters
        (Teams TTS limit). Use plain ASCII; avoid em dashes and special Unicode.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 1000)]
        [string]$RoutingMessage
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set routing-options message on default call flow")) {

        try {
            $menuPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $RoutingMessage
        } catch {
            throw "Failed to create TTS prompt: $_"
        }

        # Rebuild the menu with the new prompt, preserving existing menu options
        try {
            $existingMenu = $aa.DefaultCallFlow.Menu
            $newMenu = New-CsAutoAttendantMenu `
                -Name $existingMenu.Name `
                -Prompts @($menuPrompt) `
                -MenuOptions @($existingMenu.MenuOptions)
        } catch {
            throw "Failed to create replacement menu: $_"
        }

        # Rebuild the call flow with the new menu, preserving any existing greeting
        try {
            $existingGreetings = @($aa.DefaultCallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $aa.DefaultCallFlow.Name `
                    -Greetings $existingGreetings `
                    -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $aa.DefaultCallFlow.Name `
                    -Menu $newMenu
            }
        } catch {
            throw "Failed to create replacement default call flow: $_"
        }

        $aa.DefaultCallFlow = $newCallFlow

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Routing-options message updated on default call flow for '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRRouting

function Set-TeamsIVRRouting {
    <#
    .SYNOPSIS
        Sets the call routing action on an Auto Attendant's default call flow.
    .DESCRIPTION
        Locates the DefaultCallFlow on the specified Auto Attendant and replaces the
        automatic menu option with a new routing action.
        Supported RoutingAction values mirror Set-TeamsIVRHolidayRouting:
          Disconnect          -- Hang up the call.
          RedirectToUser      -- Transfer to a Teams user. -RedirectTarget: UPN or object ID.
          RedirectToPSTN      -- Transfer to a PSTN number. -RedirectTarget: E.164.
          RedirectToVoicemail -- Send to shared voicemail. -RedirectTarget: UPN or object ID.
          RedirectToAA        -- Transfer to another Auto Attendant. -RedirectTarget: AA identity.
        The existing greeting and routing-options message are preserved.
        Note: The DefaultCallFlow is the business-hours call flow. AfterHours routing
        is controlled by the schedule (Set-TeamsIVRSchedule), not this function.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER RoutingAction
        The routing action to apply. See Description for valid values.
    .PARAMETER RedirectTarget
        Required for all RoutingAction values except 'Disconnect'.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Disconnect', 'RedirectToUser', 'RedirectToPSTN', 'RedirectToVoicemail', 'RedirectToAA')]
        [string]$RoutingAction,

        [Parameter(Mandatory = $false)]
        [string]$RedirectTarget
    )

    # Validate that RedirectTarget is supplied for all non-Disconnect actions
    if ($RoutingAction -ne 'Disconnect' -and [string]::IsNullOrWhiteSpace($RedirectTarget)) {
        throw "-RedirectTarget is required when -RoutingAction is '$RoutingAction'."
    }

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Set routing action '$RoutingAction' on default call flow")) {

        # Build the new menu option
        try {
            $menuOption = switch ($RoutingAction) {

                'Disconnect' {
                    New-CsAutoAttendantMenuOption -Action DisconnectCall -DtmfResponse Automatic
                }

                'RedirectToUser' {
                    $resolvedUser = Resolve-TeamsUser -User $RedirectTarget
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $resolvedUser.Identity `
                        -Type User
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }

                'RedirectToPSTN' {
                    Assert-E164Format -PhoneNumber $RedirectTarget
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $RedirectTarget `
                        -Type ExternalPstn
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }

                'RedirectToVoicemail' {
                    $resolvedUser = Resolve-TeamsUser -User $RedirectTarget
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $resolvedUser.Identity `
                        -Type SharedVoicemail
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }

                'RedirectToAA' {
                    try {
                        $targetAA = Get-CsAutoAttendant -Identity $RedirectTarget
                    } catch {
                        throw "Could not find target Auto Attendant '$RedirectTarget': $_"
                    }
                    $entity = New-CsAutoAttendantCallableEntity `
                        -Identity $targetAA.Identity `
                        -Type ApplicationEndpoint
                    New-CsAutoAttendantMenuOption `
                        -Action TransferCallToTarget `
                        -DtmfResponse Automatic `
                        -CallTarget $entity
                }
            }
        } catch {
            throw "Failed to build menu option for RoutingAction '$RoutingAction': $_"
        }

        # Rebuild menu: new option, preserve any existing prompt
        try {
            $existingMenu    = $aa.DefaultCallFlow.Menu
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu `
                    -Name $existingMenu.Name `
                    -Prompts $existingPrompts `
                    -MenuOptions @($menuOption)
            } else {
                $newMenu = New-CsAutoAttendantMenu `
                    -Name $existingMenu.Name `
                    -MenuOptions @($menuOption)
            }
        } catch {
            throw "Failed to rebuild default call flow menu: $_"
        }

        # Rebuild call flow: new menu, preserve any existing greeting
        try {
            $existingGreetings = @($aa.DefaultCallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $aa.DefaultCallFlow.Name `
                    -Greetings $existingGreetings `
                    -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow `
                    -Name $aa.DefaultCallFlow.Name `
                    -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild default call flow: $_"
        }

        $aa.DefaultCallFlow = $newCallFlow

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Routing action '$RoutingAction' set on default call flow for '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Get-TeamsIVRHolidayMenuOptions

function Get-TeamsIVRHolidayMenuOptions {
    <#
    .SYNOPSIS
        Returns the keypress menu options for a named holiday's call flow on an Auto Attendant.
    .DESCRIPTION
        Retrieves the CsAutoAttendantMenu from the holiday call flow and converts each
        menu option into a PSCustomObject with properties Key, Action, and Target.
        The same data is also available as a JSON string (via Format-MenuOptionsAsJson) for
        copy-paste into -SetHolidayMenu.
    .PARAMETER Identity
        Identity of the Auto Attendant to query.
    .PARAMETER HolidayName
        Name of the holiday whose menu options should be retrieved.
    .OUTPUTS
        PSCustomObject[] with properties: Key, Action, Target
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'. Create it first with New-TeamsIVRHoliday."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    $specs = @($ctx.CallFlow.Menu.MenuOptions | ForEach-Object { Convert-MenuOptionToActionSpec -MenuOption $_ })
    return $specs
}

#endregion

#region Add-TeamsIVRHolidayMenuOption

function Add-TeamsIVRHolidayMenuOption {
    <#
    .SYNOPSIS
        Adds or replaces a single keypress option on a holiday call flow menu.
    .DESCRIPTION
        Binds the specified key (0-9, *, #) to the given routing action on the
        holiday call flow for the named AA/holiday pair.

        If the menu already has an Automatic routing option (from Set-TeamsIVRHolidayRouting),
        it is removed with a warning -- an Automatic option and keypress options cannot coexist.

        If the same key already has a binding, it is replaced.
        All other existing keypress bindings are preserved.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER HolidayName
        Name of the holiday whose menu should be updated.
    .PARAMETER Key
        The key callers press to trigger this option. Valid: 0-9, *, #
    .PARAMETER RoutingAction
        The action to take when this key is pressed.
    .PARAMETER RedirectTarget
        Required for all RoutingAction values except 'Disconnect'.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('0','1','2','3','4','5','6','7','8','9','*','#')]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Disconnect','RedirectToUser','RedirectToPSTN','RedirectToVoicemail','RedirectToAA')]
        [string]$RoutingAction,

        [Parameter(Mandatory = $false)]
        [string]$RedirectTarget
    )

    if ($RoutingAction -ne 'Disconnect' -and [string]::IsNullOrWhiteSpace($RedirectTarget)) {
        throw "-RedirectTarget is required when -RoutingAction is '$RoutingAction'."
    }

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'. Create it first with New-TeamsIVRHoliday."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    $tone = Convert-DtmfKeyToTone -Key $Key

    if ($PSCmdlet.ShouldProcess($aa.Name, "Add key '$Key' ($RoutingAction) to holiday '$HolidayName' menu")) {

        try {
            $newOption = Build-MenuOptionFromSpec -Spec ([PSCustomObject]@{
                Key    = $Key
                Action = $RoutingAction
                Target = $RedirectTarget
            })
        } catch {
            throw "Failed to build menu option: $_"
        }

        $existingMenu = $ctx.CallFlow.Menu

        # Remove any Automatic option (incompatible with keypress options) and any existing binding on this key
        $hadAutomatic = @($existingMenu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -eq 'Automatic' }).Count -gt 0
        $filteredOptions = @($existingMenu.MenuOptions | Where-Object {
            $_.DtmfResponse.ToString() -ne 'Automatic' -and $_.DtmfResponse.ToString() -ne $tone
        })

        if ($hadAutomatic) {
            Write-Warning "Removed existing automatic routing from '$HolidayName' call flow -- incompatible with keypress menu. Use Set-TeamsIVRHolidayRouting to restore a single automatic action."
        }

        $newOptions = $filteredOptions + @($newOption)

        try {
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -Prompts $existingPrompts -MenuOptions $newOptions
            } else {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -MenuOptions $newOptions
            }
        } catch {
            throw "Failed to rebuild menu for '$HolidayName': $_"
        }

        try {
            $existingGreetings = @($ctx.CallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $ctx.CallFlow.Name -Greetings $existingGreetings -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $ctx.CallFlow.Name -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild call flow for '$HolidayName': $_"
        }

        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $ctx.Assoc.ScheduleId -CallFlowId $newCallFlow.Id
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        $aa.CallFlows = @($aa.CallFlows | Where-Object { $_.Id -ne $ctx.CallFlow.Id }) + @($newCallFlow)
        $aa.CallHandlingAssociations = @($aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Key '$Key' ($RoutingAction) added to '$HolidayName' menu on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Remove-TeamsIVRHolidayMenuOption

function Remove-TeamsIVRHolidayMenuOption {
    <#
    .SYNOPSIS
        Removes a single keypress option from a holiday call flow menu.
    .DESCRIPTION
        Finds and removes the binding for the specified key from the holiday call flow menu.
        All other options are preserved. If the key is not found, a warning is written and
        no change is made.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER HolidayName
        Name of the holiday whose menu should be updated.
    .PARAMETER Key
        The key to remove. Valid: 0-9, *, #
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('0','1','2','3','4','5','6','7','8','9','*','#')]
        [string]$Key
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    $tone = Convert-DtmfKeyToTone -Key $Key
    $existingMenu = $ctx.CallFlow.Menu
    $targetOption = @($existingMenu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -eq $tone })

    if ($targetOption.Count -eq 0) {
        throw "Key '$Key' is not bound in the '$HolidayName' menu on '$($aa.Name)'. Nothing was changed."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Remove key '$Key' from holiday '$HolidayName' menu")) {

        $remainingOptions = @($existingMenu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -ne $tone })

        try {
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -Prompts $existingPrompts -MenuOptions $remainingOptions
            } else {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -MenuOptions $remainingOptions
            }
        } catch {
            throw "Failed to rebuild menu for '$HolidayName': $_"
        }

        try {
            $existingGreetings = @($ctx.CallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $ctx.CallFlow.Name -Greetings $existingGreetings -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $ctx.CallFlow.Name -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild call flow for '$HolidayName': $_"
        }

        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $ctx.Assoc.ScheduleId -CallFlowId $newCallFlow.Id
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        $aa.CallFlows = @($aa.CallFlows | Where-Object { $_.Id -ne $ctx.CallFlow.Id }) + @($newCallFlow)
        $aa.CallHandlingAssociations = @($aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Key '$Key' removed from '$HolidayName' menu on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRHolidayMenu

function Set-TeamsIVRHolidayMenu {
    <#
    .SYNOPSIS
        Replaces all keypress options on a holiday call flow menu from a JSON definition.
    .DESCRIPTION
        Accepts a JSON array of menu option specs and rebuilds the entire menu on the holiday
        call flow. Existing greeting and routing-options prompt are preserved.

        JSON format:
          [{"Key":"1","Action":"RedirectToUser","Target":"user@contoso.com"},
           {"Key":"2","Action":"Disconnect"}]

        Valid Action values: Disconnect, RedirectToUser, RedirectToPSTN, RedirectToVoicemail, RedirectToAA
        Target is required for all actions except Disconnect.
        Use -GetHolidayMenuOptions to retrieve the current menu as copy-paste JSON.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER HolidayName
        Name of the holiday whose menu should be replaced.
    .PARAMETER MenuOptionsJson
        JSON array of menu option specs. See Description for format.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(1, 256)]
        [string]$HolidayName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$MenuOptionsJson
    )

    try {
        $specs = @($MenuOptionsJson | ConvertFrom-Json)
    } catch {
        throw "Failed to parse -MenuOptionsJson: $_`nExpected format: [{`"Key`":`"1`",`"Action`":`"RedirectToUser`",`"Target`":`"user@contoso.com`"}]"
    }

    if ($specs.Count -eq 0) {
        throw "-MenuOptionsJson parsed to an empty array. Provide at least one menu option."
    }

    # Validate each spec before building any Teams objects so the caller gets a
    # precise error (which key, what's wrong) rather than a deep stack trace.
    $validActions = @('Disconnect','RedirectToUser','RedirectToPSTN','RedirectToVoicemail','RedirectToAA')
    $validKeys    = @('0','1','2','3','4','5','6','7','8','9','*','#','(auto)')
    $valErrors    = [System.Collections.Generic.List[string]]::new()
    foreach ($spec in $specs) {
        if (-not $spec.Key -or $spec.Key -notin $validKeys) {
            [void]$valErrors.Add("Key '$($spec.Key)': invalid key -- must be 0-9, *, or #.")
        }
        if (-not $spec.Action -or $spec.Action -notin $validActions) {
            [void]$valErrors.Add("Key '$($spec.Key)': unknown Action '$($spec.Action)'.")
        }
        if ($spec.Action -ne 'Disconnect' -and [string]::IsNullOrWhiteSpace($spec.Target)) {
            [void]$valErrors.Add("Key '$($spec.Key)': Action '$($spec.Action)' requires a Target.")
        }
        if ($spec.Action -eq 'RedirectToPSTN' -and $spec.Target -and $spec.Target -notmatch '^\+[1-9]\d{6,14}$') {
            [void]$valErrors.Add("Key '$($spec.Key)': Target '$($spec.Target)' is not valid E.164 format.")
        }
    }
    if ($valErrors.Count -gt 0) {
        throw "Invalid menu option(s):`n  " + ($valErrors -join "`n  ")
    }

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $ctx = Find-TeamsIVRHolidayContext -AA $aa -HolidayName $HolidayName
    if (-not $ctx) {
        throw "Holiday '$HolidayName' not found on '$($aa.Name)'."
    }
    if (-not $ctx.CallFlow) {
        throw "Holiday '$HolidayName' on '$($aa.Name)' has no call flow. The AA may be in an inconsistent state."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Replace all menu options on holiday '$HolidayName' ($($specs.Count) option(s))")) {

        $newOptions = @()
        foreach ($spec in $specs) {
            try {
                $newOptions += Build-MenuOptionFromSpec -Spec $spec
            } catch {
                throw "Failed to build option for key '$($spec.Key)': $_"
            }
        }

        $existingMenu = $ctx.CallFlow.Menu
        try {
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -Prompts $existingPrompts -MenuOptions $newOptions
            } else {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -MenuOptions $newOptions
            }
        } catch {
            throw "Failed to rebuild menu for '$HolidayName': $_"
        }

        try {
            $existingGreetings = @($ctx.CallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $ctx.CallFlow.Name -Greetings $existingGreetings -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $ctx.CallFlow.Name -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild call flow for '$HolidayName': $_"
        }

        try {
            $newAssoc = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId $ctx.Assoc.ScheduleId -CallFlowId $newCallFlow.Id
        } catch {
            throw "Failed to create replacement CallHandlingAssociation for '$HolidayName': $_"
        }

        $aa.CallFlows = @($aa.CallFlows | Where-Object { $_.Id -ne $ctx.CallFlow.Id }) + @($newCallFlow)
        $aa.CallHandlingAssociations = @($aa.CallHandlingAssociations | Where-Object { $_.ScheduleId -ne $ctx.Assoc.ScheduleId }) + @($newAssoc)

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Menu replaced on '$HolidayName' on '$($aa.Name)' ($($specs.Count) option(s))"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Get-TeamsIVRMenuOptions

function Get-TeamsIVRMenuOptions {
    <#
    .SYNOPSIS
        Returns the keypress menu options for an Auto Attendant's default call flow.
    .PARAMETER Identity
        Identity of the Auto Attendant to query.
    .OUTPUTS
        PSCustomObject[] with properties: Key, Action, Target
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $specs = @($aa.DefaultCallFlow.Menu.MenuOptions | ForEach-Object { Convert-MenuOptionToActionSpec -MenuOption $_ })
    return $specs
}

#endregion

#region Add-TeamsIVRMenuOption

function Add-TeamsIVRMenuOption {
    <#
    .SYNOPSIS
        Adds or replaces a single keypress option on an Auto Attendant's default call flow menu.
    .DESCRIPTION
        Same behaviour as Add-TeamsIVRHolidayMenuOption but targets DefaultCallFlow.
        If the menu currently has an Automatic option it is removed with a warning.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER Key
        The key callers press to trigger this option. Valid: 0-9, *, #
    .PARAMETER RoutingAction
        The action to take when this key is pressed.
    .PARAMETER RedirectTarget
        Required for all RoutingAction values except 'Disconnect'.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateSet('0','1','2','3','4','5','6','7','8','9','*','#')]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Disconnect','RedirectToUser','RedirectToPSTN','RedirectToVoicemail','RedirectToAA')]
        [string]$RoutingAction,

        [Parameter(Mandatory = $false)]
        [string]$RedirectTarget
    )

    if ($RoutingAction -ne 'Disconnect' -and [string]::IsNullOrWhiteSpace($RedirectTarget)) {
        throw "-RedirectTarget is required when -RoutingAction is '$RoutingAction'."
    }

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $tone = Convert-DtmfKeyToTone -Key $Key

    if ($PSCmdlet.ShouldProcess($aa.Name, "Add key '$Key' ($RoutingAction) to default call flow menu")) {

        try {
            $newOption = Build-MenuOptionFromSpec -Spec ([PSCustomObject]@{
                Key    = $Key
                Action = $RoutingAction
                Target = $RedirectTarget
            })
        } catch {
            throw "Failed to build menu option: $_"
        }

        $existingMenu = $aa.DefaultCallFlow.Menu
        $hadAutomatic = @($existingMenu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -eq 'Automatic' }).Count -gt 0
        $filteredOptions = @($existingMenu.MenuOptions | Where-Object {
            $_.DtmfResponse.ToString() -ne 'Automatic' -and $_.DtmfResponse.ToString() -ne $tone
        })

        if ($hadAutomatic) {
            Write-Warning "Removed existing automatic routing from default call flow -- incompatible with keypress menu. Use Set-TeamsIVRRouting to restore a single automatic action."
        }

        $newOptions = $filteredOptions + @($newOption)

        try {
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -Prompts $existingPrompts -MenuOptions $newOptions
            } else {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -MenuOptions $newOptions
            }
        } catch {
            throw "Failed to rebuild default call flow menu: $_"
        }

        try {
            $existingGreetings = @($aa.DefaultCallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $aa.DefaultCallFlow.Name -Greetings $existingGreetings -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $aa.DefaultCallFlow.Name -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild default call flow: $_"
        }

        $aa.DefaultCallFlow = $newCallFlow

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Key '$Key' ($RoutingAction) added to default call flow menu on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Remove-TeamsIVRMenuOption

function Remove-TeamsIVRMenuOption {
    <#
    .SYNOPSIS
        Removes a single keypress option from an Auto Attendant's default call flow menu.
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER Key
        The key to remove. Valid: 0-9, *, #
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateSet('0','1','2','3','4','5','6','7','8','9','*','#')]
        [string]$Key
    )

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    $tone = Convert-DtmfKeyToTone -Key $Key
    $existingMenu = $aa.DefaultCallFlow.Menu
    $targetOption = @($existingMenu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -eq $tone })

    if ($targetOption.Count -eq 0) {
        throw "Key '$Key' is not bound in the default call flow menu on '$($aa.Name)'. Nothing was changed."
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Remove key '$Key' from default call flow menu")) {

        $remainingOptions = @($existingMenu.MenuOptions | Where-Object { $_.DtmfResponse.ToString() -ne $tone })

        try {
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -Prompts $existingPrompts -MenuOptions $remainingOptions
            } else {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -MenuOptions $remainingOptions
            }
        } catch {
            throw "Failed to rebuild default call flow menu: $_"
        }

        try {
            $existingGreetings = @($aa.DefaultCallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $aa.DefaultCallFlow.Name -Greetings $existingGreetings -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $aa.DefaultCallFlow.Name -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild default call flow: $_"
        }

        $aa.DefaultCallFlow = $newCallFlow

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Key '$Key' removed from default call flow menu on '$($aa.Name)'"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
}

#endregion

#region Set-TeamsIVRMenu

function Set-TeamsIVRMenu {
    <#
    .SYNOPSIS
        Replaces all keypress options on an Auto Attendant's default call flow menu from a JSON definition.
    .DESCRIPTION
        Accepts a JSON array of menu option specs and rebuilds the entire default call flow menu.
        Existing greeting and routing-options prompt are preserved.
        Use -GetIVRMenuOptions to retrieve the current menu as copy-paste JSON.

        JSON format:
          [{"Key":"1","Action":"RedirectToUser","Target":"user@contoso.com"},
           {"Key":"2","Action":"Disconnect"}]
    .PARAMETER Identity
        Identity of the Auto Attendant to update.
    .PARAMETER MenuOptionsJson
        JSON array of menu option specs. See Description for format.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$MenuOptionsJson
    )

    try {
        $specs = @($MenuOptionsJson | ConvertFrom-Json)
    } catch {
        throw "Failed to parse -MenuOptionsJson: $_`nExpected format: [{`"Key`":`"1`",`"Action`":`"RedirectToUser`",`"Target`":`"user@contoso.com`"}]"
    }

    if ($specs.Count -eq 0) {
        throw "-MenuOptionsJson parsed to an empty array. Provide at least one menu option."
    }

    # Validate each spec before building any Teams objects so the caller gets a
    # precise error (which key, what's wrong) rather than a deep stack trace.
    $validActions = @('Disconnect','RedirectToUser','RedirectToPSTN','RedirectToVoicemail','RedirectToAA')
    $validKeys    = @('0','1','2','3','4','5','6','7','8','9','*','#','(auto)')
    $valErrors    = [System.Collections.Generic.List[string]]::new()
    foreach ($spec in $specs) {
        if (-not $spec.Key -or $spec.Key -notin $validKeys) {
            [void]$valErrors.Add("Key '$($spec.Key)': invalid key -- must be 0-9, *, or #.")
        }
        if (-not $spec.Action -or $spec.Action -notin $validActions) {
            [void]$valErrors.Add("Key '$($spec.Key)': unknown Action '$($spec.Action)'.")
        }
        if ($spec.Action -ne 'Disconnect' -and [string]::IsNullOrWhiteSpace($spec.Target)) {
            [void]$valErrors.Add("Key '$($spec.Key)': Action '$($spec.Action)' requires a Target.")
        }
        if ($spec.Action -eq 'RedirectToPSTN' -and $spec.Target -and $spec.Target -notmatch '^\+[1-9]\d{6,14}$') {
            [void]$valErrors.Add("Key '$($spec.Key)': Target '$($spec.Target)' is not valid E.164 format.")
        }
    }
    if ($valErrors.Count -gt 0) {
        throw "Invalid menu option(s):`n  " + ($valErrors -join "`n  ")
    }

    try {
        $aa = Get-CsAutoAttendant -Identity $Identity
    } catch {
        Write-Error "Failed to retrieve Auto Attendant '$Identity': $_"
        throw
    }

    if ($PSCmdlet.ShouldProcess($aa.Name, "Replace all default call flow menu options ($($specs.Count) option(s))")) {

        $newOptions = @()
        foreach ($spec in $specs) {
            try {
                $newOptions += Build-MenuOptionFromSpec -Spec $spec
            } catch {
                throw "Failed to build option for key '$($spec.Key)': $_"
            }
        }

        $existingMenu = $aa.DefaultCallFlow.Menu
        try {
            $existingPrompts = @($existingMenu.Prompts)
            if ($existingPrompts.Count -gt 0) {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -Prompts $existingPrompts -MenuOptions $newOptions
            } else {
                $newMenu = New-CsAutoAttendantMenu -Name $existingMenu.Name -MenuOptions $newOptions
            }
        } catch {
            throw "Failed to rebuild default call flow menu: $_"
        }

        try {
            $existingGreetings = @($aa.DefaultCallFlow.Greetings)
            if ($existingGreetings.Count -gt 0) {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $aa.DefaultCallFlow.Name -Greetings $existingGreetings -Menu $newMenu
            } else {
                $newCallFlow = New-CsAutoAttendantCallFlow -Name $aa.DefaultCallFlow.Name -Menu $newMenu
            }
        } catch {
            throw "Failed to rebuild default call flow: $_"
        }

        $aa.DefaultCallFlow = $newCallFlow

        try {
            Invoke-TeamsRetry { Set-CsAutoAttendant -Instance $aa -WhatIf:$WhatIfPreference } "Set-CsAutoAttendant '$($aa.Name)'"
            Write-Verbose "Default call flow menu replaced on '$($aa.Name)' ($($specs.Count) option(s))"
            Write-Verbose "[$Script:RunId] Committed. Note: Teams changes propagate asynchronously and may take up to 60 seconds to take effect."
        } catch {
            Write-Error "Failed to save Auto Attendant '$($aa.Name)': $_"
            throw
        }
    }
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
    Write-Host ""
    Write-Host "  HOLIDAY ACTIONS" -ForegroundColor Yellow
    Write-Host "  ---------------"
    Write-Host "    -GetHolidayList"
    Write-Host "        List all holiday schedules found across every Auto Attendant."
    Write-Host "        Output per holiday: name, schedule ID, linked IVR(s), and date ranges."
    Write-Host "        Note: Teams stores holidays inside AA objects, not at the tenant level."
    Write-Host "        A schedule shared by multiple AAs appears once with all linked IVRs listed."
    Write-Host ""
    Write-Host "    -GetHoliday -HolidayName <string>"
    Write-Host "        Find a named holiday across all Auto Attendants."
    Write-Host "        Output: IVR name/ID, schedule ID, and date ranges for each AA it appears on."
    Write-Host ""
    Write-Host "    -NewHoliday -Identity <string> -HolidayName <string>"
    Write-Host "        Create a new holiday on an Auto Attendant (Teams requires an AA -- no tenant-level store)."
    Write-Host "        Creates an empty schedule and a default disconnect call flow."
    Write-Host "        Follow up with -SetHolidayTime, -SetHolidayGreeting, and -SetHolidayRouting."
    Write-Host ""
    Write-Host "    -SetHolidayTime -Identity <string> -HolidayName <string> -StartDateTime <string> -EndDateTime <string>"
    Write-Host "        Set the active date/time range on a holiday. Replaces any existing range."
    Write-Host "        Accepts any .NET-parseable date/time string; converted to Teams format internally."
    Write-Host "        RESTRICTION: Both times must fall on 15-minute boundaries (minutes 0, 15, 30, or 45)."
    Write-Host "        Example: -StartDateTime '2025-12-25 00:00' -EndDateTime '2025-12-26 00:00'"
    Write-Host "        (Use next-day midnight to cover a full day -- '23:59' is not boundary-aligned.)"
    Write-Host ""
    Write-Host "    -AttachHolidayToIVR -Identity <string> -HolidayName <string>"
    Write-Host "        Copy an existing holiday (found on any AA) onto the specified Auto Attendant."
    Write-Host "        Date ranges are copied; greeting and routing must be configured separately."
    Write-Host ""
    Write-Host "    -SetHolidayGreeting -Identity <string> -HolidayName <string> -GreetingText <string>"
    Write-Host "        Set the TTS greeting callers hear when the IVR is in holiday mode."
    Write-Host ""
    Write-Host "    -SetHolidayRoutingMessage -Identity <string> -HolidayName <string> -RoutingMessage <string>"
    Write-Host "        Set the TTS menu prompt played after the greeting (before key-press options)."
    Write-Host "        Useful when the holiday call flow offers menu choices; skip if just disconnecting."
    Write-Host ""
    Write-Host "    -SetHolidayRouting -Identity <string> -HolidayName <string> -RoutingAction <action> [-RedirectTarget <string>]"
    Write-Host "        Set what happens to calls during the holiday period (single automatic action)."
    Write-Host "        For multi-option key-press menus, use -AddHolidayMenuOption instead."
    Write-Host "        RoutingAction values:"
    Write-Host "          Disconnect        -- Hang up the call. No -RedirectTarget needed."
    Write-Host "          RedirectToUser    -- Transfer to a Teams user.  -RedirectTarget: UPN or object ID."
    Write-Host "          RedirectToPSTN    -- Transfer to a phone number. -RedirectTarget: E.164 (+15551234567)."
    Write-Host "          RedirectToVoicemail -- Send to shared voicemail. -RedirectTarget: UPN or object ID."
    Write-Host "          RedirectToAA      -- Transfer to another AA.    -RedirectTarget: AA identity (GUID or name)."
    Write-Host ""
    Write-Host "    -GetHolidayMenuOptions -Identity <string> -HolidayName <string>"
    Write-Host "        List all keypress options on a holiday call flow."
    Write-Host "        Prints a table (Key / Action / Target) and a JSON blob for use with -SetHolidayMenu."
    Write-Host ""
    Write-Host "    -AddHolidayMenuOption -Identity <string> -HolidayName <string> -Key <0-9|*|#> -RoutingAction <action> [-RedirectTarget <string>]"
    Write-Host "        Add or replace a single key binding on a holiday call flow menu."
    Write-Host "        Key: 0-9, *, or #   RoutingAction: same values as -SetHolidayRouting."
    Write-Host "        If an automatic routing option exists it is removed (incompatible with key-press)."
    Write-Host ""
    Write-Host "    -RemoveHolidayMenuOption -Identity <string> -HolidayName <string> -Key <0-9|*|#>"
    Write-Host "        Remove a single key binding from a holiday call flow menu."
    Write-Host ""
    Write-Host "    -SetHolidayMenu -Identity <string> -HolidayName <string> -MenuOptionsJson <json>"
    Write-Host "        Replace the entire holiday call flow menu from a JSON array."
    Write-Host "        JSON format: [{`"Key`":`"1`",`"Action`":`"RedirectToUser`",`"Target`":`"user@contoso.com`"},{`"Key`":`"2`",`"Action`":`"Disconnect`"}]"
    Write-Host "        Existing greeting and routing-options prompt are preserved."
    Write-Host "        Use -GetHolidayMenuOptions to get current options as copy-paste JSON."

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  IVR DEFAULT CALL FLOW ACTIONS" -ForegroundColor Yellow
    Write-Host "  ---------------------------------------------------------------"
    Write-Host "  These operate on the DefaultCallFlow (business hours), not holiday call flows."
    Write-Host ""
    Write-Host "    -SetIVRGreeting -Identity <string> -GreetingText <string>"
    Write-Host "        Set the TTS greeting on the IVR's default (business-hours) call flow."
    Write-Host ""
    Write-Host "    -SetIVRRoutingMessage -Identity <string> -RoutingMessage <string>"
    Write-Host "        Set the TTS menu prompt on the IVR's default call flow."
    Write-Host "        This is the message callers hear before key-press options are read."
    Write-Host ""
    Write-Host "    -SetIVRRouting -Identity <string> -RoutingAction <action> [-RedirectTarget <string>]"
    Write-Host "        Set the call routing action on the IVR's default call flow (single automatic action)."
    Write-Host "        Same RoutingAction values and -RedirectTarget semantics as -SetHolidayRouting."
    Write-Host "        For multi-option key-press menus, use -AddIVRMenuOption instead."
    Write-Host ""
    Write-Host "    -GetIVRMenuOptions -Identity <string>"
    Write-Host "        List all keypress options on the IVR's default call flow."
    Write-Host "        Prints a table (Key / Action / Target) and a JSON blob for use with -SetIVRMenu."
    Write-Host ""
    Write-Host "    -AddIVRMenuOption -Identity <string> -Key <0-9|*|#> -RoutingAction <action> [-RedirectTarget <string>]"
    Write-Host "        Add or replace a single key binding on the default call flow menu."
    Write-Host "        Key: 0-9, *, or #   RoutingAction: same values as -SetIVRRouting."
    Write-Host "        If an automatic routing option exists it is removed (incompatible with key-press)."
    Write-Host ""
    Write-Host "    -RemoveIVRMenuOption -Identity <string> -Key <0-9|*|#>"
    Write-Host "        Remove a single key binding from the default call flow menu."
    Write-Host ""
    Write-Host "    -SetIVRMenu -Identity <string> -MenuOptionsJson <json>"
    Write-Host "        Replace the entire default call flow menu from a JSON array."
    Write-Host "        JSON format: [{`"Key`":`"1`",`"Action`":`"RedirectToUser`",`"Target`":`"user@contoso.com`"},{`"Key`":`"2`",`"Action`":`"Disconnect`"}]"
    Write-Host "        Use -GetIVRMenuOptions to get current options as copy-paste JSON."

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  OPTIONS" -ForegroundColor Yellow
    Write-Host "  -------"
    Write-Host "    -TenantId <string>   Azure AD tenant ID (optional - uses default if omitted)"
    Write-Host "    -Mode <string>       Operating mode: Interactive (default) | Headless | Api"
    Write-Host "        Interactive  Human operator: shows banner, progress, colors."
    Write-Host "        Headless     Cron/script: no banner or progress; non-interactive auth."
    Write-Host "        Api          Web-app: all output is a single JSON line per call."
    Write-Host "                       Success: {`"ok`":true,`"data`":[...]} or {`"ok`":true,`"message`":`"...`"}"
    Write-Host "                       Failure: {`"ok`":false,`"error`":`"...`"} (exit code 1)"

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
    Write-Host "    Get-TeamsIVRAuthorizedUser -Identity <string>"
    Write-Host "        Returns the authorized users list for an IVR as objects (DisplayName, UPN, ObjectId)."
    Write-Host "        Called internally by -GetAuthorizedUsers and -CheckUserAccess; also usable directly."
    Write-Host ""
    Write-Host "    Test-TeamsIVRUserAccess -Identity <string> -User <string>"
    Write-Host "        Resolves the user from Teams and returns `$true if authorized, `$false if not."
    Write-Host "        -User accepts UPN, display name (must be unique), or object ID GUID."
    Write-Host "        Called internally by -CheckUserAccess; also usable directly in scripts."
    Write-Host ""
    Write-Host "    Get-TeamsIVRByUser -User <string>"
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

    # ----------------------------------------------------------------
    Write-Host ""
    Write-Host "  TYPICAL WORKFLOW: SETTING UP A HOLIDAY ON AN IVR" -ForegroundColor Yellow
    Write-Host "  ---------------------------------------------------"
    Write-Host ""
    Write-Host "  Teams stores holiday schedules inside Auto Attendant objects, so every step"
    Write-Host "  below needs the IVR's Identity GUID. Capture it first (see STEP 0 above):"
    Write-Host "    `$id = `"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`""
    Write-Host ""
    Write-Host "  ---- STEP 1 : Check whether the holiday already exists on this IVR ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -GetHoliday -HolidayName `"Christmas`""
    Write-Host "    # Lists every AA the holiday appears on."
    Write-Host "    # If your IVR is already in the output, skip to STEP 3 to update its dates."
    Write-Host "    # If not, continue to STEP 2 to create it."
    Write-Host ""
    Write-Host "  ---- STEP 2 : Create the holiday on the IVR ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -NewHoliday -Identity `$id -HolidayName `"Christmas`""
    Write-Host "    # Creates an empty holiday with a default disconnect call flow."
    Write-Host "    # No callers are affected yet -- the schedule has no dates until STEP 3."
    Write-Host ""
    Write-Host "  ---- STEP 3 : Set the date range ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayTime -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -StartDateTime `"2025-12-25 00:00`" -EndDateTime `"2025-12-26 00:00`""
    Write-Host "    # Use 00:00 for midnight start and next-day 00:00 for end-of-day."
    Write-Host "    # IMPORTANT: Both times must fall on 15-minute boundaries (0, 15, 30, or 45 min)."
    Write-Host "    # Replaces any previously set date range on this holiday."
    Write-Host ""
    Write-Host "  ---- STEP 4 : Set the greeting callers hear ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayGreeting -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -GreetingText `"Thank you for calling. Our office is closed for Christmas.`""
    Write-Host "    # This is played immediately when a call arrives during the holiday window."
    Write-Host ""
    Write-Host "  ---- STEP 5 : Set what happens after the greeting ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Option A -- Simple: just disconnect after the greeting (most common for holidays):"
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayRouting -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -RoutingAction Disconnect"
    Write-Host ""
    Write-Host "  Option B -- Redirect to a user (e.g. on-call staff):"
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayRouting -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -RoutingAction RedirectToUser -RedirectTarget `"oncall@contoso.com`""
    Write-Host ""
    Write-Host "  Option C -- Redirect to voicemail:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayRouting -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -RoutingAction RedirectToVoicemail -RedirectTarget `"sharedmailbox@contoso.com`""
    Write-Host ""
    Write-Host "  Option D -- Offer a key-press menu:"
    Write-Host ""
    Write-Host "    # 1. Set the prompt callers hear before the key options are read:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayRoutingMessage -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -RoutingMessage `"For urgent matters press 1. To leave a voicemail press 2. To disconnect press 3.`""
    Write-Host ""
    Write-Host "    # 2. Add each key binding one at a time:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -AddHolidayMenuOption -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -Key 1 -RoutingAction RedirectToUser -RedirectTarget `"oncall@contoso.com`""
    Write-Host "    .\Manage-TeamsIVR.ps1 -AddHolidayMenuOption -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -Key 2 -RoutingAction RedirectToVoicemail -RedirectTarget `"sharedmailbox@contoso.com`""
    Write-Host "    .\Manage-TeamsIVR.ps1 -AddHolidayMenuOption -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -Key 3 -RoutingAction Disconnect"
    Write-Host ""
    Write-Host "    # 3. Verify -- prints table and JSON copy-paste blob:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -GetHolidayMenuOptions -Identity `$id -HolidayName `"Christmas`""
    Write-Host ""
    Write-Host "    # 4. To update multiple bindings at once, copy the JSON from step 3, edit it, and apply:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayMenu -Identity `$id -HolidayName `"Christmas`" ``"
    Write-Host "        -MenuOptionsJson '[{`"Key`":`"1`",`"Action`":`"RedirectToUser`",`"Target`":`"oncall@contoso.com`"},{`"Key`":`"2`",`"Action`":`"Disconnect`"}]'"
    Write-Host ""
    Write-Host "    # To remove a single binding:"
    Write-Host "    .\Manage-TeamsIVR.ps1 -RemoveHolidayMenuOption -Identity `$id -HolidayName `"Christmas`" -Key 3"
    Write-Host ""
    Write-Host "  ---- STEP 6 : Verify ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    .\Manage-TeamsIVR.ps1 -GetHoliday -HolidayName `"Christmas`""
    Write-Host "    # Confirm the date range appears under your IVR."
    Write-Host ""
    Write-Host "  ---- APPLYING THE SAME HOLIDAY TO MULTIPLE IVRs ----" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Once a holiday exists on one IVR, copy it to others with -AttachHolidayToIVR."
    Write-Host "  Dates are copied automatically; greeting and routing default to disconnect."
    Write-Host ""
    Write-Host "    `$id2 = `"yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy`"   # second IVR's GUID"
    Write-Host "    .\Manage-TeamsIVR.ps1 -AttachHolidayToIVR -Identity `$id2 -HolidayName `"Christmas`""
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayGreeting -Identity `$id2 -HolidayName `"Christmas`" ``"
    Write-Host "        -GreetingText `"Our office is closed for Christmas. Please call back soon.`""
    Write-Host "    .\Manage-TeamsIVR.ps1 -SetHolidayRouting -Identity `$id2 -HolidayName `"Christmas`" ``"
    Write-Host "        -RoutingAction Disconnect"
    Write-Host ""
}

#endregion

#region Main

$Script:IsDotSourced = $MyInvocation.InvocationName -eq '.'

if ($Mode -eq 'Interactive' -and -not $Script:IsDotSourced) {
    Show-Banner
}

# ShowUsage does not require a Teams connection -- handle it before connecting
if ($PSCmdlet.ParameterSetName -eq 'ShowUsage') {
    # When dot-sourced: load functions silently; caller uses Get-Help for documentation
    # When run directly with no args: print the usage reference
    # NOTE: exit would terminate the shell session if dot-sourced, so use return
    if (-not $Script:IsDotSourced) {
        Show-Usage
    }
    return
}

# Convenience flag: Headless and Api both suppress UI and use non-interactive auth
$isHeadless = $Mode -ne 'Interactive'

# ---------------------------------------------------------------------------
# Api mode output helpers
# All output in Api mode goes through these two functions so the web-app
# always receives exactly one JSON line on stdout per invocation.
# ---------------------------------------------------------------------------
function Send-ApiSuccess {
    param(
        $Data = $null,
        [string]$Message = ''
    )
    $body = [ordered]@{ ok = $true; runId = $Script:RunId }
    if ($null -ne $Data)  { $body.data    = $Data    }
    if ($Message -ne '')  { $body.message = $Message }
    Write-Host (ConvertTo-Json $body -Depth 10 -Compress)
}

function Send-ApiError {
    # Accepts a context string plus an optional ErrorRecord.
    # When an ErrorRecord is supplied the full exception chain is appended so the
    # web-app receives the root-cause Teams error, not just the outermost wrapper.
    param(
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$Err = $null
    )
    if ($Err -and $Err.Exception) {
        $chain = [System.Collections.Generic.List[string]]::new()
        $ex = $Err.Exception
        while ($ex) {
            if ($ex.Message -and -not ($chain.Count -gt 0 -and $chain[$chain.Count - 1] -eq $ex.Message)) {
                [void]$chain.Add($ex.Message)
            }
            $ex = $ex.InnerException
        }
        if ($chain.Count -gt 0) {
            $Message = "$Message ($($chain -join ' --> '))"
        }
    }
    Write-Host (ConvertTo-Json ([ordered]@{ ok = $false; error = $Message; runId = $Script:RunId }) -Compress)
}

# ---------------------------------------------------------------------------
# Retry / cache helpers
# ---------------------------------------------------------------------------

function Invoke-TeamsRetry {
    # Wraps a Teams API scriptblock with exponential-backoff retry logic for
    # transient errors (429 rate-limit, 503/504 service unavailable, timeouts).
    # Disconnection errors are rethrown immediately with a clear message because
    # re-authentication requires a new script invocation.
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Command,
        [string]$OperationName = 'Teams API call',
        [int]$MaxRetries = 3,
        [int]$BaseDelayMs = 2000
    )
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            return (& $Command)
        } catch {
            $msg  = $_.ToString()
            $last = $attempt -ge $MaxRetries

            $isRateLimit  = $msg -match '429|Too Many Requests|rate.?limit|throttl'
            $isTransient  = $msg -match '503|504|Service Unavailable|Gateway Timeout|timed? ?out|temporarily unavailable'
            $isDisconnect = $msg -match 'not connected|sign.?in again|session expired|no active.*session|MicrosoftTeams.*connect|authentication'

            if ($isDisconnect) {
                throw "Teams session disconnected during '$OperationName'. Re-run the script to re-authenticate. [$Script:RunId] Original: $_"
            }
            if (($isRateLimit -or $isTransient) -and -not $last) {
                $delay = [int]($BaseDelayMs * [math]::Pow(2, $attempt - 1))
                Write-Verbose "[$Script:RunId] Transient error on '$OperationName' (attempt $attempt/$MaxRetries) -- retrying in ${delay}ms. Error: $_"
                Start-Sleep -Milliseconds $delay
                continue
            }
            throw   # permanent error or last attempt
        }
    }
}

function Get-CsAutoAttendantCached {
    # Cache wrapper for Get-CsAutoAttendant -Identity. Caches within one script run;
    # callers that mutate the AA should call $Script:AACache.Remove($Identity) afterward.
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Identity
    )
    if (-not $Script:AACache.ContainsKey($Identity)) {
        $Script:AACache[$Identity] = Invoke-TeamsRetry { Get-CsAutoAttendant -Identity $Identity } "Get-CsAutoAttendant '$Identity'"
    }
    return $Script:AACache[$Identity]
}

# All other actions require an active Teams connection
try {
    Connect-TeamsEnvironment -TenantId $TenantId -Headless:$isHeadless
    # When Option A or B is activated, replace the call above with:
    # Connect-TeamsEnvironment -TenantId $TenantId -Headless:$isHeadless `
    #     -ApplicationId $ApplicationId `
    #     -CertificateThumbprint $CertificateThumbprint   # Option A
    #
    # Connect-TeamsEnvironment -TenantId $TenantId -Headless:$isHeadless `
    #     -ApplicationId $ApplicationId `
    #     -ClientSecret $ClientSecret                     # Option B
    #
    # Option C requires no extra params:
    # Connect-TeamsEnvironment -TenantId $TenantId -Headless:$isHeadless
} catch {
    if ($Mode -eq 'Api') { Send-ApiError "Failed to connect to Microsoft Teams" $_ }
    exit 1
}

switch ($PSCmdlet.ParameterSetName) {

    'GetAllIVRs' {
        # Retrieve every Auto Attendant and Call Queue
        try {
            $ivrs = Get-TeamsIVRList -Headless:$isHeadless
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($ivrs)
            } else {
                $ivrs | Format-Table -AutoSize
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve IVR list: $_" $_; exit 1 }
            Write-Error "Failed to retrieve IVR list: $_"
            exit 1
        }
    }

    'GetIVR' {
        # Retrieve a specific IVR by name (partial match supported)
        try {
            $ivrs = Get-TeamsIVRList -Headless:$isHeadless
            $match = @($ivrs | Where-Object { $_.Name -like "*$Name*" })

            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data $match
            } elseif ($match.Count -eq 0) {
                Write-Warning "No IVR found matching name: '$Name'"
            } else {
                $match | Format-Table -AutoSize
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve IVR '$Name': $_" $_; exit 1 }
            Write-Error "Failed to retrieve IVR '$Name': $_"
            exit 1
        }
    }

    'GetSchedule' {
        # Retrieve raw schedule objects for an Auto Attendant
        try {
            $scheduleData = Get-TeamsIVRSchedule -Identity $Identity
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data $scheduleData
            } else {
                Write-Host "`nSchedule for: $($scheduleData.Name)" -ForegroundColor Cyan
                Write-Host "--- After-Hours Schedule (closed periods stored by Teams) ---" -ForegroundColor Yellow
                $scheduleData.AfterHours | Format-List
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve schedule for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to retrieve schedule for '$Identity': $_"
            exit 1
        }
    }

    'DecodeSchedule' {
        # Decode the schedule into human-readable open/business hours.
        # The AfterHours schedule stores closed periods -- -InvertToOpenHours flips
        # them back so the output matches the Build-TeamsIVRSchedule input format.
        try {
            $scheduleData = Get-TeamsIVRSchedule -Identity $Identity

            if ($Mode -eq 'Api') {
                # Return structured day/hours objects rather than formatted strings
                $bhLines = @()
                if ($scheduleData.AfterHours) {
                    $bhLines = @(Format-TeamsIVRSchedule -Schedule $scheduleData.AfterHours -InvertToOpenHours |
                        ForEach-Object {
                            $parts = $_ -split '\s*:\s*', 2
                            [PSCustomObject]@{
                                Day   = $parts[0].Trim()
                                Hours = if ($parts.Count -gt 1) { $parts[1].Trim() } else { '' }
                            }
                        })
                }
                Send-ApiSuccess -Data ([PSCustomObject]@{
                    Name          = $scheduleData.Name
                    Identity      = $scheduleData.Identity
                    BusinessHours = $bhLines
                })
            } else {
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
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to decode schedule for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to decode schedule for '$Identity': $_"
            exit 1
        }
    }

    'SetSchedule' {
        # Update the schedule for an Auto Attendant
        try {
            Set-TeamsIVRSchedule -Identity $Identity -Schedule $Schedule
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Schedule updated for '$Identity'."
            } else {
                Write-Host "Schedule updated successfully for Identity: $Identity" -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set schedule for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set schedule for '$Identity': $_"
            exit 1
        }
    }

    'GetAuthorizedUsers' {
        # List all authorized users for an Auto Attendant
        try {
            $authUsers = Get-TeamsIVRAuthorizedUser -Identity $Identity
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($authUsers)
            } elseif (@($authUsers).Count -eq 0) {
                Write-Host "No authorized users configured for Identity: $Identity" -ForegroundColor Yellow
            } else {
                Write-Host "`n  Authorized users for Identity: $Identity" -ForegroundColor Cyan
                $authUsers | Format-Table DisplayName, UserPrincipalName, ObjectId -AutoSize
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve authorized users for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to retrieve authorized users for '$Identity': $_"
            exit 1
        }
    }

    'CheckUserAccess' {
        # Check whether a specific user is authorized for an Auto Attendant
        try {
            $isAuthorized = Test-TeamsIVRUserAccess -Identity $Identity -User $User
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data ([PSCustomObject]@{ authorized = $isAuthorized; user = $User; identity = $Identity })
            } else {
                if ($isAuthorized) {
                    Write-Host "AUTHORIZED   - '$User' is in the authorized users list for Identity: $Identity" -ForegroundColor Green
                } else {
                    Write-Host "NOT AUTHORIZED - '$User' is not in the authorized users list for Identity: $Identity" -ForegroundColor Red
                }
                # Return the boolean so callers can capture it (e.g. $result = .\script.ps1 -CheckUserAccess ...)
                return $isAuthorized
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to check user access for '$User' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to check user access for '$User' on '$Identity': $_"
            exit 1
        }
    }

    'GetUserIVRs' {
        # List all Auto Attendants where the given user is an authorized user
        try {
            $ivrs = Get-TeamsIVRByUser -User $User -Headless:$isHeadless
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($ivrs)
            } elseif (@($ivrs).Count -eq 0) {
                Write-Host "`n  '$User' is not an authorized user for any Auto Attendant." -ForegroundColor Yellow
            } else {
                Write-Host "`n  Auto Attendants where '$User' is authorized:" -ForegroundColor Cyan
                $ivrs | Format-Table Name, Identity -AutoSize
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve IVRs for user '$User': $_" $_; exit 1 }
            Write-Error "Failed to retrieve IVRs for user '$User': $_"
            exit 1
        }
    }

    'GetHolidayList' {
        try {
            $holidays = Get-TeamsIVRHolidayList -Headless:$isHeadless
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($holidays)
            } elseif (@($holidays).Count -eq 0) {
                Write-Host "`n  No holiday schedules found across any Auto Attendant." -ForegroundColor Yellow
            } else {
                Write-Host "`n  Holiday schedules found: $(@($holidays).Count)" -ForegroundColor Cyan
                foreach ($h in $holidays) {
                    Write-Host ""
                    Write-Host "  $($h.Name)" -ForegroundColor White
                    Write-Host "    ID         : $($h.ScheduleId)"
                    Write-Host "    Linked IVRs: $($h.LinkedIVRs -join ', ')"
                    if (@($h.DateTimeRanges).Count -eq 0) {
                        Write-Host "    Dates      : (none configured)"
                    } else {
                        foreach ($r in $h.DateTimeRanges) {
                            Write-Host "    Date range : $($r.Start) -- $($r.End)"
                        }
                    }
                }
                Write-Host ""
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve holiday list: $_" $_; exit 1 }
            Write-Error "Failed to retrieve holiday list: $_"
            exit 1
        }
    }

    'GetHoliday' {
        try {
            $results = Get-TeamsIVRHoliday -HolidayName $HolidayName -Headless:$isHeadless
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($results)
            } elseif (@($results).Count -eq 0) {
                Write-Host "`n  Holiday '$HolidayName' was not found on any Auto Attendant." -ForegroundColor Yellow
            } else {
                Write-Host "`n  Holiday '$HolidayName' found on $(@($results).Count) Auto Attendant(s):" -ForegroundColor Cyan
                foreach ($r in $results) {
                    Write-Host ""
                    Write-Host "  IVR        : $($r.IVRName)"
                    Write-Host "  IVR ID     : $($r.IVRIdentity)"
                    Write-Host "  Schedule ID: $($r.ScheduleId)"
                    if (@($r.DateTimeRanges).Count -eq 0) {
                        Write-Host "  Dates      : (none configured)"
                    } else {
                        foreach ($dr in $r.DateTimeRanges) {
                            Write-Host "  Date range : $($dr.Start) -- $($dr.End)"
                        }
                    }
                }
                Write-Host ""
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve holiday '$HolidayName': $_" $_; exit 1 }
            Write-Error "Failed to retrieve holiday '$HolidayName': $_"
            exit 1
        }
    }

    'NewHoliday' {
        try {
            New-TeamsIVRHoliday -Identity $Identity -HolidayName $HolidayName
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday '$HolidayName' created on '$Identity'."
            } else {
                Write-Host "Holiday '$HolidayName' created on IVR '$Identity'." -ForegroundColor Green
                Write-Host "  Next steps: -SetHolidayTime to add dates, -SetHolidayGreeting to add a greeting."
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to create holiday '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to create holiday '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'SetHolidayTime' {
        try {
            Set-TeamsIVRHolidayTime -Identity $Identity `
                                    -HolidayName $HolidayName `
                                    -StartDateTime $StartDateTime `
                                    -EndDateTime $EndDateTime
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday '$HolidayName' date range set on '$Identity'."
            } else {
                Write-Host "Holiday '$HolidayName' date range set: $StartDateTime -- $EndDateTime" -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set time for holiday '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set time for holiday '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'AttachHolidayToIVR' {
        try {
            Add-TeamsIVRHolidayToIVR -Identity $Identity -HolidayName $HolidayName -Headless:$isHeadless
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday '$HolidayName' copied to '$Identity'."
            } else {
                Write-Host "Holiday '$HolidayName' copied to IVR '$Identity'." -ForegroundColor Green
                Write-Host "  Configure greeting/routing separately with -SetHolidayGreeting and -SetHolidayRouting."
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to attach holiday '$HolidayName' to '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to attach holiday '$HolidayName' to '$Identity': $_"
            exit 1
        }
    }

    'SetHolidayGreeting' {
        try {
            Set-TeamsIVRHolidayGreeting -Identity $Identity `
                                        -HolidayName $HolidayName `
                                        -GreetingText $GreetingText
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday greeting updated for '$HolidayName' on '$Identity'."
            } else {
                Write-Host "Holiday greeting updated for '$HolidayName' on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set holiday greeting for '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set holiday greeting for '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'SetHolidayRoutingMessage' {
        try {
            Set-TeamsIVRHolidayRoutingMessage -Identity $Identity `
                                              -HolidayName $HolidayName `
                                              -RoutingMessage $RoutingMessage
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday routing-options message updated for '$HolidayName' on '$Identity'."
            } else {
                Write-Host "Holiday routing-options message updated for '$HolidayName' on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set holiday routing message for '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set holiday routing message for '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'SetHolidayRouting' {
        try {
            Set-TeamsIVRHolidayRouting -Identity $Identity `
                                       -HolidayName $HolidayName `
                                       -RoutingAction $RoutingAction `
                                       -RedirectTarget $RedirectTarget
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday routing action '$RoutingAction' set for '$HolidayName' on '$Identity'."
            } else {
                Write-Host "Holiday routing action '$RoutingAction' set for '$HolidayName' on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set holiday routing for '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set holiday routing for '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'SetIVRGreeting' {
        try {
            Set-TeamsIVRGreeting -Identity $Identity -GreetingText $GreetingText
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "IVR greeting updated for '$Identity'."
            } else {
                Write-Host "IVR greeting updated for '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set IVR greeting for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set IVR greeting for '$Identity': $_"
            exit 1
        }
    }

    'SetIVRRoutingMessage' {
        try {
            Set-TeamsIVRRoutingMessage -Identity $Identity -RoutingMessage $RoutingMessage
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "IVR routing-options message updated for '$Identity'."
            } else {
                Write-Host "IVR routing-options message updated for '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set IVR routing message for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set IVR routing message for '$Identity': $_"
            exit 1
        }
    }

    'SetIVRRouting' {
        try {
            Set-TeamsIVRRouting -Identity $Identity `
                                -RoutingAction $RoutingAction `
                                -RedirectTarget $RedirectTarget
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "IVR routing action '$RoutingAction' set for '$Identity'."
            } else {
                Write-Host "IVR routing action '$RoutingAction' set for '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set IVR routing for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set IVR routing for '$Identity': $_"
            exit 1
        }
    }

    'GetHolidayMenuOptions' {
        try {
            $specs = Get-TeamsIVRHolidayMenuOptions -Identity $Identity -HolidayName $HolidayName
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($specs)
            } elseif (@($specs).Count -eq 0) {
                Write-Host "`n  No menu options found for holiday '$HolidayName' on '$Identity'." -ForegroundColor Yellow
            } else {
                Write-Host "`n  Menu options for holiday '$HolidayName' on '$Identity':" -ForegroundColor Cyan
                $specs | Format-Table -AutoSize @{L='Key';E={$_.Key}}, @{L='Action';E={$_.Action}}, @{L='Target';E={if ($_.Target) {$_.Target} else {''}}}
                Write-Host "  JSON (for -SetHolidayMenu -MenuOptionsJson):"
                Write-Host "  $(Format-MenuOptionsAsJson -Specs $specs)" -ForegroundColor White
                Write-Host ""
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve menu options for '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to retrieve menu options for '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'AddHolidayMenuOption' {
        try {
            Add-TeamsIVRHolidayMenuOption -Identity $Identity `
                                          -HolidayName $HolidayName `
                                          -Key $Key `
                                          -RoutingAction $RoutingAction `
                                          -RedirectTarget $RedirectTarget
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Key '$Key' ($RoutingAction) added to '$HolidayName' menu on '$Identity'."
            } else {
                Write-Host "Key '$Key' ($RoutingAction) added to '$HolidayName' menu on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to add menu option to '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to add menu option to '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'RemoveHolidayMenuOption' {
        try {
            Remove-TeamsIVRHolidayMenuOption -Identity $Identity -HolidayName $HolidayName -Key $Key
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Key '$Key' removed from '$HolidayName' menu on '$Identity'."
            } else {
                Write-Host "Key '$Key' removed from '$HolidayName' menu on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to remove menu option from '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to remove menu option from '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'SetHolidayMenu' {
        try {
            Set-TeamsIVRHolidayMenu -Identity $Identity -HolidayName $HolidayName -MenuOptionsJson $MenuOptionsJson
            $count = @($MenuOptionsJson | ConvertFrom-Json).Count
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Holiday '$HolidayName' menu replaced with $count option(s) on '$Identity'."
            } else {
                Write-Host "Holiday '$HolidayName' menu replaced with $count option(s) on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set holiday menu for '$HolidayName' on '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set holiday menu for '$HolidayName' on '$Identity': $_"
            exit 1
        }
    }

    'GetIVRMenuOptions' {
        try {
            $specs = Get-TeamsIVRMenuOptions -Identity $Identity
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Data @($specs)
            } elseif (@($specs).Count -eq 0) {
                Write-Host "`n  No menu options found on default call flow for '$Identity'." -ForegroundColor Yellow
            } else {
                Write-Host "`n  Menu options on default call flow for '$Identity':" -ForegroundColor Cyan
                $specs | Format-Table -AutoSize @{L='Key';E={$_.Key}}, @{L='Action';E={$_.Action}}, @{L='Target';E={if ($_.Target) {$_.Target} else {''}}}
                Write-Host "  JSON (for -SetIVRMenu -MenuOptionsJson):"
                Write-Host "  $(Format-MenuOptionsAsJson -Specs $specs)" -ForegroundColor White
                Write-Host ""
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to retrieve menu options for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to retrieve menu options for '$Identity': $_"
            exit 1
        }
    }

    'AddIVRMenuOption' {
        try {
            Add-TeamsIVRMenuOption -Identity $Identity `
                                   -Key $Key `
                                   -RoutingAction $RoutingAction `
                                   -RedirectTarget $RedirectTarget
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Key '$Key' ($RoutingAction) added to default call flow menu on '$Identity'."
            } else {
                Write-Host "Key '$Key' ($RoutingAction) added to default call flow menu on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to add menu option to '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to add menu option to '$Identity': $_"
            exit 1
        }
    }

    'RemoveIVRMenuOption' {
        try {
            Remove-TeamsIVRMenuOption -Identity $Identity -Key $Key
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Key '$Key' removed from default call flow menu on '$Identity'."
            } else {
                Write-Host "Key '$Key' removed from default call flow menu on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to remove menu option from '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to remove menu option from '$Identity': $_"
            exit 1
        }
    }

    'SetIVRMenu' {
        try {
            Set-TeamsIVRMenu -Identity $Identity -MenuOptionsJson $MenuOptionsJson
            $count = @($MenuOptionsJson | ConvertFrom-Json).Count
            if ($Mode -eq 'Api') {
                Send-ApiSuccess -Message "Default call flow menu replaced with $count option(s) on '$Identity'."
            } else {
                Write-Host "Default call flow menu replaced with $count option(s) on '$Identity'." -ForegroundColor Green
            }
        } catch {
            if ($Mode -eq 'Api') { Send-ApiError "Failed to set IVR menu for '$Identity': $_" $_; exit 1 }
            Write-Error "Failed to set IVR menu for '$Identity': $_"
            exit 1
        }
    }
}

#endregion
