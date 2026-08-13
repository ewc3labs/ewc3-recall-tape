<#
    Register the RecallTape spike add-in with OneNote (per-user, dev loop).

    WHY this exists as a script: registering a COM add-in is several facts that must agree (CLSID,
    ProgID, codebase path, LoadBehavior, surrogate), plus a URL protocol and its trust entry. Doing it
    by hand is how you end up debugging a "why won't it load" that is really a typo.

    Run tools/unregister.ps1 to undo all of it.
    Needs an elevated shell only because RegAsm writes to HKLM\SOFTWARE\Classes.
#>
param([string] $Configuration = 'Release')
$ErrorActionPreference = 'Stop'

$Clsid = '{AA568A3C-2A53-479B-B188-2367D2E27CE4}'
$repo  = Split-Path -Parent $PSScriptRoot
$dll   = Join-Path $repo ('src\RecallTape.OneNote\bin\' + $Configuration + '\net48\RecallTape.OneNote.dll')
$exe   = Join-Path $repo ('src\RecallTape.ProtocolHandler\bin\' + $Configuration + '\net48\RecallTape.ProtocolHandler.exe')

if (-not (Test-Path $dll)) { throw "Build first: $dll not found" }

$regasm = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
Write-Host "RegAsm /codebase -> $dll"
& $regasm $dll /codebase | Write-Host

# --- COM surrogate -------------------------------------------------------------------------
# WHY: OneNote will not activate a managed add-in in-process. Registered the plain RegAsm way
# (InprocServer32 -> mscoree.dll) OneNote fails the load and demotes LoadBehavior 3 -> 2 before
# our OnConnection ever runs. An empty DllSurrogate value moves the add-in into the generic
# dllhost.exe surrogate, which is what OneMore does and why it works. Measured, not copied:
# in-proc was tried first on this machine and failed.
$clsidKey = "HKLM:\SOFTWARE\Classes\CLSID\$Clsid"
$appidKey = "HKLM:\SOFTWARE\Classes\AppID\$Clsid"
Set-ItemProperty $clsidKey -Name 'AppID' -Value $Clsid
New-Item -Path $appidKey -Force | Out-Null
Set-ItemProperty $appidKey -Name '(default)'    -Value 'RecallTape'
Set-ItemProperty $appidKey -Name 'DllSurrogate' -Value ''

# LaunchPermission: self-relative security descriptor granting local launch and local activate to
# Authenticated Users, SYSTEM and Administrators.
#   SDDL: O:BAG:BAD:(A;;0x0b;;;AU)(A;;0x0b;;;SY)(A;;0x0b;;;BA)
#
# WHY THIS IS NOT OPTIONAL: on ARM64 Windows the machine-wide DCOM default launch security is more
# restrictive than on x64 and does not grant local-launch to BUILTIN\Users. Without this, COM refuses
# to start dllhost.exe and the add-in NEVER LOADS -- silently.
#
# This bites even a user who IS a local administrator. With UAC on, an admin's normal processes run
# with a FILTERED token in which BUILTIN\Administrators is marked deny-only: usable to deny access,
# never to grant it. OneNote runs non-elevated, so an ACE granting launch to BA does nothing for it,
# and the check falls through to what the filtered token can actually grant with. That makes AU
# (Authenticated Users) the load-bearing entry here rather than belt-and-braces.
#
# Snapdragon Surfaces are the target hardware, so this is the common case, not the edge case.
# Harmless on x64: it just makes the implicit grant explicit. Credit: OneMore's OneMoreSetup/Registry.wxs.
$sd = '010004801400000024000000000000003400000001020000000000052000000020020000010200000000000520000000200200000200480003000000000014000b00000001010000000000050b000000000014000b000000010100000000000512000000000018000b00000001020000000000052000000020020000'
$bytes = [byte[]]::new($sd.Length / 2)
for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [Convert]::ToByte($sd.Substring($i * 2, 2), 16) }
Set-ItemProperty $appidKey -Name 'LaunchPermission' -Value $bytes -Type Binary
Write-Host "Surrogate registered: AppID $Clsid DllSurrogate=(empty) + LaunchPermission"

# --- OneNote add-in discovery --------------------------------------------------------------
# LoadBehavior 3 = load at startup and keep loading. OneNote silently rewrites this to 2 when a
# load fails, so re-arm it before every test or the next run is meaningless.
$key = 'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn'
New-Item -Path $key -Force | Out-Null
Set-ItemProperty $key -Name 'FriendlyName' -Value 'RecallTape'
Set-ItemProperty $key -Name 'Description'  -Value 'Active-recall masking for OneNote (spike build)'
Set-ItemProperty $key -Name 'LoadBehavior' -Type DWord -Value 3
Write-Host "Registered $key (LoadBehavior=3)"

# --- recalltape:// URL protocol ------------------------------------------------------------
# Clicking a tape strip makes OneNote hand its Image `hyperlink` to the shell, which launches this
# handler. It cannot touch the page itself -- external processes get a dead OneNote Application on
# this build -- so it couriers the command to the add-in over a named pipe.
# HKCU rather than HKCR: per-user, no elevation needed for this part.
if (Test-Path $exe) {
    $proto = 'HKCU:\SOFTWARE\Classes\recalltape'
    New-Item -Path $proto -Force | Out-Null
    Set-ItemProperty $proto -Name '(default)'   -Value 'URL:RecallTape Protocol'
    Set-ItemProperty $proto -Name 'URL Protocol' -Value ''
    New-Item -Path "$proto\shell\open\command" -Force | Out-Null
    Set-ItemProperty "$proto\shell\open\command" -Name '(default)' -Value ('"' + $exe + '" "%1"')
    Write-Host "Protocol registered: recalltape:// -> $exe"

    # Office warns before launching an unfamiliar protocol. This entry is what suppresses that
    # prompt -- OneMore ships the same thing for onemore:. Without it, every peek costs a dialog.
    $trusted = 'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications\recalltape:'
    New-Item -Path $trusted -Force | Out-Null
    Write-Host "Protocol trusted for Office (no prompt on click)"
} else {
    Write-Host "NOTE: protocol handler not built ($exe) - tape will be clickable but nothing will answer"
}

Write-Host ""
Write-Host "Restart OneNote to load it."
