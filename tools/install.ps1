<#
.SYNOPSIS
    Install RecallTape into OneNote.

.DESCRIPTION
    Registers the add-in in place, from wherever you extracted this folder. Nothing is copied
    elsewhere, so KEEP THIS FOLDER where it is -- deleting or moving it breaks the add-in.

    Needs Administrator once, to register a COM class machine-wide. Everything else is per-user.

    To remove it completely, run uninstall.ps1 as Administrator.

.NOTES
    Pre-alpha. Not a signed installer; Windows will be suspicious and it is right to be.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$Clsid = '{AA568A3C-2A53-479B-B188-2367D2E27CE4}'
$here  = $PSScriptRoot
$dll   = Join-Path $here 'RecallTape.OneNote.dll'
$exe   = Join-Path $here 'RecallTape.ProtocolHandler.exe'

function Fail($message) { Write-Host "ERROR: $message" -ForegroundColor Red; exit 1 }

# --- preflight: fail with an explanation, never half-install -------------------------------
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Fail "Administrator required. Right-click install.ps1 and choose 'Run with PowerShell as Administrator'."
}
if (-not (Test-Path $dll)) { Fail "RecallTape.OneNote.dll is not next to this script. Extract the whole zip and run it from there." }
if (-not (Test-Path $exe)) { Fail "RecallTape.ProtocolHandler.exe is not next to this script. Extract the whole zip and run it from there." }

$regasm = Join-Path $env:WinDir 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'
if (-not (Test-Path $regasm)) { Fail ".NET Framework 4.x not found. RecallTape needs .NET Framework 4.8." }

if (-not (Test-Path 'C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE')) {
    Write-Host "WARNING: OneNote desktop was not found where expected." -ForegroundColor Yellow
    Write-Host "         RecallTape needs OneNote for Windows desktop, not the Store app. Continuing anyway." -ForegroundColor Yellow
}

Write-Host "Installing RecallTape from $here"
Write-Host ""

# --- 1. the COM class ----------------------------------------------------------------------
& $regasm $dll /codebase | Out-Null
Write-Host "  [1/4] COM class registered"

# --- 2. the surrogate ----------------------------------------------------------------------
# OneNote will not activate a managed add-in in-process: it fails the load and silently demotes
# LoadBehavior 3 -> 2. An empty DllSurrogate moves us into dllhost.exe, which also means a crash
# in RecallTape costs a dllhost and leaves OneNote - and your notes - untouched.
Set-ItemProperty "HKLM:\SOFTWARE\Classes\CLSID\$Clsid" -Name 'AppID' -Value $Clsid
New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$Clsid" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\AppID\$Clsid" -Name '(default)'    -Value 'RecallTape'
Set-ItemProperty "HKLM:\SOFTWARE\Classes\AppID\$Clsid" -Name 'DllSurrogate' -Value ''

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
Set-ItemProperty "HKLM:\SOFTWARE\Classes\AppID\$Clsid" -Name 'LaunchPermission' -Value $bytes -Type Binary
Write-Host "  [2/4] COM surrogate configured (with ARM64 launch permission)"

# --- 3. OneNote add-in registration --------------------------------------------------------
# HKCU ONLY, deliberately. OneMore also registers under HKLM as a fallback for Intune/SYSTEM
# deployments, and we should not copy that: an add-in registered machine-wide cannot be disabled from
# OneNote's COM Add-ins dialog by a non-elevated user, because unticking the box writes LoadBehavior
# into that hive. A per-user install should be a user's to switch off without asking for admin.
$addin = 'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn'
New-Item -Path $addin -Force | Out-Null
Set-ItemProperty $addin -Name 'FriendlyName' -Value 'RecallTape'
Set-ItemProperty $addin -Name 'Description'  -Value 'Active-recall masking for OneNote'
Set-ItemProperty $addin -Name 'LoadBehavior' -Type DWord -Value 3
Write-Host "  [3/4] Registered with OneNote"

# --- 4. the recalltape:// protocol ---------------------------------------------------------
# Clicking a tape strip makes OneNote hand its hyperlink to the shell. The handler cannot touch
# the page itself, so it couriers the command to the add-in over a named pipe.
$proto = 'HKCU:\SOFTWARE\Classes\recalltape'
New-Item -Path $proto -Force | Out-Null
Set-ItemProperty $proto -Name '(default)'    -Value 'URL:RecallTape Protocol'
Set-ItemProperty $proto -Name 'URL Protocol' -Value ''
New-Item -Path "$proto\shell\open\command" -Force | Out-Null
Set-ItemProperty "$proto\shell\open\command" -Name '(default)' -Value ('"' + $exe + '" "%1"')

# Without this, Office prompts every single time you click a piece of tape.
New-Item -Path 'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Security\Trusted Protocols\All Applications\recalltape:' -Force | Out-Null
Write-Host "  [4/4] Click-to-peek enabled"

Write-Host ""
Write-Host "Done. Restart OneNote and look for the RecallTape tab." -ForegroundColor Green
Write-Host ""
Write-Host "Keep this folder where it is -- RecallTape runs from here."
Write-Host "To remove it, run uninstall.ps1 as Administrator."
