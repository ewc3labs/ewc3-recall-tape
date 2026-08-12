<#
    Register the RecallTape spike add-in with OneNote (per-user, dev loop).

    WHY this exists as a script: registering a COM add-in is four separate facts that must agree
    (CLSID, ProgID, codebase path, LoadBehavior). Doing it by hand is how you end up debugging a
    "why won't it load" that is really a typo. Run tools/unregister.ps1 to undo all of it.

    Needs an elevated shell only because RegAsm writes to HKLM\SOFTWARE\Classes.
#>
param([string] $Configuration = 'Release')
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
$dll  = Join-Path $repo ('src\RecallTape.OneNote\bin\' + $Configuration + '\net48\RecallTape.OneNote.dll')
if (-not (Test-Path $dll)) { throw "Build first: $dll not found" }

$Clsid = '{AA568A3C-2A53-479B-B188-2367D2E27CE4}'
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
Set-ItemProperty $appidKey -Name '(default)'   -Value 'RecallTape'
Set-ItemProperty $appidKey -Name 'DllSurrogate' -Value ''
Write-Host "Surrogate registered: AppID $Clsid DllSurrogate=(empty)"

# OneNote discovers add-ins here. LoadBehavior 3 = load at startup and keep loading.
# HKCU rather than HKLM: this is a dev registration for one user, not an install.
$key = 'HKCU:\SOFTWARE\Microsoft\Office\OneNote\AddIns\RecallTape.AddIn'
New-Item -Path $key -Force | Out-Null
Set-ItemProperty $key -Name 'FriendlyName' -Value 'RecallTape'
Set-ItemProperty $key -Name 'Description'  -Value 'Active-recall masking for OneNote (spike build)'
Set-ItemProperty $key -Name 'LoadBehavior' -Type DWord -Value 3
Write-Host "Registered $key (LoadBehavior=3). Restart OneNote to load it."
