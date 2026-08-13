<#
.SYNOPSIS
    Close OneNote properly, wait for it to actually exit, then start exactly one instance.

.DESCRIPTION
    The dev loop needs OneNote restarted constantly, because the COM surrogate holds the built DLL
    open. Doing that by hand with Stop-Process -Force is how you end up with TWO ONENOTE.EXE
    processes contending for one cache, and the second one sits forever on
    "We're sorry. OneNote is cleaning up from the last time it was open."

    That happened on 2026-08-13. The rules that came out of it:

      1. Ask first. CloseMainWindow lets OneNote flush its cache and release the surrogate.
      2. WAIT for the process to actually leave. OneNote lingers well after its window closes.
      3. Force only as a last resort, and only after a real timeout.
      4. NEVER start a new instance until the old one is gone. This is the one that bit.

    Also: a modal dialog we own (Remove All's confirmation) runs on the thread OneNote called us on,
    so OneNote is blocked while it is up. Killing OneNote at that moment is an unclean shutdown by
    construction. Close gracefully and any open dialog goes with it.
#>
[CmdletBinding()]
param(
    [switch] $NoStart,
    [int]    $TimeoutSeconds = 45
)

$exe = 'C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE'

function Wait-Gone([int] $seconds) {
    $deadline = (Get-Date).AddSeconds($seconds)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-Process ONENOTE -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return -not (Get-Process ONENOTE -ErrorAction SilentlyContinue)
}

$running = @(Get-Process ONENOTE -ErrorAction SilentlyContinue)
if ($running.Count -gt 1) {
    Write-Host "  NOTE: $($running.Count) OneNote processes are running. That is the cache-contention state." -ForegroundColor Yellow
}

if ($running) {
    Write-Host "  asking $($running.Count) OneNote process(es) to close..."
    $running | ForEach-Object { $_.CloseMainWindow() | Out-Null }

    if (Wait-Gone $TimeoutSeconds) {
        Write-Host "  closed cleanly"
    } else {
        $left = @(Get-Process ONENOTE -ErrorAction SilentlyContinue)
        Write-Host "  still here after ${TimeoutSeconds}s; forcing $($left.Count)" -ForegroundColor Yellow
        $left | ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }
        if (-not (Wait-Gone 15)) { throw "OneNote will not exit. Do not start another one." }
        Write-Host "  forced"
    }
}

# Belt and braces: a surrogate holding our DLL keeps the build locked even with OneNote gone.
Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -match 'AA568A3C' } |
    ForEach-Object {
        Write-Host "  releasing our COM surrogate (pid $($_.ProcessId))"
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

if (-not $NoStart) {
    if (Get-Process ONENOTE -ErrorAction SilentlyContinue) { throw "Refusing to start a second instance." }
    # via explorer so it launches de-elevated even from an elevated shell - an elevated OneNote
    # cannot talk to a non-elevated shell, and vice versa.
    Start-Process explorer.exe -ArgumentList "`"$exe`""
    Write-Host "  started one instance"
}
