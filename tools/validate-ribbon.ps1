<#
.SYNOPSIS
    Validate the add-in's ribbon XML against the CustomUI schema.

.DESCRIPTION
    A single undeclared attribute makes Office reject the WHOLE customUI. Not the control, not the
    group - all of it. The add-in still loads, every callback still exists, the log looks perfectly
    healthy, and the RecallTape tab is simply absent.

    That happened on 2026-08-13 with getItemImageMso, which reads like a real attribute, is what you
    would reach for to put an imageMso on a gallery item, and does not exist. The schema has no such
    thing: a dynamic gallery can only return IPictureDisp images. The tab vanished and nothing
    anywhere said why.

    So the XML gets checked before it ships. This extracts the literal out of AddIn.cs and validates
    it, which takes a second and turns an invisible failure into a build error.

.NOTES
    The schema is customui14.xsd, matching the 2009/07 namespace we declare.
#>
[CmdletBinding()]
param(
    [string] $Source,
    [string] $Schema = 'C:\DEV\ewc3labs\OneMore\Reference\customui14.xsd'
)

$ErrorActionPreference = 'Stop'

# Resolved here, not in the param block: $PSScriptRoot is not populated there under PowerShell 5.1,
# which is what a stock Windows box runs.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Source) { $Source = Join-Path $root '..\src\RecallTape.OneNote\AddIn.cs' }

if (-not (Test-Path $Schema)) {
    Write-Host "SKIPPED: schema not found at $Schema" -ForegroundColor Yellow
    Write-Host "         (customui14.xsd ships with the Office SDK / OneMore's Reference folder)"
    exit 0
}

$code = Get-Content $Source -Raw
$start = $code.IndexOf('return @"<customUI')
if ($start -lt 0) { Write-Host "ERROR: no customUI literal found in $Source" -ForegroundColor Red; exit 1 }
$start += 'return @"'.Length
$end = $code.IndexOf('</customUI>', $start) + '</customUI>'.Length

# Verbatim string: "" is one literal quote. Our generated items go in at runtime; a placeholder
# would not validate, so substitute one representative item to prove the shape is legal.
$xml = $code.Substring($start, $end - $start).Replace('""', '"')
$xml = $xml.Replace('{ICONS}', "<item id='ico0' imageMso='Bold' label='Bold' screentip='Bold'/>")

$tmp = Join-Path $env:TEMP 'recalltape-ribbon-check.xml'
Set-Content -Path $tmp -Value $xml -Encoding UTF8

$problems = New-Object System.Collections.ArrayList
$settings = New-Object System.Xml.XmlReaderSettings
$settings.ValidationType = [System.Xml.ValidationType]::Schema
[void]$settings.Schemas.Add('http://schemas.microsoft.com/office/2009/07/customui', $Schema)
$settings.add_ValidationEventHandler([System.Xml.Schema.ValidationEventHandler]{
    param($sender, $e) [void]$problems.Add($e.Message)
})

$reader = [System.Xml.XmlReader]::Create($tmp, $settings)
try { while ($reader.Read()) { } }
catch { [void]$problems.Add("not well-formed: $($_.Exception.Message)") }
finally { $reader.Close() }

if ($problems.Count -gt 0) {
    Write-Host ""
    Write-Host "RIBBON XML IS INVALID - Office would show NO RecallTape tab at all:" -ForegroundColor Red
    $problems | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    Write-Host ""
    exit 1
}

Write-Host "  ribbon XML validates against customui14.xsd" -ForegroundColor Green
exit 0
