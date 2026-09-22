<#
.SYNOPSIS
    Register the classic OneNote type library for 64-bit callers (per user).

.DESCRIPTION
    Classic OneNote 2016 registers its type library only under the `win32` key:

        HKCR\TypeLib\{0EA692EE-BB50-4E3C-AEF0-356D91732725}\1.1\0\win32

    A 64-bit process (such as 64-bit Python driving m365-mcp) therefore gets
    `TYPE_E_LIBNOTREGISTERED` (0x8002801D, "Library not registered") the moment
    it calls GetHierarchy / FindPages / UpdatePageContent, even though OneNote
    itself is 64-bit and starts fine.

    This script adds the matching `win64` entry under HKCU\Software\Classes,
    which Windows merges into HKCR. It needs NO administrator rights and it
    touches nothing outside the current user's own class registrations.

    Run -WhatIf first to see exactly what would be written, and -Revert to
    remove the key again.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\fix_onenote_typelib.ps1 -WhatIf
    powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\fix_onenote_typelib.ps1
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$Revert
)

$ErrorActionPreference = 'Stop'

$libGuid = '{0EA692EE-BB50-4E3C-AEF0-356D91732725}'
$userKey = "HKCU:\Software\Classes\TypeLib\$libGuid\1.1\0"
$win64Key = "$userKey\win64"

if ($Revert) {
    if (Test-Path $win64Key) {
        if ($PSCmdlet.ShouldProcess($win64Key, 'Remove registry key')) {
            Remove-Item $win64Key -Recurse -Force
            Write-Host "Removed $win64Key"
        }
    } else {
        Write-Host "Nothing to revert: $win64Key does not exist."
    }
    return
}

# 1. Where does the machine say the 32-bit typelib lives?
$machineWin32 = "HKLM:\SOFTWARE\Classes\TypeLib\$libGuid\1.1\0\win32"
if (-not (Test-Path $machineWin32)) {
    throw "OneNote type library is not registered at all ($machineWin32 missing). Is classic OneNote installed?"
}
$tlbPath = (Get-ItemProperty $machineWin32).'(default)'
Write-Host "Machine win32 entry: $tlbPath"

# 2. Sanity-check the file it points at (path may carry a resource index like \3).
$filePart = $tlbPath -replace '\\\d+$', ''
if (-not (Test-Path -LiteralPath $filePart)) {
    throw "The registered type library file does not exist: $filePart"
}

# 3. Is the app actually 64-bit? If OneNote is 32-bit, this fix is wrong -
#    run the MCP server on 32-bit Python instead.
$bytes = [System.IO.File]::ReadAllBytes($filePart)
$peOffset = [BitConverter]::ToInt32($bytes, 0x3C)
$machine = [BitConverter]::ToUInt16($bytes, $peOffset + 4)
$is64 = ($machine -eq 0x8664)
Write-Host ("OneNote executable is {0}-bit" -f $(if ($is64) { '64' } else { '32' }))
if (-not $is64) {
    throw "OneNote is 32-bit; a 64-bit typelib registration would be a lie. Run m365-mcp on 32-bit Python instead."
}

if (Test-Path $win64Key) {
    $current = (Get-ItemProperty $win64Key).'(default)'
    Write-Host "Already present: $win64Key -> $current"
    return
}

if ($PSCmdlet.ShouldProcess($win64Key, "Create with value '$tlbPath'")) {
    New-Item -Path $win64Key -Force | Out-Null
    Set-ItemProperty -Path $win64Key -Name '(default)' -Value $tlbPath
    Write-Host "Created $win64Key -> $tlbPath"
    Write-Host "Restart the MCP server (and any Python that already loaded the cache), then retry onenote_hierarchy."
}
