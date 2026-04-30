[CmdletBinding()]
param(
    [string]$ServiceIdentifier = "5885",
    [string]$OutputRoot = "C:\Xerox\FFCoreForChatGPT\files\diagnostics",
    [int]$EventLogDays = 7
)

$ErrorActionPreference = "Continue"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$caseRoot = Join-Path $OutputRoot "FFCore_Aborted_Service_${ServiceIdentifier}_$timestamp"
New-Item -ItemType Directory -Force -Path $caseRoot | Out-Null

$ffInstall = "C:\Program Files\Xerox\FreeFlow Core"
$ffRuntime = "C:\Xerox\FreeFlow\Core"
$logRoots = @(
    (Join-Path $ffInstall "Logs"),
    (Join-Path $ffRuntime "Logs")
) | Where-Object { Test-Path -LiteralPath $_ }

function Write-Text {
    param([string]$Name, [scriptblock]$Command)
    $path = Join-Path $caseRoot $Name
    try {
        & $Command | Out-File -FilePath $path -Encoding UTF8 -Width 240
    }
    catch {
        "ERROR: $($_.Exception.Message)" | Out-File -FilePath $path -Encoding UTF8
    }
}

Write-Text "00_ReadMe.txt" {
    @"
FFCore aborted service diagnostic pickup.
Service Identifier: $ServiceIdentifier
Collected: $(Get-Date)
Computer: $env:COMPUTERNAME

This script is read-only. It does not restart services or modify FreeFlow Core.
"@
}

Write-Text "01_Xerox_Services.txt" {
    Get-Service | Where-Object {
        $_.Name -like "*Xerox*" -or $_.DisplayName -like "*Xerox*" -or
        $_.Name -like "*FreeFlow*" -or $_.DisplayName -like "*FreeFlow*"
    } | Sort-Object DisplayName | Select-Object Status, Name, DisplayName, StartType
}

Write-Text "02_Recent_Application_Events.txt" {
    $start = (Get-Date).AddDays(-1 * [Math]::Abs($EventLogDays))
    Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = $start; Level = 1,2,3 } -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Message -like "*$ServiceIdentifier*" -or $_.Message -like "*FreeFlow*" -or
            $_.Message -like "*Xerox*" -or $_.Message -like "*Barcode*" -or
            $_.Message -like "*aborted*"
        } | Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
}

Write-Text "03_Recent_System_Events.txt" {
    $start = (Get-Date).AddDays(-1 * [Math]::Abs($EventLogDays))
    Get-WinEvent -FilterHashtable @{ LogName = "System"; StartTime = $start; Level = 1,2,3 } -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Message -like "*$ServiceIdentifier*" -or $_.Message -like "*FreeFlow*" -or
            $_.Message -like "*Xerox*" -or $_.ProviderName -like "*Service Control Manager*"
        } | Select-Object TimeCreated, ProviderName, Id, LevelDisplayName, Message
}

Write-Text "04_Log_Search_ServiceIdentifier_Barcode_Aborted.txt" {
    foreach ($root in $logRoots) {
        "===== Searching $root ====="
        Get-ChildItem -LiteralPath $root -Recurse -File -Include "*.log","*.txt","*.xml" -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -lt 50MB } |
            Select-String -Pattern $ServiceIdentifier,"Barcode","aborted","Échec","Echec","failed","exception","error" -CaseSensitive:$false -Context 3,3 -ErrorAction SilentlyContinue |
            ForEach-Object {
                "File: $($_.Path)"
                "Line: $($_.LineNumber)"
                $_.Context.PreContext
                $_.Line
                $_.Context.PostContext
                ""
            }
    }
}

foreach ($root in $logRoots) {
    $dest = Join-Path $caseRoot ("Logs_" + (($root -replace "[:\\ ]","_").Trim("_")))
    New-Item -ItemType Directory -Force -Path $dest | Out-Null
    Get-ChildItem -LiteralPath $root -Recurse -File -Include "*.log","*.txt","*.xml" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 150 |
        ForEach-Object {
            $safeName = ($_.FullName.Substring($root.Length).TrimStart("\") -replace "[\\/:*?""<>|]", "_")
            Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $dest $safeName) -Force
        }
}

$zip = Join-Path (Split-Path -Parent $caseRoot) ("FFCore_Aborted_Service_${ServiceIdentifier}_$timestamp.zip")
Compress-Archive -LiteralPath $caseRoot -DestinationPath $zip -Force

Write-Host "Collected folder: $caseRoot"
Write-Host "Created zip: $zip"
