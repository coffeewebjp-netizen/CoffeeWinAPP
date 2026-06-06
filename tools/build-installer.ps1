param(
    [ValidateSet("Auto", "Inno", "IExpress")]
    [string]$InstallerType = "Auto",
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "1.0.0",
    [switch]$SelfContained,
    [string]$PublishPath = "publish\CoffeeAutoButton",
    [string]$OutputDir = "installer\out"
)

$ErrorActionPreference = "Stop"

function Resolve-RepoPath([string]$Path) {
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path $script:Root.Path $Path
}

function Assert-UnderRoot([string]$Path) {
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $rootPath = [System.IO.Path]::GetFullPath($script:Root.Path)
    if (!$fullPath.StartsWith($rootPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Path is outside repository root: $fullPath"
    }
}

function New-InstallerPayload([string]$PublishFullPath, [string]$PayloadDir) {
    Assert-UnderRoot $PayloadDir

    if (Test-Path $PayloadDir) {
        Remove-Item -LiteralPath $PayloadDir -Recurse -Force
    }

    New-Item -ItemType Directory -Path $PayloadDir | Out-Null
    Copy-Item -Path (Join-Path $PublishFullPath "*") -Destination $PayloadDir -Recurse -Force

    Set-Content -Path (Join-Path $PayloadDir "install.cmd") -Encoding ASCII -Value @'
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
exit /b %ERRORLEVEL%
'@

    Set-Content -Path (Join-Path $PayloadDir "install.ps1") -Encoding UTF8 -Value @'
$ErrorActionPreference = "Stop"

$appName = "Coffee AutoButton"
$appFolder = "CoffeeAutoButton"
$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$installDir = Join-Path $env:LOCALAPPDATA $appFolder
$programsRoot = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$programsDir = Join-Path $programsRoot $appName

New-Item -ItemType Directory -Path $installDir -Force | Out-Null
New-Item -ItemType Directory -Path $programsDir -Force | Out-Null

Get-ChildItem -LiteralPath $sourceDir -File |
    Where-Object { $_.Name -notin @("install.cmd", "install.ps1") } |
    Copy-Item -Destination $installDir -Force

function New-Shortcut($Path, $Target, $Arguments = "") {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($Path)
    $shortcut.TargetPath = $Target
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = Split-Path -Parent $Target
    $shortcut.Save()
}

New-Shortcut `
    -Path (Join-Path $programsDir "$appName.lnk") `
    -Target (Join-Path $installDir "CoffeeAutoButton.exe")

New-Shortcut `
    -Path (Join-Path $programsRoot "$appName.lnk") `
    -Target (Join-Path $installDir "CoffeeAutoButton.exe")

New-Shortcut `
    -Path (Join-Path $programsRoot "CoffeeAutoButton.lnk") `
    -Target (Join-Path $installDir "CoffeeAutoButton.exe")

New-Shortcut `
    -Path (Join-Path $programsDir "Uninstall $appName.lnk") `
    -Target "powershell.exe" `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$installDir\uninstall.ps1`""

Start-Process -FilePath (Join-Path $installDir "CoffeeAutoButton.exe")
'@

    Set-Content -Path (Join-Path $PayloadDir "uninstall.ps1") -Encoding UTF8 -Value @'
$ErrorActionPreference = "Stop"

$appName = "Coffee AutoButton"
$appFolder = "CoffeeAutoButton"
$installDir = Join-Path $env:LOCALAPPDATA $appFolder
$programsRoot = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
$programsDir = Join-Path $programsRoot $appName
$programShortcut = Join-Path $programsRoot "$appName.lnk"
$programShortcutAlias = Join-Path $programsRoot "CoffeeAutoButton.lnk"

if (Test-Path $programsDir) {
    Remove-Item -LiteralPath $programsDir -Recurse -Force
}

if (Test-Path $programShortcut) {
    Remove-Item -LiteralPath $programShortcut -Force
}

if (Test-Path $programShortcutAlias) {
    Remove-Item -LiteralPath $programShortcutAlias -Force
}

if (Test-Path $installDir) {
    Remove-Item -LiteralPath $installDir -Recurse -Force
}

Write-Host "$appName was uninstalled. User settings in %APPDATA%\CoffeeAutoButton were not removed."
'@
}

function Build-IExpressInstaller([string]$PayloadDir, [string]$OutputFullPath) {
    $iexpress = Get-Command iexpress.exe -ErrorAction SilentlyContinue
    if ($iexpress -eq $null) {
        throw "IExpress was not found. Install Inno Setup or use a Windows environment with iexpress.exe."
    }

    Assert-UnderRoot $OutputFullPath
    $outputParent = Split-Path -Parent $OutputFullPath
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null

    $sedPath = Join-Path (Split-Path -Parent $PayloadDir) "CoffeeAutoButton.iexpress.sed"
    Assert-UnderRoot $sedPath

    $payloadFiles = Get-ChildItem -LiteralPath $PayloadDir -File | Sort-Object Name
    $fileStrings = New-Object System.Collections.Generic.List[string]
    $sourceFileLines = New-Object System.Collections.Generic.List[string]

    for ($index = 0; $index -lt $payloadFiles.Count; $index++) {
        $fileStrings.Add("FILE$index=`"$($payloadFiles[$index].Name)`"")
        $sourceFileLines.Add("%FILE$index%=")
    }

    $sed = @"
[Version]
Class=IEXPRESS
SEDVersion=3

[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=%InstallPrompt%
DisplayLicense=%DisplayLicense%
FinishMessage=%FinishMessage%
TargetName=%TargetName%
FriendlyName=%FriendlyName%
AppLaunched=%AppLaunched%
PostInstallCmd=%PostInstallCmd%
AdminQuietInstCmd=%AdminQuietInstCmd%
UserQuietInstCmd=%UserQuietInstCmd%
SourceFiles=SourceFiles

[Strings]
InstallPrompt=Install Coffee AutoButton?
DisplayLicense=
FinishMessage=Coffee AutoButton has been installed.
TargetName=$OutputFullPath
FriendlyName=Coffee AutoButton Setup
AppLaunched=install.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=install.cmd
UserQuietInstCmd=install.cmd
$($fileStrings -join "`r`n")

[SourceFiles]
SourceFiles0=$PayloadDir\

[SourceFiles0]
$($sourceFileLines -join "`r`n")
"@

    Set-Content -Path $sedPath -Encoding ASCII -Value $sed
    & $iexpress.Source /N /Q $sedPath
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Write-Host "Built IExpress installer: $OutputFullPath"
}

function Build-InnoInstaller([string]$PublishFullPath, [string]$OutputFullDir, [string]$Version) {
    $iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($iscc -eq $null) {
        throw "ISCC.exe was not found. Install Inno Setup or use -InstallerType IExpress."
    }

    Assert-UnderRoot $OutputFullDir
    New-Item -ItemType Directory -Path $OutputFullDir -Force | Out-Null

    $issPath = Resolve-RepoPath "installer\CoffeeAutoButton.iss"
    & $iscc.Source `
        "/DPublishDir=$PublishFullPath" `
        "/DOutputDir=$OutputFullDir" `
        "/DAppVersion=$Version" `
        $issPath
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }

    Write-Host "Built Inno Setup installer: $(Join-Path $OutputFullDir 'CoffeeAutoButtonSetup.exe')"
}

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$publishFullPath = Resolve-RepoPath $PublishPath
$outputFullDir = Resolve-RepoPath $OutputDir
$setupExe = Join-Path $outputFullDir "CoffeeAutoButtonSetup.exe"

$publishArgs = @{
    Configuration = $Configuration
    Runtime = $Runtime
    OutputPath = $publishFullPath
}
if ($SelfContained) {
    $publishArgs.SelfContained = $true
}

& (Join-Path $PSScriptRoot "publish.ps1") @publishArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

if ($InstallerType -eq "Auto") {
    $InstallerType = if (Get-Command ISCC.exe -ErrorAction SilentlyContinue) { "Inno" } else { "IExpress" }
}

if ($InstallerType -eq "Inno") {
    Build-InnoInstaller -PublishFullPath $publishFullPath -OutputFullDir $outputFullDir -Version $Version
} else {
    $payloadDir = Resolve-RepoPath "installer\payload"
    New-InstallerPayload -PublishFullPath $publishFullPath -PayloadDir $payloadDir
    Build-IExpressInstaller -PayloadDir $payloadDir -OutputFullPath $setupExe
}
