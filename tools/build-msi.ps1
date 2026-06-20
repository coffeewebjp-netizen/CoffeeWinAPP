param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "1.0.10",
    [string]$PublishPath = "publish\CoffeeAutoButton",
    [string]$OutputPath = "installer\out\CoffeeAutoButtonSetup.msi",
    [switch]$InstallWixTool
)

$ErrorActionPreference = "Stop"

function Resolve-RepoPath([string]$Path) {
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return Join-Path $script:Root.Path $Path
}

function Resolve-WixExe {
    $localWix4 = Resolve-RepoPath ".tools\wix4\wix.exe"
    if (Test-Path $localWix4) {
        return $localWix4
    }

    if ($InstallWixTool) {
        $toolPath = Resolve-RepoPath ".tools\wix4"
        New-Item -ItemType Directory -Path $toolPath -Force | Out-Null
        dotnet tool install wix --version 4.0.6 --tool-path $toolPath | ForEach-Object { Write-Host $_ }
        if ($LASTEXITCODE -ne 0) {
            exit $LASTEXITCODE
        }

        if (Test-Path $localWix4) {
            return $localWix4
        }
    }

    $pathCommand = Get-Command wix.exe -ErrorAction SilentlyContinue
    if ($pathCommand -ne $null) {
        return $pathCommand.Source
    }

    throw "WiX Toolset was not found. Install it with: dotnet tool install wix --version 4.0.6 --tool-path .\.tools\wix4"
}

function Ensure-WixUiExtension([string]$WixExe) {
    $extensions = & $WixExe extension list
    if ($extensions -match "WixToolset\.UI\.wixext") {
        return
    }

    & $WixExe extension add WixToolset.UI.wixext/4.0.6 | ForEach-Object { Write-Host $_ }
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
}

$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$publishFullPath = Resolve-RepoPath $PublishPath
$outputFullPath = Resolve-RepoPath $OutputPath
$outputDir = Split-Path -Parent $outputFullPath
$wxsPath = Resolve-RepoPath "installer\msi\Product.wxs"

& (Join-Path $PSScriptRoot "publish.ps1") `
    -Configuration $Configuration `
    -Runtime $Runtime `
    -OutputPath $publishFullPath
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

$wixExe = Resolve-WixExe
Ensure-WixUiExtension $wixExe
& $wixExe build `
    $wxsPath `
    -arch x64 `
    -ext WixToolset.UI.wixext `
    -d "PublishDir=$publishFullPath" `
    -d "AppVersion=$Version" `
    -o $outputFullPath
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "Built MSI installer: $outputFullPath"
