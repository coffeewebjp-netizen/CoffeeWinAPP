param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [switch]$SelfContained,
    [string]$OutputPath = "publish\CoffeeAutoButton"
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$projectPath = Join-Path $root.Path "CoffeeAutoButton\CoffeeAutoButton.csproj"
$outputFullPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath
} else {
    Join-Path $root.Path $OutputPath
}

$publishArgs = @(
    "publish",
    $projectPath,
    "-c",
    $Configuration,
    "-r",
    $Runtime,
    "--self-contained",
    $SelfContained.IsPresent.ToString().ToLowerInvariant(),
    "-p:NuGetAudit=false",
    "-o",
    $outputFullPath
)

dotnet @publishArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "Published Coffee AutoButton to $outputFullPath"
