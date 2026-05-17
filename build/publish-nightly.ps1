# Script used in the github nightly release workflow to publish a pre-built NuGet package
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$packageZip = "./build/package/PnP.Framework-packages.zip"
$extractPath = "./build/package/extracted"

Write-Host "Extracting $packageZip"
Expand-Archive -Path $packageZip -DestinationPath $extractPath -Force

$nupkg = Get-ChildItem -Path $extractPath -Filter "*.nupkg" -Recurse | Select-Object -First 1

if ($null -eq $nupkg) {
    Write-Error "No .nupkg file found in $extractPath"
    exit 1
}

Write-Host "Publishing $($nupkg.Name) to NuGet"
$apiKey = $env:NUGET_API_KEY

dotnet nuget push $nupkg.FullName --api-key $apiKey --source https://api.nuget.org/v3/index.json
