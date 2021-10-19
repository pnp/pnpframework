# Script used in the github nightly release workflow
$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$versionIncrement = Get-Content ./build/version.debug.increment -Raw
$versionIncrement = $versionIncrement -as [int]
$versionIncrement++

$version = Get-Content ./build/version.debug -Raw

$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework version $version"
dotnet build ./src/lib/PnP.Framework/PnP.Framework.csproj --configuration Release --no-incremental --force /p:Version=$version

Write-Host "Packinging PnP.Framework version $version"
dotnet pack ./src/lib/PnP.Framework/PnP.Framework.csproj --configuration Release --no-build /p:PackageVersion=$version

Write-Host "Publishing to nuget"
$nupkg = $("./src/lib/PnP.Framework/bin/Release/PnP.Framework.$version.nupkg")
$apiKey = $("$env:NUGET_API_KEY")

dotnet nuget push $nupkg --api-key $apiKey --source https://api.nuget.org/v3/index.json

Write-Host "Writing $version to git"
Set-Content -Path ./build/version.debug.increment -Value $versionIncrement -NoNewline