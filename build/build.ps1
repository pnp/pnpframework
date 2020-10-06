$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$versionIncrement = Get-Content .\version.debug.increment -Raw
$versionIncrement = $versionIncrement -as [int]
$versionIncrement++

$version = Get-Content .\version.debug -Raw

$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework .NET Standard 2.0 version $version"
dotnet build ./src/lib/PnP.Framework/PnP.Framework.csproj --no-incremental /p:Version=$version

Write-Host "Packinging PnP.Framework .NET Standard 2.0 version $version"
dotnet pack ./src/lib/PnP.Framework/PnP.Framework.csproj --no-build /p:PackageVersion=$version

Write-Host "Publishing to nuget"
$nupkg = $("./src/lib/PnP.Framework/bin/Debug/PnP.Framework.$version.nupkg")
$apiKey = $("$env:NUGET_API_KEY")

dotnet nuget push $nupkg --api-key $apiKey --source https://api.nuget.org/v3/index.json

Write-Host "Writing $version to git"
Set-Content -Path ./build/version.debug.increment -Value $versionIncrement