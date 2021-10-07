#Requires -RunAsAdministrator

# Script used to release an official PnP Framework build
$versionIncrement = Get-Content "$PSScriptRoot\version.debug.increment" -Raw
$versionIncrement = $versionIncrement -as [int]
$versionIncrement++

$version = Get-Content "$PSScriptRoot\version.debug" -Raw

$version = $version.Replace("{incremental}", $versionIncrement)

# Build the release version
Write-Host "Building PnP.Framework version $version..."
dotnet build $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --configuration Release --no-incremental --force --nologo /p:Version=$version

# Sign the binaries
Write-Host "Signing the binaries..."
d:\github\SharePointPnP\CodeSigning\PnP\sign-pnpbinaries.ps1 -SignJson pnpframeworkassemblies

# Package the release version
Write-Host "Packinging PnP.Framework version $version..."
dotnet pack $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --configuration Release --no-build /p:PackageVersion=$version

# Sign the nuget package is not needed as Nuget signs the package automatically

# Publish
Write-host "Verify the created NuGet package in folder d:\github\pnpframework\src\lib\PnP.Framework\bin\release. If OK enter the nuget API key to publish the package, press enter to cancel." -ForegroundColor Yellow 
$apiKey = Read-Host "NuGet API key" 

if ($apiKey.Length -gt 0)
{
    # Push the actual package and the symbol package
    nuget push d:\github\pnpframework\src\lib\PnP.Framework\bin\release\PnP.Framework.$version.nupkg -ApiKey $apiKey -source https://api.nuget.org/v3/index.json

    # Persist last used version
    Write-Host "Writing $version to git"
    Set-Content -Path .\version.debug.increment -Value $versionIncrement -NoNewline

    # Push change to the repo
    Write-Host "!!Ensure you push in all changes!!" -ForegroundColor Yellow 
}
else 
{
    Write-Host "Publishing of the NuGet package cancelled!"
}