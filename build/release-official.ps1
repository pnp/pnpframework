#Requires -RunAsAdministrator

# Script used to release an official PnP Framework build
$versionIncrement = Get-Content "$PSScriptRoot\version.release.increment" -Raw
$versionIncrement = $versionIncrement -as [int]
$versionIncrement++

$version = Get-Content "$PSScriptRoot\version.release" -Raw

$version = $version.Replace("{minorrelease}", $versionIncrement)

# Build the release version
Write-Host "Building PnP.Framework version $version"
dotnet build $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --configuration Release --no-incremental --force --nologo /p:Version=$version

# Sign the binaries
d:\github\SharePointPnP\CodeSigning\PnP\sign-pnpbinaries.ps1 -SignJson pnpframeworkassemblies

# Package the release version
Write-Host "Packinging PnP.Framework version $version"
dotnet pack $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --configuration Release --no-build /p:PackageVersion=$version

# Copy to the package name used in the json sign file
copy-item d:\github\pnpframework\src\lib\PnP.Framework\bin\release\PnP.Framework.$version.nupkg d:\github\pnpframework\src\lib\PnP.Framework\bin\release\PnP.Framework.nupkg -Force

# Sign the nuget package
d:\github\SharePointPnP\CodeSigning\PnP\sign-pnpbinaries.ps1 -SignJson pnpframeworknuget

# Publish
# manual

# Persist last used version
Write-Host "Writing $version to git"
# Set-Content -Path .\version.release.increment -Value $versionIncrement

# Push to the repo
Write-Host "Pushing updated version file to git"
# git add .\version.release.increment
# git commit -m "Build increment - release version $versionIncrement"
# git push