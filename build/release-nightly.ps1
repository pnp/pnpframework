# Script used to test the nighly release script as local version (= build.ps1 in the GitHub workflow)
$versionIncrement = Get-Content "$PSScriptRoot\version.debug.increment" -Raw
$versionIncrement = $versionIncrement -as [int]
$versionIncrement++

$version = Get-Content "$PSScriptRoot\version.debug" -Raw

$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework version $version"
dotnet build $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --no-incremental /p:Version=$version

Write-Host "Packinging PnP.Framework version $version"
dotnet pack $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --no-build /p:PackageVersion=$version

Write-Host "Writing $version to git"
#Set-Content -Path .\version.debug.increment -Value $versionIncrement

#Push to the repo
Write-Host "Pushing updated version file to git"
# git add .\version.debug.increment
# git commit -m "Build increment - debug version $versionIncrement"
# git push