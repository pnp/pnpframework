$versionIncrement = Get-Content .\version.debug.increment -Raw
$versionIncrement = $versionIncrement -as [int]
$versionIncrement++

$version = Get-Content .\version.debug -Raw

$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework.Net Standard 2.0 version $version"
dotnet build ..\src\lib\PnP.Framework\PnP.Framework.csproj --no-incremental /p:Version=$version

Write-Host "Packinging PnP.Core .Net Standard 2.0 version $version"
dotnet pack ..\src\lib\PnP.Framework\PnP.Framework.csproj --no-build /p:PackageVersion=$version

#Write-Host "Writing $version to git"
#Set-Content -Path .\version.debug.increment -Value $versionIncrement

#Push to the repo
# git add .\version.debug.increment
# git commit -m "Build increment - debug version $versionIncrement"
# git push