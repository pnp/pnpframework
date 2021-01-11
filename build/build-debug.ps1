$versionIncrement = Get-Content "$PSScriptRoot\version.debug.increment" -Raw
$version = Get-Content "$PSScriptRoot\version.debug" -Raw
$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework version $version"
dotnet build $PSScriptRoot\..\src\lib\PnP.Framework\PnP.Framework.csproj --no-incremental /p:Version=$version
