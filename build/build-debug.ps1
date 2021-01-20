Param(
	[Parameter(Mandatory = $false, ValueFromPipeline = $false)]
	[switch]
	$NoIncremental,
	[Parameter(Mandatory = $false, ValueFromPipeline = $false)]
	[switch]
    $Force
)

$versionIncrement = Get-Content "$PSScriptRoot\version.debug.increment" -Raw
$version = Get-Content "$PSScriptRoot\version.debug" -Raw
$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework version $version"

$buildCmd = "dotnet build `"$PSScriptRoot/../src/lib/PnP.Framework/PnP.Framework.csproj`"" + "--nologo --configuration Debug -p:VersionPrefix=$version -p:VersionSuffix=debug";

if ($NoIncremental) {
	$buildCmd += " --no-incremental";
}
if ($Force) {
	$buildCmd += " --force"
}

Write-Host "Executing $buildCmd" -ForegroundColor Yellow

Invoke-Expression $buildCmd

