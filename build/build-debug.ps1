Param(
	[Parameter(Mandatory = $false, ValueFromPipeline = $false)]
	[switch]
	$NoIncremental,
	[Parameter(Mandatory = $false, ValueFromPipeline = $false)]
	[switch]
    $Force,
	[Parameter(Mandatory = $false,
		ValueFromPipeline = $false)]
	[switch]
	$LocalPnPCore
)

$localPnPCoreSdkPathValue = $env:PnPCoreSdkPath
$env:PnPCoreSdkPath = ""

$versionIncrement = Get-Content "$PSScriptRoot\version.debug.increment" -Raw
$version = Get-Content "$PSScriptRoot\version.debug" -Raw
$version = $version.Replace("{incremental}", $versionIncrement)

Write-Host "Building PnP.Framework version $version"

$buildCmd = "dotnet build `"$PSScriptRoot/../src/lib/PnP.Framework/PnP.Framework.csproj`"" + "--nologo --configuration Debug -p:VersionPrefix=$version -p:VersionSuffix=debug";

if ($LocalPnPCore) {
	# Check if available
	$pnpCoreAssembly = Join-Path $PSScriptRoot -ChildPath "..\..\pnpcore\src\sdk\PnP.Core\bin\Debug\netstandard2.0\PnP.Core.dll"
	$pnpCoreAssembly5 = Join-Path $PSScriptRoot -ChildPath "..\..\pnpcore\src\sdk\PnP.Core\bin\Debug\net5.0\PnP.Core.dll"
	$pnpCoreAssembly6 = Join-Path $PSScriptRoot -ChildPath "..\..\pnpcore\src\sdk\PnP.Core\bin\Debug\net6.0\PnP.Core.dll"
	$pnpCoreAssembly7 = Join-Path $PSScriptRoot -ChildPath "..\..\pnpcore\src\sdk\PnP.Core\bin\Debug\net7.0\PnP.Core.dll"
	$pnpCoreAssembly = [System.IO.Path]::GetFullPath($pnpCoreAssembly)
	$pnpCoreAssembly5 = [System.IO.Path]::GetFullPath($pnpCoreAssembly5)
	$pnpCoreAssembly6 = [System.IO.Path]::GetFullPath($pnpCoreAssembly6)
	$pnpCoreAssembly7 = [System.IO.Path]::GetFullPath($pnpCoreAssembly7)
	if (Test-Path $pnpCoreAssembly -PathType Leaf) {
		$buildCmd += " -p:PnPCoreSdkPath=`"$pnpCoreAssembly`""
		$buildCmd += " -p:PnPCoreSdkPathNet5=`"$pnpCoreAssembly5`""
		$buildCmd += " -p:PnPCoreSdkPathNet6=`"$pnpCoreAssembly6`""
		$buildCmd += " -p:PnPCoreSdkPathNet7=`"$pnpCoreAssembly7`""
	} 
	else {
		Write-Error -Message "PnP Core Assembly path $pnpCoreAssembly not found"
		exit 1
	}
}
else {
	$localFolder = Join-Path $PSScriptRoot -ChildPath "..\..\pnpcore"
	$localFolder = [System.IO.Path]::GetFullPath($localFolder)
	Write-Error -Message "Please make sure you have a local copy of the PnP.Core repository installed at $localFolder"
}

if ($NoIncremental) {
	$buildCmd += " --no-incremental";
}
if ($Force) {
	$buildCmd += " --force"
}

Write-Host "Executing $buildCmd" -ForegroundColor Yellow

Invoke-Expression $buildCmd

$env:PnPCoreSdkPath = $localPnPCoreSdkPathValue
