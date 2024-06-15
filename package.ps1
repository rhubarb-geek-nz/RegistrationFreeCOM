# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param(
	$CertificateThumbprint = '601A8B683F791E51F647D34AD102C38DA4DDB65F'
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

trap
{
	throw $PSItem
}

foreach ($EDITION in 'Community', 'Professional')
{
	$VCVARSDIR = "${Env:ProgramFiles}\Microsoft Visual Studio\2022\$EDITION\VC\Auxiliary\Build"

	if ( Test-Path -LiteralPath $VCVARSDIR -PathType Container )
	{
		break
	}
}

$VCVARSARM = 'vcvarsarm.bat'
$VCVARSARM64 = 'vcvarsarm64.bat'
$VCVARSAMD64 = 'vcvars64.bat'
$VCVARSX86 = 'vcvars32.bat'
$VCVARSHOST = 'vcvars32.bat'

switch ($Env:PROCESSOR_ARCHITECTURE)
{
	'AMD64' {
		$VCVARSX86 = 'vcvarsamd64_x86.bat'
		$VCVARSARM = 'vcvarsamd64_arm.bat'
		$VCVARSARM64 = 'vcvarsamd64_arm64.bat'
		$VCVARSHOST = $VCVARSAMD64
	}
	'ARM64' {
		$VCVARSX86 = 'vcvarsarm64_x86.bat'
		$VCVARSARM = 'vcvarsarm64_arm.bat'
		$VCVARSAMD64 = 'vcvarsarm64_amd64.bat'
		$VCVARSHOST = $VCVARSARM64
	}
	'X86' {
		$VCVARSXARM64 = 'vcvarsx86_arm64.bat'
		$VCVARSARM = 'vcvarsx86_arm.bat'
		$VCVARSAMD64 = 'vcvarsx86_amd64.bat'
	}
	Default {
		throw "Unknown architecture $Env:PROCESSOR_ARCHITECTURE"
	}
}

$VCVARSARCH = @{'arm' = $VCVARSARM; 'arm64' = $VCVARSARM64; 'x86' = $VCVARSX86; 'x64' = $VCVARSAMD64}

$ARCHLIST = ( $VCVARSARCH.Keys | ForEach-Object {
	$VCVARS = $VCVARSARCH[$_];
	if ( Test-Path -LiteralPath "$VCVARSDIR/$VCVARS" -PathType Leaf )
	{
		$_
	}
} | Sort-Object )

$ARCHLIST | ForEach-Object {
	New-Object PSObject -Property @{
		Architecture=$_;
		Environment=$VCVARSARCH[$_]
	}
} | Format-Table -Property Architecture,'Environment'

foreach ($DIR in 'obj', 'bin')
{
	if (Test-Path -LiteralPath $DIR)
	{
		Remove-Item -LiteralPath $DIR -Force -Recurse
	}
}

$ARCHLIST | ForEach-Object {
	$ARCH = $_

	$VCVARS = ( '{0}\{1}' -f $VCVARSDIR, $VCVARSARCH[$ARCH] )

	@"
CALL "$VCVARS"
IF ERRORLEVEL 1 EXIT %ERRORLEVEL%
NMAKE /NOLOGO clean
IF ERRORLEVEL 1 EXIT %ERRORLEVEL%
NMAKE /NOLOGO CertificateThumbprint="$CertificateThumbprint"
EXIT %ERRORLEVEL%
"@ | & "$env:COMSPEC"

	if ($LastExitCode -ne 0)
	{
		exit $LastExitCode
	}
}

$ARCHLIST | ForEach-Object {
	$ARCH = $_
	$VCVARS = ( '{0}\{1}' -f $VCVARSDIR, $VCVARSARCH[$ARCH] )
	$EXE = "bin\$ARCH\displib.dll"

	"bin\$ARCH\displib.dll", "bin\$ARCH\dispapp.exe" | ForEach-Object {
		$EXE = $_

		$MACHINE = ( @"
@CALL "$VCVARS" > NUL:
IF ERRORLEVEL 1 EXIT %ERRORLEVEL%
dumpbin /headers $EXE
IF ERRORLEVEL 1 EXIT %ERRORLEVEL%
EXIT %ERRORLEVEL%
"@ | & "$env:COMSPEC" /nologo /Q | Select-String -Pattern " machine " )

		$MACHINE = $MACHINE.ToString().Trim()

		$MACHINE = $MACHINE.Substring($MACHINE.LastIndexOf(' ')+1)

		New-Object PSObject -Property @{
			Architecture=$ARCH;
			Executable=$EXE;
			Machine=$MACHINE;
			FileVersion=(Get-Item $EXE).VersionInfo.FileVersion;
			ProductVersion=(Get-Item $EXE).VersionInfo.ProductVersion;
			FileDescription=(Get-Item $EXE).VersionInfo.FileDescription
		}
	}
} | Format-Table -Property Architecture, Executable, Machine, FileVersion, ProductVersion, FileDescription

$null = New-Item -Path '.' -Name 'base' -ItemType 'directory'

try
{
	$null = New-Item -Path 'base' -Name 'runtimes' -ItemType 'directory'

	foreach ($ARCH in $ARCHLIST)
	{

		$null = New-Item -Path 'base\runtimes' -Name "win-$ARCH" -ItemType 'directory'

		$null = New-Item -Path "base\runtimes\win-$ARCH" -Name 'native' -ItemType 'directory'

		Copy-Item "bin\$ARCH\displib.dll" "base\runtimes\win-$ARCH\native\displib.dll"
	}

	Copy-Item 'README.md' 'base\README.md'

	& nuget pack 'displib\displib.nuspec' -BasePath 'base'

	If ( $LastExitCode -ne 0 )
	{
		Exit $LastExitCode
	}
}
finally
{
	Remove-Item 'base' -Recurse
}
