#!/usr/bin/env pwsh
# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param($ProjectName, $IntermediateOutputPath, $OutDir, $PublishDir)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

trap
{
	throw $PSItem
}

function Get-SingleNodeValue([System.Xml.XmlDocument]$doc,[string]$path)
{
	return $doc.SelectSingleNode($path).FirstChild.Value
}

$xmlDoc = [System.Xml.XmlDocument](Get-Content "$ProjectName.csproj")

$ModuleId = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/PackageId'
$Version = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/Version'
$ProjectUri = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/PackageProjectUrl'
$Description = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/Description'
$Author = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/Authors'
$Copyright = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/Copyright'
$AssemblyName = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/AssemblyName'
$CompanyName = Get-SingleNodeValue $xmlDoc '/Project/PropertyGroup/Company'

$moduleSettings = @{
	Path = "$OutDir$ModuleId.psd1"
	RootModule = "$AssemblyName.dll"
	ModuleVersion = $Version
	Guid = 'f7ba451d-8bc2-4a5c-8e42-d251d4f99f9f'
	Author = $Author
	CompanyName = $CompanyName
	Copyright = $Copyright
	Description = $Description
	FunctionsToExport = @()
	CmdletsToExport = @('Test-RegistrationFreeCOM','New-RegistrationFreeCOM')
	VariablesToExport = '*'
	AliasesToExport = @()
	ProjectUri = $ProjectUri
}

New-ModuleManifest @moduleSettings

Import-PowerShellDataFile -LiteralPath "$OutDir$ModuleId.psd1" | Export-PowerShellDataFile | Set-Content -LiteralPath "$PublishDir$ModuleId.psd1" -Encoding utf8BOM

$NuGetRuntimePath = "$HOME\.nuget\packages\rhubarb-geek-nz.registrationfreecom\$Version\runtimes"

Get-ChildItem -LiteralPath $NuGetRuntimePath | ForEach-Object {
	$platformDir = $_.Name
	$platformPath = $_.FullName

	if (Test-Path -LiteralPath "$PublishDir\$platformDir")
	{
		Remove-Item -LiteralPath "$PublishDir\$platformDir" -Recurse
	}

	$null = New-Item $PublishDir -Name $platformDir -ItemType 'Directory'

	Copy-Item "$platformPath\native\*" "$PublishDir$platformDir\"
}
