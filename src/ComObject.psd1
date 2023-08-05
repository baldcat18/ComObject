@{
	ModuleVersion = '0.2.0'
	GUID = 'afe8c49d-9b7b-4740-872f-6d926c4272e9'
	Author = 'BaldCat'
	Copyright = '(c) 2022 BaldCat. All rights reserved.'
	Description = 'Create and release COM objects..'
	PowerShellVersion = '5.1'
	CompatiblePSEditions = @('Core', 'Desktop')
	RootModule = 'ComObject.psm1'
	FunctionsToExport = @('Get-ComObject', 'New-ComObjectFromCLSID', 'Remove-ComObject')
	CmdletsToExport = @()
	AliasesToExport = @()
	PrivateData = @{
		PSData = @{
			Prerelease = 'beta'
			Tags = @('COM', 'Windows')
		}
	}
}
