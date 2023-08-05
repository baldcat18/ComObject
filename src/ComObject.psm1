using namespace System.Runtime.InteropServices

function Get-ComObject {
	<#
	.SYNOPSIS
	Returns a reference to an Automation object from a file.
	.DESCRIPTION
	Returns a reference to an Automation object from a file.
	.PARAMETER PathName
	Full path and name of the file containing the object to retrieve.
	.OUTPUTS
	System.__ComObject
	.EXAMPLE
	PS >Get-ComObject "C:\foo.html"
	#>

	[CmdletBinding()]
	[OutputType([__ComObject])]
	param (
		[parameter(Mandatory)]
		[string]$PathName
	)

	try {
		return [Marshal]::BindToMoniker($PathName)
	} catch {
		$PSCmdlet.WriteError($PSItem)
	}
}

function New-ComObjectFromCLSID {
	<#
	.SYNOPSIS
	Create the instance of the type associated with the specified CLSID.
	.DESCRIPTION
	Create the instance of the type associated with the specified CLSID.
	.PARAMETER Clsid
	The CLSID of the object to get.
	.OUTPUTS
	System.__ComObject
	.EXAMPLE
	PS >New-ComObjectFromCLSID "{72C24DD5-D70A-438B-8A42-98424B88AFB8}"
	#>

	[CmdletBinding(SupportsShouldProcess)]
	[OutputType([__ComObject])]
	param (
		[parameter(Mandatory)]
		[Guid]$Clsid
	)

	if ($PSCmdlet.ShouldProcess($Clsid)){
		try {
			return [Marshal]::BindToMoniker("new:$Clsid")
		} catch {
			$PSCmdlet.WriteError($PSItem)
		}
	}
}

function Remove-ComObject {
	<#
	.SYNOPSIS
	Decrements the reference count of the Runtime Callable Wrapper (RCW) associated with the specified COM object.
	.DESCRIPTION
	Decrements the reference count of the Runtime Callable Wrapper (RCW) associated with the specified COM object.
	.PARAMETER Object
	The COM object to release.
	.PARAMETER Force
	Releases all references to a RCW by setting its reference count to 0.
	.OUTPUTS
	System.Int32
		 The new value of the reference count of the RCW associated with Object.
	.EXAMPLE
	PS >$excel = New-Object -ComObject Excel.Application
	# Processing with $excel
	PS >Remove-ComObject $excel
	#>

	[CmdletBinding(SupportsShouldProcess)]
	[OutputType([int])]
	param (
		[parameter(Mandatory, ValueFromPipeline, Position = 0)]
		[__ComObject]$Object,
		[switch]$Force
	)

	process {
		try {
			if ($PSCmdlet.ShouldProcess($Object)) {
				if ($Force) {
					return [Marshal]::FinalReleaseComObject($Object)
				} else {
					return [Marshal]::ReleaseComObject($Object)
				}
			}
		} catch {
			$PSCmdlet.WriteError($PSItem)
		}
	}
}
