<#
	version 1.0.0.2
	aliases not working as expected when using pipeline and piping different types of objects
	added if($Computer.Name){$Computer=$Computer.Name} in processblock
	
	version 1.0.0.1
	test connectivity now with wmi
	
	version 1.0.0
	Initial upload
	Added Aliases to ClientName parameter to support pipeline in from WMI,SCCM & Active Directory
#>
Function Get-HiXVersion {
	<#
		.SYNOPSIS
			Retieves version of EZIS where it has been copied.

		.DESCRIPTION
			A long description of the function.

		.PARAMETER ClientName
			The ComputerName(s) on which to operate.(Accepts value from pipeline)

		.EXAMPLE
			Get-HiXVersion -ComputerName 'srv-ts10','srv-ts11'

		.EXAMPLE
			'srv-ts10','srv-ts11' | Get-HiXVersion

		.EXAMPLE
			Get-HiXVersion -ClientName (Get-Content C:\computers.txt)

		.EXAMPLE
			(Get-Content C:\computers.txt) | Get-HiXVersion -Acceptatie

		.EXAMPLE
			Get-QADComputer srv-ts* | Get-HiXVersion

		.INPUTS
			[System.String],[System.String[]]

		.OUTPUTS
			[System.Management.Automation.PSObject]

		.NOTES
			Additional information about this function go here.

		.LINK
			Get-Help about_Functions_advanced

		.LINK
			Get-Help about_comment_based_help

		.SEE ALSO
			Get-Help about_Functions
			Get-Help about_Functions_Advanced_Methods
			Get-Help about_Functions_Advanced_Parameters
			Get-Help about_Functions_CmdletBindingAttribute
			Get-Help about_Functions_OutputTypeAttribute
			http://go.microsoft.com/fwlink/?LinkID=135279
	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Management.Automation.PSObject])]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

		[Switch]
		$Produktie,

		[Switch]
		$Acceptatie,

		[Switch]
		$Ontwikkel,

		[Switch]
		$Support,

		[Switch]
		$Update,

		[Switch]
		$Sedatie,

		[Switch]
		$ZCD
	)

	# set initial values in the begin block (populate variables, check dependent modules etc.)
	begin {
		if($host.Version.Major -lt '3'){
			Write-verbose 'This function works only on powershell version 3.0 and higher.`nUpgrade to .NET 4.0 and install WinRM 3.0'
			#break
		}
	} # end beginblock

	# processblock
	process {

		# add -Whatif and -Confirm support to the CmdLet
		if($PSCmdlet.ShouldProcess("$ClientName", "Get-HiXVersion")){

			if($Computer.Name){$Computer=$Computer.Name}

			# loop through collection $ComputerName
			ForEach($Computer in $ClientName){

				# test connection to each $Computer
				if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0)
				{

					Write-Verbose "$Computer is online..."

					# start try
					try{
						$objComputerSystem=Get-WmiObject -Class Win32_ComputerSystem -Property SystemType,DomainRole -ComputerName $Computer -ErrorAction Stop
					} # end try

					# start catch specific
					catch [System.Runtime.InteropServices.COMException] {
						Write-Warning "Cannot connect to $Computer through WMI"
						$Error[0].Exception.Message
					} # end catch specific error

					# catch rest of errors
					catch {
						$Error[0].Exception.Message
					} # end catch rest of errors
					

					# end # set Program Files directory according to systemtype 32-bit or 64-bit

					$ChipsoftDirectory="Chipsoft"
					
					# find out if we are dealing with server or workstation
					if($objComputerSystem.DomainRole -ge 3){
						$rootdrive="E$"
						Write-Verbose "$Computer is a server....setting rootdrive to $rootdrive"}
					else{
						$rootdrive="C$"
						Write-Verbose "$Computer is a workstation....setting rootdrive to $rootdrive"
					}


					If($Produktie){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_Produktie\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}
					If($Acceptatie){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_Acceptatie\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}
					If($Ontwikkel){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_Ontwikkel\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}
					If($Support){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_Support\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}
					If($Update){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_Update\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}
					If($Sedatie){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_Sedatie\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}
					If($ZCD){
						$Exepath="$rootdrive\$ChipsoftDirectory\HiX_ZCD\ChipSoft.FCL.ClassRegistry.dll"
						Write-Verbose "Setting path to $Exepath"
					}

					# test for existing executablepath
					if(Test-Path "\\$Computer\$ExePath"){
						
						# open the file and read the fileinfo
						$file=Get-Item "\\$Computer\$ExePath" -Force | Select-Object @{Expression={[System.Version]$_.VersionInfo.FileVersion.Replace(',','.').Split(' _')[0]};Label="FileVersion"},LastWriteTime
						
						# create psobject $output
						$output=New-Object PSObject | Select-Object ComputerName,FileVersion,LastWriteTime
						
						# write info to $output
						$output.ComputerName=$Computer
						$output.FileVersion=$file.FileVersion
						$output.LastWriteTime=$file.LastWriteTime
						
						# output info....
						Write-Output $output

					}
					else{
						Write-Warning "Filepath \\$Computer\$Exepath does not exist!"
					}
					# end test for existing executablepath
					# end do the things you want to do after this line
				
			}
			# else $Computer not online
			else {
				Write-Warning "$Computer is not online!"
				} # end test connection to each $Computer

			} # end for each $Computer

		} # end if $pscmdlet.ShouldProcess

	} # end processblock

	# remove variables  in the endblock
	end {

		Remove-Variable objComputerSystem -Force -ErrorAction SilentlyContinue

	} # end endblock

} # end function