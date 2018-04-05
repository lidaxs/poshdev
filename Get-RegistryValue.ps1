<#
	version 1.0.0
	Initial upload
	Added aliases to ClientName parameter to support pipeline in from WMI,SCCM & Active Directory
#>
Function Get-RegistryValue {
	<#
		.SYNOPSIS
			Gets a registry values remote or local.

		.DESCRIPTION
			Gets a registry values remote or local.

		.PARAMETER  ComputerName
			The ComputerName(s) on which to operate.

		.PARAMETER  Hive
			The Hive to pick ...LocalMachine.

		.PARAMETER  Key
			The key to query.

		.EXAMPLE
			Get-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName "MyTestValue"
			
		.INPUTS
			System.String,System.String[]

		.OUTPUTS
			System.String

		.NOTES
			Additional information about the function go here.

		.LINK
			about_functions_advanced

		.LINK
			about_comment_based_help
#Requires –Version 3
	#>
	[CmdletBinding()]
	[OutputType([System.Object])]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

		[Parameter(Position=1, Mandatory=$false)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		[ValidateSet('ClassesRoot','CurrentConfig','CurrentUser','DynData','LocalMachine','PerformanceData','Users')]
		$Hive='LocalMachine',
		
		[Parameter(Position=2, Mandatory=$false)]
		[System.String]
		$Key,
		
		[Parameter(Position=3, Mandatory=$false)]
		[System.String]
		$ValueName

	)

begin{

	}

process{

		# Loop through collection
		ForEach($Computer in $ClientName){

			# Test connectivity
			if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue){
			
				$note=New-Object -TypeName PSObject | Select-Object Computer,Hive,Key,ValueName,Value
				Write-Verbose "Workstation $Computer is online..."
				
				try {
					# Connect to remoteregistry servive through WMI
					$svc=Get-WmiObject -Class Win32_Service -Filter "Name='RemoteRegistry'" -ComputerName $Computer -ErrorAction Stop
				}
				catch [System.Runtime.InteropServices.COMException] {
					Write-Warning "Cannot connect to $Computer through WMI"
					#$Error[0].Exception | Select Source,Message,ErrorCode | Format-List -Property Source,Message,ErrorCode
					}

				try{
					# Start service if stopped
					if($svc.State -eq 'Stopped'){
					[void]$svc.StartService()
					Write-Verbose "Succesfully started $($svc.Name)"}
				}
				catch [System.Management.Automation.RuntimeException]{
					#$Error[0].Exception | Select Source,Message,ErrorCode | Format-List -Property Source,Message,ErrorCode
					Write-Warning "Cannot start RemoteRegistry service on $Computer."
				}

				try{
					$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$Hive,$Computer,[Microsoft.Win32.RegistryView]::Registry64)
				}
				catch [System.Management.Automation.MethodInvocationException]{
					#$Error[0].Exception | Select Source,Message,ErrorCode | Format-List -Property Source,Message,ErrorCode
					Write-Warning "Cannot open RegistryKey $Key on $Computer"
				}

				$SubKey=$RemoteRegistry.OpenSubKey($Key)

				try{
					$RegValue=$SubKey.GetValue($ValueName)
					$note.Computer=$Computer
					$note.Hive=$Hive
					$note.Key=$Key
					$note.ValueName=$ValueName
					$note.Value=$RegValue
				}
				catch [System.Management.Automation.RuntimeException] {
					Write-Warning "Cannot query value $ValueName in $Key on $Computer"
					$note.Computer=$Computer
					$note.Hive=$Hive
					$note.Key=$Key
					$note.ValueName=$ValueName
					$note.Value="Unknown!!"
				}



			Write-Output $note
			
			}
		else{
				Write-Warning "$Computer is not online!"
			}
		}
	}
	
end	{
	try{
		Clear-Variable RegValue -Force -ErrorAction SilentlyContinue
		Clear-Variable SubKey -Force -ErrorAction SilentlyContinue
		Clear-Variable RemoteRegistry -Force -ErrorAction SilentlyContinue
		}
	catch{}
	}
}