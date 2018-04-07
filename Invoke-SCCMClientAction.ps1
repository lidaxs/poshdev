<#
	version 1.0.1
	Added Aliases to ClientName parameter to support pipeline in from WMI,SCCM & Active Directory

	version 1.0.0
	Initial upload

	wishlist
	Usage of invoke-wmimethod
	one parameter actions for all switches(hashtable?)
	$PSBoundParameters
	replace -asjob with multithreading functionality
	replace test-connection with more speedy option(bucket fill in processblock?)
#>
Function Invoke-SCCMClientAction {
	<#
		.SYNOPSIS
			Runs sccm client actions on local\remote workstations.

		.DESCRIPTION
			Runs sccm client actions on local\remote workstations.
			Possible actions:HardwareInventory,SoftwareInventory,DiscoveryDataCollection,FileInventory,SoftwareUpdatesScan,SoftwareUpdatesDeployment

		.PARAMETER ClientName
			The ComputerName(s) on which to operate.(Accepts value from pipeline)

		.EXAMPLE
			Invoke-SCCMClientAction -ClientName 'C120VMXP','C120WIN7'

		.EXAMPLE
			'C120VMXP','C120WIN7' | Invoke-SCCMClientAction

		.EXAMPLE
			Invoke-SCCMClientAction -ClientName (Get-Content C:\computers.txt)

		.EXAMPLE
			(Get-Content C:\computers.txt) | Invoke-SCCMClientAction

		.INPUTS
			[System.String],[System.String[]],[System.Boolean]

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
		$RequestEvaluateMachinePolicies,
		
		[Switch]
		$RequestMachineAssignments,
		
		[Switch]
		$HardwareInventory,
		
		[Switch]
		$SoftwareInventory,
		
		[Switch]
		$DiscoveryDataCollection,
		
		[Switch]
		$FileCollection,
		
		[Switch]
		$UpdateDeployment,
		
		[Switch]
		$UpdateScan,
	
		[Switch]
		$Full,
		
		[Switch]
		$AsJob

	)

	# set initial values in the begin block (populate variables, check dependent modules etc.)
	begin {

	} # end beginblock

	# processblock
	process {

		# add -Whatif and -Confirm support to the CmdLet
		if($PSCmdlet.ShouldProcess("$ClientName", "Invoke-SCCMClientAction")){

			# loop through collection $ClientName
			ForEach($Computer in $ClientName){

				# test connection to each $Computer
				if ( Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {

					Write-Verbose "$Computer is online..."


######### START SCRIPTBLOCK ######################

$ScriptBlock = {
[CmdletBinding()]
	param(
		[System.String]
		$Computer,

		[System.Boolean]
		$RequestEvaluateMachinePolicies,

		[System.Boolean]
		$RequestMachineAssignments,		
		
		[System.Boolean]
		$HardwareInventory,
		
		[System.Boolean]
		$SoftwareInventory,
		
		[System.Boolean]
		$DiscoveryDataCollection,
		
		[System.Boolean]
		$FileCollection,
		
		[System.Boolean]
		$UpdateDeployment,
		
		[System.Boolean]
		$UpdateScan,
	
		[System.Boolean]
		$Full
	)

	# start try connecting to SMS_Client class
	try{
		$SMSCli = [wmiclass]"\\$Computer\root\ccm:SMS_Client"
	} # end try

	# start catch specific
	catch [System.Runtime.InteropServices.COMException] {
		Write-Warning "Cannot connect to $Computer through WMI"
		Write-Warning $Error[0].Exception.Message
	} # end catch specific error

	# catch rest of errors
	catch {
		Write-Warning $Error[0].Exception.Message
	} # end catch rest of errors


	if($HardwareInventory){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000001}"
		$scheduleFriendlyName="Hardware Inventory Cycle"
		
		$objWmi=[wmi]"\\$Computer\root\ccm\invagt:InventoryActionStatus.InventoryActionID='$($scheduleID)'"
		
		$LastCycleStarted=$objWmi.ConvertToDateTime($objWmi.LastCycleStartedDate)
		Write-Verbose "LastCycleStartedDate $LastCycleStarted"
		
		$LastReportDate=$objWmi.ConvertToDateTime($objWmi.LastReportDate)
		Write-Verbose "LastReportDate $LastReportDate"
		
		$message="Invoking $scheduleFriendlyName (Delta)"
		
		if($Full){
			try{
				$message=$message.Replace("Delta","Full")
				Write-Verbose "Deleting previous inventory data for $scheduleFriendlyName...performing full resync(This can take a while)"
				$objWmi.Delete()
			}
			catch [System.Management.Automation.RuntimeException]{
				Write-Verbose "InventoryAction $scheduleFriendlyName is already deleted probably"
				Write-Warning $Error[0].Exception.Message
			}
			catch{
				Write-Warning $Error[0].Exception.Message
			}
		}
		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule("$scheduleID")
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}

	if($SoftwareInventory){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000002}"
		$scheduleFriendlyName="Software Inventory Cycle"
		
		$objWmi=[wmi]"\\$Computer\root\ccm\invagt:InventoryActionStatus.InventoryActionID='$($scheduleID)'"
		
		$LastCycleStarted=$objWmi.ConvertToDateTime($objWmi.LastCycleStartedDate)
		Write-Verbose "LastCycleStartedDate $LastCycleStarted"
		
		$LastReportDate=$objWmi.ConvertToDateTime($objWmi.LastReportDate)
		Write-Verbose "LastReportDate $LastReportDate"
		
		$message="Invoking $scheduleFriendlyName (Delta)"
		
		if($Full){
			try{
				$message=$message.Replace("Delta","Full")
				Write-Verbose "Deleting previous inventory data for $scheduleFriendlyName...performing full resync(This can take a while)"
				$objWmi.Delete()
			}
			catch [System.Management.Automation.RuntimeException]{
				Write-Verbose "InventoryAction $scheduleFriendlyName is already deleted probably"
				Write-Warning $Error[0].Exception.Message
			}
			catch{
				Write-Warning $Error[0].Exception.Message
			}
		}
		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule("$scheduleID")
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}

	if($DiscoveryDataCollection){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000003}"
		$scheduleFriendlyName="Discovery Data Collection Cycle"
		
		$objWmi=[wmi]"\\$Computer\root\ccm\invagt:InventoryActionStatus.InventoryActionID='$($scheduleID)'"
		
		$LastCycleStarted=$objWmi.ConvertToDateTime($objWmi.LastCycleStartedDate)
		Write-Verbose "LastCycleStartedDate $LastCycleStarted"
		
		$LastReportDate=$objWmi.ConvertToDateTime($objWmi.LastReportDate)
		Write-Verbose "LastReportDate $LastReportDate"
		
		$message="Invoking $scheduleFriendlyName (Delta)"
		
		if($Full){
			try{
				$message=$message.Replace("Delta","Full")
				Write-Verbose "Deleting previous inventory data for $scheduleFriendlyName...performing full resync(This can take a while)"
				$objWmi.Delete()
			}
			catch [System.Management.Automation.RuntimeException]{
				Write-Verbose "InventoryAction $scheduleFriendlyName is already deleted probably"
				Write-Warning $Error[0].Exception.Message
			}
			catch{
				Write-Warning $Error[0].Exception.Message
			}
		}
		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule("$scheduleID")
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}

	if($FileCollection){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000010}"
		$scheduleFriendlyName="File Collection Cycle"
		
		$objWmi=[wmi]"\\$Computer\root\ccm\invagt:InventoryActionStatus.InventoryActionID='$($scheduleID)'"
		
		$LastCycleStarted=$objWmi.ConvertToDateTime($objWmi.LastCycleStartedDate)
		Write-Verbose "LastCycleStartedDate $LastCycleStarted"
		
		$LastReportDate=$objWmi.ConvertToDateTime($objWmi.LastReportDate)
		Write-Verbose "LastReportDate $LastReportDate"
		
		$message="Invoking $scheduleFriendlyName (Delta)"
		
		if($Full){
			try{
				$message=$message.Replace("Delta","Full")
				Write-Verbose "Deleting previous inventory data for $scheduleFriendlyName...performing full resync(This can take a while)"
				$objWmi.Delete()
			}
			catch [System.Management.Automation.RuntimeException]{
				Write-Verbose "InventoryAction $scheduleFriendlyName is already deleted probably"
				Write-Warning $Error[0].Exception.Message
			}
			catch{
				Write-Warning $Error[0].Exception.Message
			}
		}
		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule("$scheduleID")
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}

	if($UpdateDeployment){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000108}"
		$scheduleFriendlyName="Software Updates Deployment Cycle"
		
		$message="Invoking $scheduleFriendlyName"

		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule($scheduleID)
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}

	if($UpdateScan){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000113}"
		$scheduleFriendlyName="Software Updates Scan Cycle"

		$message="Invoking $scheduleFriendlyName"

		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule($scheduleID)
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}

	if($RequestMachineAssignments){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000021}"
		$scheduleFriendlyName="Request machine policies"

		$message="Invoking $scheduleFriendlyName"

		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule($scheduleID)
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}
	}
	
	if($RequestEvaluateMachinePolicies){
	
		$scheduleID = "{00000000-0000-0000-0000-000000000022}"
		$scheduleFriendlyName="Request and evaluate machine policies"

		$message="Invoking $scheduleFriendlyName"

		try{
			Write-Verbose "$message on $Computer"
			$return = $SMSCli.TriggerSchedule($scheduleID)
		}
		catch{
			Write-Warning $Error[0].Exception.Message
		}

	}

} # END SCRIPTBLOCK

if( $AsJob ){
	$jobname="SCCMScheduleTrigger"
	Start-Job -ScriptBlock $ScriptBlock -ArgumentList $Computer,$RequestEvaluateMachinePolicies,$RequestMachineAssignments,$HardwareInventory,$SoftwareInventory,$DiscoveryDataCollection,$FileCollection,$UpdateDeployment,$UpdateScan,$Full -Name $jobname
}
else {
	Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$RequestEvaluateMachinePolicies,$RequestMachineAssignments,$HardwareInventory,$SoftwareInventory,$DiscoveryDataCollection,$FileCollection,$UpdateDeployment,$UpdateScan,$Full
}

			} # if test-connection

			# else $Computer not online
			else {
				Write-Warning "$Computer is not online!"
				} # end test connection to each $Computer

			} # end for each $Computer

		} # end if $pscmdlet.ShouldProcess

	} # end processblock

	# remove variables  in the endblock
	end {

	} # end endblock


} # end function