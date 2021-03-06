<#
	version 1.0.6.3
	aliases not working as expected when using pipeline and piping different types of objects
	added if($Computer.Name){$Computer=$Computer.Name} in processblock
	
	version 1.0.6.2
	test connectivity with wmi
	
	version 1.0.6.1
	added output for multithreading in End block

	version 1.0.5
	bugfixes solved (namespace added to wmi methods)
	updated synopsis

	version 1.0.4
	test-connection replaced with faster port 139 connect
	added reversed hashtable $RequestID to get name from scheduleID
	added arraylist fullcollection
	added verbosing
	added wmi datetimeconversion

	version 1.0.3
	rework with multithread
	removed switches for each action
	added $RequestAction[] param
	changed to Invoke-WMIMethod

	todo:bucket,test-onlinefast,lastscheduledate,lastreportdate,doc

	version 1.0.2
	added hashtable $ScheduleIDs
	version 1.0.1
	Added Aliases to ClientName parameter to support pipeline in from WMI,SCCM & Active Directory
		
	version 1.0.0
	Initial upload

	wishlist
	Usage of invoke-wmimethod...done
	one parameter actions for all switches(hashtable?)...done
	replace -asjob with multithreading functionality...done
	replace test-connection with more speedy option(bucket fill in processblock?)...done
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
			Invoke-SCCMClientAction -ClientName 'C120VMXP','C120WIN7' -RequestAction RequestMachineAssignments,RequestEvaluateMachinePolicies

		.EXAMPLE
			'C120VMXP','C120WIN7' | Invoke-SCCMClientAction -RequestAction RequestMachineAssignments,RequestEvaluateMachinePolicies

		.EXAMPLE
			Invoke-SCCMClientAction -ClientName (Get-Content C:\computers.txt) -RequestAction SoftwareUpdatesDeployment

		.EXAMPLE
			(Get-Content C:\computers.txt) | Invoke-SCCMClientAction -RequestAction HardwareInventory -Full

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

		[ValidateSet('RequestMachineAssignments','RequestEvaluateMachinePolicies','SoftwareUpdatesScan','SoftwareUpdatesDeployment','HardwareInventory','SoftwareInventory','DiscoveryDataCollection','FileInventory')]
		[System.String[]]
		$RequestAction,

		[Switch]
		$Full,

		# run the script multithreaded against multiple computers
		[Parameter(Mandatory=$false)]
		[Switch]
		$MultiThread,

		# maximum number of threads that can run simultaniously
		[Parameter(Mandatory=$false)]
		[Int]
		$MaxThreads=20,

		# Maximum time(seconds) in which a thread must finish before a timeout occurs
		[Parameter(Mandatory=$false)]
		[Int]
		$MaxResultTime=20,

		[Parameter(Mandatory=$false)]
		[Int]
		$SleepTimer=1000

	)

	# set initial values in the begin block (populate variables, check dependent modules etc.)
	begin
	{

		if ($MultiThread)
		{
			Write-Verbose "Creating Default Initial Session State"
			$ISS = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
			
			Write-Verbose "Creating RunspacePool in which the threads will run"
			$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $ISS, $Host)
			
			Write-Verbose "Opening RunspacePool"
			$RunspacePool.Open()
			
			Write-Verbose "Creating Jobs array which will hold each job"
			$Jobs = @()
		}
	} # end beginblock

	# processblock
	process {

		# add -Whatif and -Confirm support to the CmdLet
		if($PSCmdlet.ShouldProcess("$ClientName", "Invoke-SCCMClientAction"))
		{

			if($Computer.Name){$Computer=$Computer.Name}

			# loop through collection $ClientName
			ForEach($Computer in $ClientName){

				# test connection to each $Computer...modify
				#if([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,139).AsyncWaitHandle.WaitOne(1000,$false))
				#if ( Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue)
				if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0)
				{

					Write-Verbose "$Computer is online..."


######### START SCRIPTBLOCK ######################

$ScriptBlock = {
[CmdletBinding(SupportsShouldProcess=$true)]
	param(
		[System.String]
		$Computer,

		[System.String[]]
		$RequestAction,

		[System.Boolean]
		$Full
	)

	# collection of schedules which have delta and full inventory
	[System.Collections.ArrayList]$FullCollection=@(
		"{00000000-0000-0000-0000-000000000001}"
		"{00000000-0000-0000-0000-000000000002}"
		"{00000000-0000-0000-0000-000000000003}"
		"{00000000-0000-0000-0000-000000000010}"
	)

	# hashtable to translate Name to ScheduleID
	$RequestActions = @{
		HardwareInventory              = "{00000000-0000-0000-0000-000000000001}"
		SoftwareInventory              = "{00000000-0000-0000-0000-000000000002}"
		DiscoveryDataCollection        = "{00000000-0000-0000-0000-000000000003}"
		FileInventory                  = "{00000000-0000-0000-0000-000000000010}"
		SoftwareUpdatesScan            = "{00000000-0000-0000-0000-000000000113}"
		SoftwareUpdatesDeployment      = "{00000000-0000-0000-0000-000000000108}"
		RequestMachineAssignments      = "{00000000-0000-0000-0000-000000000021}"
		RequestEvaluateMachinePolicies = "{00000000-0000-0000-0000-000000000022}"
	}

	# reverse the $RequestAction hashtable to make reverse lookup possible
	# hashtable to translate ScheduleID to Name
	$RequestID=@{}
	foreach ($key  in $RequestActions.Keys) {
		$RequestID.Add($RequestActions[$key],$key)
	}
	
	# start try invoking method to SMS_Client class
	try{
		foreach($ScheduleID in $RequestActions[$RequestAction])
		{
			if ($Full) {
				if ($ScheduleID -in $FullCollection) {
					Write-Verbose "Deleting previous inventory data for $($RequestID.$ScheduleID)($ScheduleID) ...performing full resync(This can take a while)"
					$Status=Get-WmiObject -ComputerName $Computer -Namespace ROOT\ccm\invagt -Class InventoryActionStatus -Filter "InventoryActionID='$ScheduleID'"
				}
			}

			Write-Verbose "Requesting action $($RequestID.$ScheduleID)($ScheduleID) for $Computer"
			Invoke-WmiMethod -Namespace ROOT\ccm -ComputerName $Computer -Class SMS_Client -Name TriggerSchedule -ArgumentList "$($ScheduleID)" | Out-Null

			if($Status)
			{
				$output = [PSCustomObject]@{
					ComputerName         = $Computer
					InventoryActionName  = $($RequestID.$ScheduleID)
					InventoryActionID    = $($ScheduleID)
					LastCycleStartedDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($Status.LastCycleStartedDate)
					LastReportDate       = [System.Management.ManagementDateTimeConverter]::ToDateTime($Status.LastReportDate)
				}

				
			}
			Write-Output $output
		}
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



} # END SCRIPTBLOCK

				if ($MultiThread)
				{
					$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
					$PowershellThread.AddParameter("Computer", $Computer) | Out-Null
					$PowershellThread.AddParameter("RequestAction", $RequestAction) | Out-Null
					$PowershellThread.AddParameter("Full", $Full) | Out-Null

					if($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
					{
						$PowershellThread.AddParameter("Verbose") | out-null
					}

					$PowershellThread.RunspacePool = $RunspacePool
					$Handle     = $PowershellThread.BeginInvoke()
					$Job        = "" | Select-Object Handle, Thread, object
					$Job.Handle = $Handle
					$Job.Thread = $PowershellThread
					$Job.Object = $Computer.ToString()
					$Jobs      += $Job
				}

				else
				{
					if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
					{
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$RequestAction,$Full,$Verbose
					}

					else
					{
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$RequestAction,$Full
					}
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
	end
	{
		if($MultiThread)
		{
			$ResultTimer = Get-Date
			While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0)
			{
			
				$Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
				If ($Remaining.Length -gt 60)
				{
					$Remaining = $Remaining.Substring(0,60) + "..."
				}
				Write-Progress `
					-Activity "Waiting for Jobs - $($MaxThreads - $($RunspacePool.GetAvailableRunspaces())) of $MaxThreads threads running" `
					-PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
					-Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $remaining" 

				ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True}))
				{
					$Job.Thread.EndInvoke($Job.Handle)
					$Job.Thread.Dispose()
					$Job.Thread = $Null
					$Job.Handle = $Null
					$ResultTimer = Get-Date
				}
				If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime)
				{
					Write-Warning "Child script appears to be frozen, try increasing MaxResultTime...CTRL + C to abort operation"
				}
				Start-Sleep -Milliseconds $SleepTimer
				
			} # end while

		$RunspacePool.Close() | Out-Null
		$RunspacePool.Dispose() | Out-Null
	} # end if multithread
	} # end endblock

} # end function

#Invoke-SCCMClientAction -ClientName C1204000 -RequestAction RequestMachineAssignments,RequestEvaluateMachinePolicies,HardwareInventory -Full -Verbose