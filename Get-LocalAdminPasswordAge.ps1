<#
	version 1.0.0.1
	test input

	version 1.0.0.0
	initial upload
#>
function Get-LocalAdminPasswordAge {
	<#
		.SYNOPSIS
			Short description of function.

		.DESCRIPTION
			long description of function.

		.PARAMETER  ClientName
			The ClientName(s) on which to operate.
			This can be a string or collection

		.PARAMETER MultiThread
			Enable multithreading

		.PARAMETER MaxThreads
			Maximum number of threads to run simultaneously

		.PARAMETER MaxResultTime
			Max time in which a thread must finish(seconds)

		.PARAMETER SleepTimer
			Time to wait between checks if thread has finished

		.EXAMPLE
			PS C:\> Get-LocalAdminPasswordAge -ClientName C120VMXP,C120WIN7

		.EXAMPLE
			PS C:\> $mycollection | Get-LocalAdminPasswordAge

		.INPUTS
			System.String,System.String[]

		.OUTPUTS
			System.Object[]

		.NOTES
			Additional information about the function go here.

		.LINK
			about_functions_advanced

		.LINK
			about_comment_based_help
	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Object])]
	param(
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[Alias("Name","CN","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
        $ClientName=@($env:COMPUTERNAME),
        
        [String]
        $LocalAccountName = "Administrator",

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
	
	begin
	{
		if($ClientName.Name){$ClientName=$ClientName.Name}
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
	}
	process
	{
		# loop through collection
		ForEach($Computer in $ClientName)
		{
			if($Computer.Name){$Computer=$Computer.Name}
			if($PSCmdlet.ShouldProcess("$Computer", "Get-LocalAdminPasswordAge"))
			{

				$ScriptBlock=
				{[CmdletBinding(SupportsShouldProcess=$true)]
				param
				(
					[String]
                    $Computer,
                    
                    [String]
                    $LocalAccountName
				)

					try
					{
						$adc = [adsisearcher]"CN=$($Computer)"
						$adc.PropertiesToLoad.AddRange(@("lastlogontimestamp"))
						$adc.SearchRoot = [ADSI]"LDAP://OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local"
						$lastad = $adc.FindOne()

						$output = New-Object -TypeName PSObject -Property @{
							ComputerName = $Computer
							Account = $LocalAccountName
							PasswordAge = 'Unknown'
							LastLogonAD =  [datetime]::FromFileTime($lastad.Properties["lastlogontimestamp"][0])
						}

					}
					catch
					{
						Write-Verbose "$Computer not found in Active Directory under OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local"
					}

					# Test connectivity
					if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0) 
					{
						#the code to execute in each thread
						try
						{
                            $output.PasswordAge =([System.Math]::Round([int]([adsi]"WinNT://$Computer/$LocalAccountName,user").passwordage.value/86400)).ToString()
						}
						catch
						{
						}
						
						
						
					} # end if test-connection
					
					else # computer is online
					{
						Write-Warning "$Computer is not online!"
                    }
                    $output | Select-Object ComputerName,PasswordAge,LastLogonAD,Account
				} # end scriptblock


			} # end if $PSCmdlet.ShouldProcess


			if ($MultiThread)
			{
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
                $PowershellThread.AddParameter("Computer", $Computer) | out-null
                $PowershellThread.AddParameter("LocalAccountName", $LocalAccountName) | out-null
				if($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
				{
					$PowershellThread.AddParameter("Verbose") | out-null
				}
				$PowershellThread.RunspacePool = $RunspacePool

				$Handle = $PowershellThread.BeginInvoke()
				$Job = "" | Select-Object Handle, Thread, object
				$Job.Handle = $Handle
				$Job.Thread = $PowershellThread
				$Job.Object = $Computer.ToString()
				$Jobs += $Job				
			}
			else # $MultiThread
			{
				if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$LocalAccountName,$Verbose
				}
				# for each parameter in the scriptblock add the same argument to the argumentlist
				else
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$LocalAccountName
				}
			}

		} # end foreach $computer

	} # end processblock

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