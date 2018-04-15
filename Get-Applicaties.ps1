<#
	version 1.0.0.1
	Added aliases to parameter Clientname to support pipelineinput from AD,SCCM and WMI
	changed test for connectivity
	moved datetimetemplate to scriptblock

	version 1.0.0.0
	Initial upload
#>
Function Get-Applicaties {
	<#
		.SYNOPSIS
			Returns list of installed applications(32- and 64-bits) filtered by DisplayName,Version and or Publisher.

		.DESCRIPTION
			Returns list of installed applications(32- and 64-bits) filtered by DisplayName,Version and or Publisher.

		.PARAMETER ClientName
			The ComputerName(s) on which to operate.(Accepts value from pipeline)

		.PARAMETER DisplayName
			The Displayname to filter

		.PARAMETER Publisher
			The publisher to filter

		.PARAMETER Version
			The version to filter 

		.PARAMETER MultiThread
			Enable multithreading

		.PARAMETER MaxThreads
			Maximum number of threads to run simultaneously

		.PARAMETER MaxResultTime
			Max time in which a thread must finish(seconds)

		.PARAMETER SleepTimer
			Time to wait between checks if thread has finished

		.EXAMPLE
			Get-Applicaties -ClientName 'C120VMXP','C120WIN7' -DisplayName "*.NET*"

		.EXAMPLE
			Get-Applicaties -ClientName 'C120VMXP','C120WIN7' -Version "5.2"

		.EXAMPLE
			Get-Applicaties -ClientName 'C120VMXP','C120WIN7' -Publisher "SRWare"

		.EXAMPLE
			'C120VMXP','C120WIN7' | Get-Applicaties

		.EXAMPLE
			Get-Applicaties -ClientName (Get-Content C:\computers.txt) -MultiThread

		.EXAMPLE
			(Get-Content C:\computers.txt) | Get-Applicaties

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
		[Parameter(Position=0, Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[Alias("PSComputerName","ComputerName","CN","MachineName","Workstation","ServerName","HostName","Name")]
		$ClientName=@($env:COMPUTERNAME),

	    [Parameter(Mandatory=$false)]
		[Alias("Application","Program","Software")]
	    [string]$DisplayName="*",

	    [Parameter(Mandatory=$false)]
	    [string]$Publisher="*",

	    [Parameter(Mandatory=$false)]
	    [string]$Version="*",
		
		[Switch]
		$MultiThread,
		
		$MaxThreads=20,
		
		$MaxResultTime=20,
		
		$SleepTimer=1000
	)

	# set initial values in the begin block (populate variables, check dependent modules etc.)
	begin {
		# Hive of the Registry
		$hive='LocalMachine'
		Write-Verbose "Setting Hive to '$hive'"

		# DateTimeTemplate for parsing installdate
		$datetimetemplate='yyyyMMdd'
		Write-Verbose "Setting datetimetemplate to '$datetimetemplate' for parsing purposes of the installdate"

		# returning parameter values
		Write-Verbose "DisplayName is '$($DisplayName)'"
		Write-Verbose "Publisher is '$($Publisher)'"
		Write-Verbose "Version is '$($Version)'"
		
		if($MultiThread)
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
	process
	{

		# add -Whatif and -Confirm support to the CmdLet
		if($PSCmdlet.ShouldProcess("$ClientName", "Get-Applicaties"))
		{
			# loop through collection $ClientName
			ForEach($Computer in $ClientName)
			{

				$ScriptBlock=
				{
					param(
					$Computer,
					$DisplayName="*",
					$Publisher="*",
					$Version="*"
					)

					$datetimetemplate='yyyyMMdd'
					
					$Hive='LocalMachine'
	
						# test connection to each $Computer
						if([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,139).AsyncWaitHandle.WaitOne(1000,$false))
						{
							Write-Verbose "$Computer is online..."
							# start try
							try
							{
								$svc=Get-WmiObject -Class Win32_Service -Filter "Name='RemoteRegistry'" -ComputerName $Computer -ErrorAction Stop
								if($svc.StartMode -eq 'Disabled')
								{
									Write-Verbose "Enabling RemoteRegistry service on $Computer....."
									if($svc.ChangeStartMode('Automatic').ReturnValue -ne 0)
									{
										Write-Warning "Error enabling RemoteRegistry service on $Computer!"
									}
									else
									{
										Write-Verbose "Succesfully enabled RemoteRegistry service on $Computer."
									}
								}
								
								elseif($svc.State -eq 'Stopped')
								{
									Write-Verbose "Starting RemoteRegistry service on $Computer....."	
									if($svc.StartService().ReturnValue -ne 0)
										{
											Write-Warning " ## Error starting service RemoteRegistry on $Computer! ##"
											continue
										}
									else
									{
										Write-Verbose "Succesfully started service RemoteRegistry on $Computer"
									}
								}
							} # end try
							# start catch specific
							catch [System.Runtime.InteropServices.COMException]
							{
								Write-Warning "Cannot connect to $Computer through WMI"
									$Error[0].Exception.Message
							} # end catch specific error
							# catch rest of errors
							catch
							{
								$Error[0].Exception.Message
							} # end catch rest of errors


							# do the things you want to do after this line

							# let's get some operatingsysteminfo
							$client_os=Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer


							# let's get some computersysteminfo
							$client_system=Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
							if($client_system.DomainRole -eq 0){$DomainRole="Standalone Workstation"}
							if($client_system.DomainRole -eq 1){$DomainRole="Member Workstation"}
							if($client_system.DomainRole -eq 2){$DomainRole="Standalone Server"}
							if($client_system.DomainRole -eq 3){$DomainRole="Member Server"}
							if($client_system.DomainRole -eq 4){$DomainRole="Backup Domain Controller"}
							if($client_system.DomainRole -eq 5){$DomainRole="Primary Domain Controller"}
							
							# connect remote registry
							try
							{
								Write-Verbose "Using the registry64 view on $Computer."
								$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$Hive,$Computer,[Microsoft.Win32.RegistryView]::Registry64)
								if($RemoteRegistry)
								{
									Write-Verbose "64-bit registry view loaded."
								}
								else
								{
									Write-Verbose "Using the registry32 view on $ClientName."
									$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive,$Computer)
								}
							}
							catch [System.Management.Automation.MethodInvocationException]
							{
									Write-Warning "Cannot open Registry on $Computer"
									Write-Warning $Error[0].Exception.Message
							}
							catch
							{
									Write-Warning "Some other error occurred opening the registry of computer $Computer"
									Write-Warning $Error[0].Exception.ErrorRecord
							}# end connect remote registry

							# create result variable
							$result=New-Object -TypeName System.Collections.ArrayList

							# set uninstallkey for 32-bit apps
							$uninstallkey="SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

							try
							{
								$uninstall_subkeys=$RemoteRegistry.OpenSubKey($uninstallkey)
								$programs=$uninstall_subkeys.GetSubKeyNames()

								foreach ($program in $programs)
								{
									$output=New-Object -TypeName PSObject | Select-Object ComputerName,KeyName,DisplayName,Version,UninstallString,InstallDate,InstallLocation,InstallSource,Language,Publisher,LongKeyName,OperatingSystemName,OperatingSystemServicePack,SystemType,DomainRole
									
									# mind the escape '\' character when using doublequotes
									$programkey=$uninstallkey + '\' + $program
									
									$programvalues=$RemoteRegistry.OpenSubKey($programkey)
									
									$output.ComputerName=$Computer
									$output.KeyName=$program
									$output.LongKeyName=$programkey
									$output.DisplayName=$programvalues.GetValue("DisplayName")
									$output.Version=$programvalues.GetValue("DisplayVersion")
									$output.UninstallString=$programvalues.GetValue("UninstallString")
									$output.InstallLocation=$programvalues.GetValue("InstallLocation")
									$output.InstallSource=$programvalues.GetValue("InstallSource")
									try
									{
										$output.InstallDate=Get-Date -Date ([DateTime]::ParseExact($programvalues.GetValue("InstallDate"), $datetimetemplate, $null)) -Format "yyyy-MM-dd"
									}
									catch
									{
										$output.InstallDate="Unknown"
									}
									
									$output.Language=$programvalues.GetValue("Language")
									$output.Publisher=$programvalues.GetValue("Publisher")
									$output.OperatingSystemName=$client_os.Caption
									$output.OperatingSystemServicePack=$client_os.CSDVersion
									$output.SystemType=$client_system.SystemType
									$output.DomainRole=$DomainRole
							
									[void]$result.Add($output)

								}
						
							}	
							catch
							{
								Write-Verbose "$uninstallkey cannot be opened on $Computer"
							}
							# set uninstallkey for 32-bit\64-bit apps
							$uninstallkey="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

						try
						{
							$uninstall_subkeys=$RemoteRegistry.OpenSubKey($uninstallkey)
							$programs=$uninstall_subkeys.GetSubKeyNames()

							foreach ($program in $programs)
							{
							
								$output=New-Object -TypeName PSObject | Select-Object ComputerName,KeyName,DisplayName,Version,UninstallString,InstallDate,InstallLocation,InstallSource,Language,Publisher,LongKeyName,OperatingSystemName,OperatingSystemServicePack,SystemType,DomainRole
									
								# mind the escape '\' character when using doublequotes
								$programkey=$uninstallkey + '\' + $program
									
								$programvalues=$RemoteRegistry.OpenSubKey($programkey)
									
								$output.ComputerName=$Computer
								$output.KeyName=$program
								$output.LongKeyName=$programkey
								$output.DisplayName=$programvalues.GetValue("DisplayName")
								$output.Version=$programvalues.GetValue("DisplayVersion")
								$output.UninstallString=$programvalues.GetValue("UninstallString")
								$output.InstallLocation=$programvalues.GetValue("InstallLocation")
								$output.InstallSource=$programvalues.GetValue("InstallSource")
								try
								{
									$output.InstallDate=Get-Date -Date ([DateTime]::ParseExact($programvalues.GetValue("InstallDate"), $datetimetemplate, $null)) -Format "yyyy-MM-dd"
								}
								catch
								{
									$output.InstallDate="Unknown"
								}
									
								$output.Language=$programvalues.GetValue("Language")
								$output.Publisher=$programvalues.GetValue("Publisher")
								$output.OperatingSystemName=$client_os.Caption
								$output.OperatingSystemServicePack=$client_os.CSDVersion
								$output.SystemType=$client_system.SystemType
								$output.DomainRole=$DomainRole
							
								[void]$result.Add($output)

							} # end foreach program

						} # end try	
						catch
						{
							Write-Verbose "$uninstallkey cannot be opened on $Computer"
						}

						# output results
						$result | Where-Object {(($_.DisplayName -like $DisplayName) -and ($_.DisplayName -ne $null) -and ($_.DisplayName -ne '')) -and ($_.Publisher -like $Publisher) -and ($_.Version -like $Version)}
						

					} # if test-connection 
							
					else
					{
						Write-Warning "$Computer is not online!"
					} # end test connection to each $Computer
				} # end scriptblock

				if($MultiThread)
				{
					$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
					$PowershellThread.AddParameter("Computer", $Computer) | out-null
					$PowershellThread.AddParameter("DisplayName", $DisplayName) | out-null
					$PowershellThread.AddParameter("Publisher", $Publisher) | out-null
					$PowershellThread.AddParameter("Version", $Version) | out-null
					$PowershellThread.RunspacePool = $RunspacePool
					$Handle = $PowershellThread.BeginInvoke()
					$Job = "" | Select-Object Handle, Thread, object
					$Job.Handle = $Handle
					$Job.Thread = $PowershellThread
					$Job.Object = $Computer.ToString()
					$Jobs += $Job
				}
				else
				{
					if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
					{
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$DisplayName,$Publisher,$Version,$Verbose
					}
					# for each parameter in the scriptblock add the same argument to the argumentlist
					else
					{
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$DisplayName,$Publisher,$Version
					}
				}

			} # end for each $Computer
				
		} # end if $pscmdlet.ShouldProcess

	} # end processblock


	# remove variables  in the endblock
	end
	{
		if($MultiThread)
		{
		    $ResultTimer = Get-Date
			
		    While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0)  {
		    
		        $Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
		        If ($Remaining.Length -gt 60){
		            $Remaining = $Remaining.Substring(0,60) + "..."
		        }
		        Write-Progress `
		            -Activity "Waiting for Jobs - $($MaxThreads - $($RunspacePool.GetAvailableRunspaces())) of $MaxThreads threads running" `
		            -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
		            -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $remaining" 

		        ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True})){
				    $Job.Thread.EndInvoke($Job.Handle)
		            $Job.Thread.Dispose()
		            $Job.Thread = $Null
		            $Job.Handle = $Null
		            $ResultTimer = Get-Date
		        }
		        If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime){
                    Write-Warning "Child script appears to be frozen, try increasing MaxResultTime...CTRL + C to abort operation"
		        }
		        Start-Sleep -Milliseconds $SleepTimer
		        
		    }
			
		    $RunspacePool.Close() | Out-Null
		    $RunspacePool.Dispose() | Out-Null
		}
	} # end endblock
} # end function