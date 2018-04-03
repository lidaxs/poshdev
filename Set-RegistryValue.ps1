<#
	version 1.0.1
	changed verb to Set
	
	version 1.0.0
	Initial upload
#>

Function Set-RegistryValue {
	<#
		.SYNOPSIS
			Adds a registry key remote or local.

		.DESCRIPTION
			Adds a registry key remote or local.

		.PARAMETER  ComputerName
			The ComputerName(s) on which to operate.(String[])

		.PARAMETER  Hive
			The Hive to pick ...LocalMachine.(String)
			Possible Values for this parameter are ClassesRoot,CurrentUser,LocalMachine,Users,PerformanceData,CurrentConfig,DynData
			...see possible values...use switch -ListRegistryHive

		.PARAMETER  Key
			The key to add.(String)

		.PARAMETER  ValueName
			The ValueName to add.(String)

		.PARAMETER  Value
			The Value to add or change.(String,Int16,Int32,Byte[],String[])

		.PARAMETER  Type
			The Type to add (String,DWord,Binary,MultiString,ExpandString,QWord)
			...see possible values...use switch -ListRegistryValueKind

		.EXAMPLE
			Add-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName "My ValueName" -Value "My Value" -Type String
			
		.EXAMPLE
			Get-Content C:\temp\computers.txt | Add-RegistryValue -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName DWTest -Value 123 -Type DWord
			
		.EXAMPLE
			Add-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName BinTest -Value ([Byte[]]@(12,23,34,23)) -Type Binary
			
		.EXAMPLE
			Add-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName MSTest -Value ([String[]]@("one","two","three")) -Type MultiString
			
		.EXAMPLE
			Add-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName ESTest -Value "Dit is een expanded string" -Type ExpandString
			
		.EXAMPLE
			Add-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName QTest -Value 2234437342 -Type QWord
			
		.INPUTS
			System.String,System.String[],Switch

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
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Object])]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

		[Parameter(Mandatory=$false,HelpMessage="Possible Values for this parameter are ClassesRoot,CurrentUser,LocalMachine,Users,PerformanceData,CurrentConfig and DynData")]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$Hive='LocalMachine',
		
		[Parameter(Mandatory=$false)]
		[System.String]
		$Key,
		
		[Parameter(Mandatory=$false)]
		[System.String]
		$ValueName,
		
		[Parameter(Mandatory=$false)]
		$Value,
		
		[Parameter(Mandatory=$true,HelpMessage="Possible values for this parameter are String,DWord,Binary,MultiString,ExpandString and QWord")]
		[System.String]
		$Type,
		
		[Parameter(Mandatory=$false)]
		[Switch]
		$ListRegistryValueKind,
		
		[Parameter(Mandatory=$false)]
		[Switch]
		$ListRegistryHive,

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
		if ($MultiThread)
		{
			# If you don't know what kind of regtypes you can use this switch to get a list
			# Default is String(REG_SZ)
			if($ListRegistryValueKind){
				return [enum]::GetNames([Microsoft.Win32.RegistryValueKind])
				break
			}
			elseif($ListRegistryHive){
				return [enum]::GetNames([Microsoft.Win32.RegistryHive])
				break
			}

			Write-Verbose "Creating Default Initial Session State"
			$ISS = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
			
			Write-Verbose "Creating RunspacePool in which the threads will run"
			$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $ISS, $Host)
			
			Write-Verbose "Opening RunspacePool"
			$RunspacePool.Open()
			
			Write-Verbose "Creating Jobs array which will hold each job"
			$Jobs = @()
		}
	} # end begin

	process
    {
		#
		# test pipeline input and pick the right attributes from the incoming objects
		if($ClientName.__NAMESPACE -like 'root\sms\site_*')
        {
			Write-Verbose "Object received from sccm."
			$ClientName=$ClientName.Name
		}
		elseif($ClientName.objectclass -eq 'computer')
        {
			Write-Verbose "Object received from Active Directory module."
			$ClientName=$ClientName.Name
		}
		elseif($ClientName.__NAMESPACE -like 'root\cimv2*')
        {
			Write-Verbose "Object received from WMI"
			$ClientName=$ClientName.PSComputerName
		}
		elseif($ClientName.ComputerName)
        {
			Write-Verbose "Object received from pscustom"
			$ClientName=$ClientName.ComputerName
		}
		else
        {
			Write-Verbose "No pipeline or no specified attribute from inputobject"
		}
		# end test pipeline input and pick the right attributes from the incoming objects
		#


	# Loop through collection
		ForEach($Computer in $ClientName)
        {

			# Test connectivity
			if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue)
            {
				If($PSCmdLet.ShouldProcess("$ComputerName", "Add-RegistryValue $ValueName with value $Value in hive $Hive in key $Key of type $Type."))
                {
				    Write-Verbose "Workstation $Computer is online..."
                        $ScriptBlock=
                        {
                            param($Computer,
							$Hive='LocalMachine',
							$Key,
							$ValueName,
							$Value,
							$Type)
                        try
                        {
                            # Connect to remoteregistry servive through WMI
                            Write-Verbose "Connecting to RemoteRegistry service on $Computer!"
                            $svc=Get-WmiObject -Class Win32_Service -Filter "Name='RemoteRegistry'" -ComputerName $Computer -ErrorAction Stop
                        }
                        catch [System.Runtime.InteropServices.COMException]
                        {
                            Write-Warning "Cannot connect to $Computer through WMI"
                            Write-Warning $Error[0].Exception.Message
                        }

                        try
                        {
                            # Start remoteregistry service if stopped
                            if($svc.State -eq 'Stopped')
                            {
                                Write-Verbose "Starting service $($svc.Name) on $Computer"
                                $result=$svc.StartService()
                                if($result.Returnvalue -ne 0)
                                {
                                    Write-Warning "Cannot start RemoteRegistry service on $Computer."
                                }
                            }
                        }
                        catch [System.Management.Automation.RuntimeException]
                        {
                            Write-Warning "Cannot start RemoteRegistry service on $Computer."
                            Write-Warning $Error[0].Exception.Message
                        }

                        try
                        {
                            Write-Verbose "Opening remote base key $Hive on $Computer."
                            $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$Hive,$Computer,[Microsoft.Win32.RegistryView]::Registry64)
                        }
                        catch [System.Management.Automation.MethodInvocationException]
                        {
                            Write-Warning "Cannot open remote basekey $Hive on $Computer"
                            Write-Warning $Error[0].Exception.Message
                        }


                        try
                        {
                            Write-Verbose "Creating key $key on $Computer in $Hive Hive."
                            [void]$RemoteRegistry.CreateSubKey($Key)
                        }
                        catch
                        {
                            Write-Warning $Error[0].Exception.Message
                        }
                        try
                        {
                            Write-Verbose "Opening key $Key for writing."
                            $CreatedKey=$RemoteRegistry.OpenSubKey($Key,$True)
                        }
                        catch
                        {
                            Write-Warning $Error[0].Exception.Message
                        }
                        try
                        {
                            Write-Verbose "Creating valuename $valuename with value $value on $Computer in $CreatedKey."
                            $CreatedKey.SetValue($ValueName,$Value,[Microsoft.Win32.RegistryValueKind]::$Type)
                        }
                        catch
                        {
                            Write-Warning "Cannot create key $Key and\or set value $Value on $Computer in $Hive Hive of type $Type."
                            Write-Warning $Error[0].Exception.Message
                        }
                    }
				} # end if $PSCmdlet.ShouldProcess

				if ($MultiThread) {
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
				$PowershellThread.AddParameter("Computer", $Computer) | Out-Null
				$PowershellThread.AddParameter("Hive", $Hive) | Out-Null
				$PowershellThread.AddParameter("Key", $Key) | Out-Null
				$PowershellThread.AddParameter("ValueName", $ValueName) | Out-Null
				$PowershellThread.AddParameter("Value", $Value) | Out-Null
				$PowershellThread.AddParameter("Type", $Type) | Out-Null
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
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Hive,$Key,$ValueName,$Value,$Type
				}
			} # end if test-connection

		
            else
            {
                Write-Warning "$Computer is not online!"
            }
		} # end foreach $Computer

	} # end process
	end
    {
		if ($MultiThread)
			{
			$ResultTimer = Get-Date
			
			While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0)  {
			
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
					Write-Error "Child script appears to be frozen, try increasing MaxResultTime"
					Exit
				}
				Start-Sleep -Milliseconds $SleepTimer
			} 
			$RunspacePool.Close() | Out-Null
			$RunspacePool.Dispose() | Out-Null	
		}
	}
}