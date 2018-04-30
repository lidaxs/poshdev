<#
	version 1.0.5.4
	updated examples
	
	version 1.0.5.3
	aliases not working as expected when using pipeline and piping different types of objects
	added if($Computer.Name){$Computer=$Computer.Name} in processblock

	version 1.0.5.2
	test connection with wmi
	
	version 1.0.5.1
	set default value for dynamic parameter Hive in beginblock and set mandatory to $false
	
	version 1.0.5
	fixed invoke-command(replaced variables $Hive and $Type with $($PSBoundParameters.Hive) and $($PSBoundParameters.Type))

	version 1.0.4
	added dynamic parameters for  Hive and Type
	added verbosing to multithreading
	removed obsolete parameters ListRegistryHive and ListValueKind
	removed parameters Hive and Type...they are now in the DynamicParam block

	version 1.0.3
	added alias Add-RegistryValue	

	version 1.0.2
	Added Aliases to ClientName parameter to support pipeline in from WMI,SCCM & Active Directory

	version 1.0.1
	changed verb to Set
	
	version 1.0.0
	Initial upload

	wishlist:
	dynamic parameter for Hive..done
	dynamic parameter for Type..done
	verbose in multithread..done
	removal of list switches..done

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
			The Value to add/set or change.(String,Int16,Int32,Byte[],String[])

		.PARAMETER  Type
			The Type to set or add (String,DWord,Binary,MultiString,ExpandString,QWord)
			...see possible values...use switch -ListRegistryValueKind

		.EXAMPLE
			Set-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName "My ValueName" -Value "My Value" -Type String
			
		.EXAMPLE
			Get-Content C:\temp\computers.txt | Set-RegistryValue -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName DWTest -Value 123 -Type DWord
			
		.EXAMPLE
			Set-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName BinTest -Value ([Byte[]]@(12,23,34,23)) -Type Binary
			
		.EXAMPLE
			Set-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName MSTest -Value ([String[]]@("one","two","three")) -Type MultiString
			
		.EXAMPLE
			Set-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName ESTest -Value "Dit is een expanded string" -Type ExpandString
			
		.EXAMPLE
			Set-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName QTest -Value 2234437342 -Type QWord
			
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
	[Alias('Add-RegistryValue')]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

		[Parameter(Mandatory=$false)]
		[System.String]
		$Key,
		
		[Parameter(Mandatory=$false)]
		[System.String]
		$ValueName,
		
		[Parameter(Mandatory=$false)]
		$Value,
		
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
    DynamicParam
    {          
        $HiveParameterAttributes                   = New-Object System.Management.Automation.ParameterAttribute
        $HiveParameterAttributes.Mandatory         = $false
        $HiveParameterAttributes.HelpMessage       = "Press `'TAB`' to cycle through the different values"
		$HiveParameterAttributes.ParameterSetName  = '__AllParameterSets'
		$HiveAttributeCollection                   = New-Object  System.Collections.ObjectModel.Collection[System.Attribute]
		$HiveAttributeCollection.Add($HiveParameterAttributes)
		$Hive                                      = [enum]::GetNames([Microsoft.Win32.RegistryHive])
		$HiveAttributeCollection.Add((New-Object  System.Management.Automation.ValidateSetAttribute($Hive)))
		$HiveRuntimeParameters                     = New-Object System.Management.Automation.RuntimeDefinedParameter('Hive', [System.String[]], $HiveAttributeCollection)

        $TypeParameterAttributes                   = New-Object System.Management.Automation.ParameterAttribute
        $TypeParameterAttributes.Mandatory         = $false
        $TypeParameterAttributes.HelpMessage       = "Press `'TAB`' to cycle through the different typevalues"
		$TypeParameterAttributes.ParameterSetName  = '__AllParameterSets'
		$TypeAttributeCollection                   = New-Object  System.Collections.ObjectModel.Collection[System.Attribute]
		$TypeAttributeCollection.Add($TypeParameterAttributes)
		$Type                                      = [enum]::GetNames([Microsoft.Win32.RegistryValueKind])
		$TypeAttributeCollection.Add((New-Object  System.Management.Automation.ValidateSetAttribute($Type)))
		$TypeRuntimeParameters                     = New-Object System.Management.Automation.RuntimeDefinedParameter('Type', [System.String[]], $TypeAttributeCollection)

		$RuntimeParametersDictionary               = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
		$RuntimeParametersDictionary.Add('Hive', $HiveRuntimeParameters)
		$RuntimeParametersDictionary.Add('Type', $TypeRuntimeParameters)
        return  $RuntimeParametersDictionary
	}
	begin
	{    
		if ( -not ($PSBoundParameters.Hive))
		{
        	$PSBoundParameters.Hive = 'LocalMachine'
		}
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
	} # end begin

	process
    {

		# Loop through collection
		ForEach($Computer in $ClientName)
        {

			if($Computer.Name){$Computer=$Computer.Name}
			
			# Test connectivity
			if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0) 
            {
				If($PSCmdLet.ShouldProcess("$Computer", "Add-RegistryValue $ValueName with value $Value in hive $Hive in key $Key of type $Type."))
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

				if ($MultiThread)
				{
					$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
					$PowershellThread.AddParameter("Computer", $Computer) | Out-Null
					$PowershellThread.AddParameter("Hive", $($PSBoundParameters.Hive)) | Out-Null
					$PowershellThread.AddParameter("Key", $Key) | Out-Null
					$PowershellThread.AddParameter("ValueName", $ValueName) | Out-Null
					$PowershellThread.AddParameter("Value", $Value) | Out-Null
					$PowershellThread.AddParameter("Type", $($PSBoundParameters.Type)) | Out-Null

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
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$($PSBoundParameters.Hive),$Key,$ValueName,$Value,$($PSBoundParameters.Type),$Verbose
					}

					else
					{
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$($PSBoundParameters.Hive),$Key,$ValueName,$Value,$($PSBoundParameters.Type)
					}
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
					$Job.Thread  = $Null
					$Job.Handle  = $Null
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