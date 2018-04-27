<#
    version 1.0.0.0
    initial upload
#>
#region functions
function Is64BitOperatingSystem
{
<#
    .Synopsis
        Return true if Operatingsystem is 64-bit
    .DESCRIPTION
        Return true if Operatingsystem is 64-bit
    .EXAMPLE
        Is64BitOperatingSystem
    .EXAMPLE
        Is64BitOperatingSystem -ComputerName MyTestMachine,MyOtherServer
#>

    [CmdletBinding()]
    [Alias()]
    [OutputType([Boolean])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        $ComputerName=$env:COMPUTERNAME
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        foreach ($Computer in $ComputerName)
        {
            if ((Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer).SystemType -eq 'x64-based PC')
            {
                Write-EventLog -LogName Application -EntryType Information -EventId 1 -Source $regsource -Message "Operating system is 64 bits"
                Write-Verbose "64-bits OS"
                $true
            }
            else
            {
                Write-EventLog -LogName Application -EntryType Information -EventId 1 -Source $regsource -Message "Operating system is 32 bits"
                Write-Verbose "32-bits OS"
                $false
            }
        }
    }
    End
    {
    }
}

function Is64BitProcess
{
    return [IntPtr]::Size -eq 8
}

Function Add-RegistryValue
{
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
        [ValidateSet('ClassesRoot','CurrentUser','LocalMachine','Users','PerformanceData','CurrentConfig','DynData')]
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
        [ValidateSet('String','DWord','Binary','MultiString','ExpandString','QWord')]
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

        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue

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
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Creating key `'$Key`' in Hive: $Hive"
                        }
                        catch
                        {
                            $eMessage=$Error[0].Exception.Message
                            Write-Warning $eMessage
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Creating key `'$Key`' in Hive: $Hive failed....error $eMessage"
                        }
                        try
                        {
                            Write-Verbose "Opening key $Key for writing."
                            $CreatedKey=$RemoteRegistry.OpenSubKey($Key,$True)
                        }
                        catch
                        {
                            $eMessage=$Error[0].Exception.Message
                            Write-Warning $eMessage
                        }
                        try
                        {
                            Write-Verbose "Creating valuename `'$valuename`' with value `'$value`' on $Computer in `'$CreatedKey`' of type : $Type"
                            $CreatedKey.SetValue($ValueName,$Value,[Microsoft.Win32.RegistryValueKind]::$Type)
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Creating valuename `'$valuename`' with value `'$value`' on $Computer in `'$CreatedKey`' of type : $Type"
                        }
                        catch
                        {
                            $eMessage=$Error[0].Exception.Message
                            Write-Warning "Cannot create key `'$Key`' and\or set value `'$Value`' on $Computer in $Hive Hive of type $Type."
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Creating valuename `'$valuename`' with value `'$value`' on $Computer in `'$CreatedKey`' of type : $Type failed :-(....$eMessage"
                            Write-Warning $eMessage
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
				else {
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

Function Get-RegistryValue
{
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
		[Parameter(Position=0, Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","ServerName","Workstation","HostName")]
		$ComputerName=@($env:COMPUTERNAME),

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
		ForEach($Computer in $ComputerName){
		
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

Function Remove-RegistryKey
{
	<#
		.SYNOPSIS
			Removes a registry key remote or local.(No subkeys)

		.DESCRIPTION
			Removes a registry key remote or local.(No subkeys)

		.PARAMETER  ComputerName
			The ComputerName(s) on which to operate.

		.PARAMETER  Hive
			The Hive to pick ...LocalMachine.
			...see possible values...use switch -ListRegistryHive

		.PARAMETER  Key
			The key to remove.

		.PARAMETER IncludeSubKeys
			Will also remove the subkeys

		.EXAMPLE
			Remove-RegistryKey -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey"
			
		.INPUTS
			System.String,System.String

		.OUTPUTS
			System.String

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
		[Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","MachineName","Workstation","ServerName","HostName")]
		[ValidateNotNullOrEmpty()]
		$ComputerName=@($env:COMPUTERNAME),

		[Parameter(Position=1, Mandatory=$false)]
        [ValidateSet('ClassesRoot','CurrentUser','LocalMachine','Users','PerformanceData','CurrentConfig','DynData')]
		[System.String]
		$Hive='LocalMachine',

		[Parameter(Position=2, Mandatory=$false)]
		[System.String]
		$Key,

		[Parameter(Position=3, Mandatory=$false)]
		[Switch]
		$IncludeSubKeys,

		[Parameter(Position=4, Mandatory=$false)]
		[Switch]
		$ListRegistryHive
	)

	begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
	} # end begin

	process
    {
	
		# If you don't know what kind of regtypes you can use this switch to get a list
		# Default is String(REG_SZ)
		if($ListRegistryHive){
		return [enum]::GetNames([Microsoft.Win32.RegistryHive])
		break
		}

	# Loop through collection
		ForEach($Computer in $ComputerName){

			# Test connectivity
			if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue){

				Write-Verbose "Workstation $Computer is online..."
				
				try {
					# Connect to remoteregistry servive through WMI
					Write-Verbose "Connecting to RemoteRegistry service on $Computer!"
					$svc=Get-WmiObject -Class Win32_Service -Filter "Name='RemoteRegistry'" -ComputerName $Computer -ErrorAction Stop
				}
				catch [System.Runtime.InteropServices.COMException] {
					Write-Warning "Cannot connect to $Computer through WMI"
					Write-Warning $Error[0].Exception.Message
				}

				try{
					# Start remoteregistry service if stopped
					if($svc.State -eq 'Stopped'){
						Write-Verbose "Starting service $($svc.Name) on $Computer"
						$result=$svc.StartService()
						if($result.Returnvalue -ne 0){
							Write-Warning "Cannot start RemoteRegistry service on $Computer."
						}
					}
				}
				catch [System.Management.Automation.RuntimeException]{
						Write-Warning "Cannot start RemoteRegistry service on $Computer."
						Write-Warning $Error[0].Exception.Message
					}

				try{
					Write-Verbose "Opening remote base key $Hive on $Computer."
					$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$Hive,$Computer,[Microsoft.Win32.RegistryView]::Registry64)
				}
				catch [System.Management.Automation.MethodInvocationException]{
					Write-Warning "Cannot open remote basekey $Hive on $Computer"
					Write-Warning $Error[0].Exception.Message
				}
				If($PSCmdLet.ShouldProcess("$ComputerName", "Remove-RegistryKey $Key in hive $Hive.")){
					$Leaf=Split-Path $Key -Leaf
					$Parent=Split-Path $Key
					if($IncludeSubKeys){
						try{
							Write-Verbose "Deleting key $Key including subkeys on $Computer in $Hive."
							$objKey=$RemoteRegistry.OpenSubKey($Parent,$True)
							$objKey.DeleteSubKeyTree($Leaf)
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Removing key `'$Key`' including subkeys in Hive: $Hive"
						}
						catch {
							Write-Warning "Cannot delete key $Leaf including subkeys in $Key on $Computer in Hive $Hive."
							$eMessage = $Error[0].Exception.Message
                            Write-Warning $eMessage
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Removing key `'$Key`' including subkeys in Hive: $Hive failed...$eMessage"
						}
					}
					else{
						try{
							Write-Verbose "Deleting key $Key on $Computer in $Hive."
							$objKey=$RemoteRegistry.OpenSubKey($Parent,$True)
							$objKey.DeleteSubKey($Leaf)
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Removing key `'$Key`' in Hive: $Hive"
						}
						catch {
							Write-Warning "Cannot delete key $Leaf in $Key on $Computer in Hive $Hive."
							$eMessage = $Error[0].Exception.Message
                            Write-Warning $eMessage
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Removing key `'$Key`' in Hive: $Hive failed...$eMessage"

						}
					}
				} # end if $PSCmdlet.ShouldProcess
			} # end if test-connection

		
		else{
				Write-Warning "$Computer is not online!"
			}
		} # end foreach $Computer

	} # end process

	end	{
		
	}
}

Function Remove-RegistryValue
{
	<#
		.SYNOPSIS
			Removes a registry value remote or local.

		.DESCRIPTION
			Removes a registry value remote or local.

		.PARAMETER  ComputerName
			The ComputerName(s) on which to operate.

		.PARAMETER  Hive
			The Hive to pick ...LocalMachine.

		.PARAMETER  Key
			The key to query.

		.EXAMPLE
			Remove-RegistryValue -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -ValueName "MyTestValue"
			
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

	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Object])]
	param(
		[Parameter(Position=0, Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","ServerName","Workstation","HostName")]
		$ComputerName=@($env:COMPUTERNAME),

		[Parameter(Position=1, Mandatory=$false)]
		[ValidateNotNullOrEmpty()]
        [ValidateSet('ClassesRoot','CurrentUser','LocalMachine','Users','PerformanceData','CurrentConfig','DynData')]
		[System.String]
		$Hive='LocalMachine',
		
		[Parameter(Position=2, Mandatory=$false)]
		[System.String]
		$Key,
		
		[Parameter(Position=3, Mandatory=$false)]
		[System.String]
		$ValueName

	)

begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
	}

process{

		# Loop through collection
		ForEach($Computer in $ComputerName){
		
			# Test connectivity
			if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue){
			
				#$note=New-Object -TypeName PSObject | Select-Object Computer,Hive,Key,ValueName,Value
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

				$SubKey=$RemoteRegistry.OpenSubKey($Key,$True)

				if($PSCmdlet.ShouldProcess("Removing valuename $ValueName in key $Key in hive $Hive on computer $Computer", "Remove Registry Value")){
					try{
						$SubKey.DeleteValue($ValueName)
						Write-Verbose "Succesfully removed valuename $ValueName in key $Key in hive $Hive on computer $Computer"
                        Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Removing valuename `'$ValueName`' in `'$Key`' in Hive: $Hive"
					}
					catch [System.Management.Automation.RuntimeException] {
						Write-Warning "Cannot delete $ValueName in $Key on $Computer"
							$eMessage = $Error[0].Exception.Message
                            Write-Warning $eMessage
                            Write-EventLog -LogName Application -EntryType Information -EventId 12 -Source $regsource -Message "Removing valuename `'$ValueName`' in `'$Key`' in Hive: $Hive failed...$eMessage"

					}
				}
			}
		else{
				Write-Warning "$Computer is not online!"
			}
		}
	}
	
end	{
		Remove-Variable RegValue -Force -ErrorAction SilentlyContinue
		Remove-Variable SubKey -Force -ErrorAction SilentlyContinue
		Remove-Variable RemoteRegistry -Force -ErrorAction SilentlyContinue
	}
}

function ADSI-MemberOf
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $GroupName
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue

        $adcomputer = ([adsisearcher]"CN=$($env:COMPUTERNAME)").FindOne().GetDirectoryEntry()

    }
    Process
    {
        if (([adsisearcher]"CN=$GroupName").FindOne().GetDirectoryEntry().member -contains $adcomputer.distinguishedName) {
            #statements
            Write-Verbose "$env:COMPUTERNAME is member of the group $GroupName"
            Write-EventLog -LogName Application -EntryType Information -EventId 5 -Source $regsource -Message "$env:COMPUTERNAME is member of $GroupName"
            return $true
        }
        else
        {
            #statements
            Write-Verbose "$env:COMPUTERNAME is NOT member of the group $GroupName"
            Write-EventLog -LogName Application -EntryType Information -EventId 5 -Source $regsource -Message "$env:COMPUTERNAME is NOT member of $GroupName"
            return $false
        }
    }
    End
    {
    }
}

function Process-Command
{
    [CmdletBinding(SupportsShouldProcess=$false)]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Commandline,

        # Param2 help description
        $Argumentlist,

        $WorkingDirectory=$env:TEMP
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        try
        {
            $arguments2string = $Argumentlist -join ","
            Write-Verbose "Starting process $Commandline with arguments $arguments2string"
            $result = Start-Process -FilePath $Commandline -ArgumentList $Argumentlist -PassThru
            Wait-Process -InputObject $result
            Write-EventLog -LogName Application -EntryType Information -EventId 10 -Source $regsource -Message "Process `"$process`" started with arguments `"$arguments2string`""
        }
        catch
        {
            Write-Warning "Starting process $Commandline with arguments $arguments2string failed!!"
            $eMessage=$Error[0].Exception.Message
            Write-EventLog -LogName Application -EntryType Warning -EventId 10 -Source $regsource -Message "Processing `"$Commandline`" with arguments `"$arguments2string`" failed!...$eMessage"
        }
    
        Write-Verbose "Process $Commandline exited with exitcode $($result.ExitCode)"
        Write-EventLog -LogName Application -EntryType Information -EventId 10 -Source $regsource -Message "Processing `'$Commandline`' with arguments returned exitcode $($result.ExitCode)"

        return $result.ExitCode
    }
    End
    {
    }
}

function Process-Service
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $ServiceName,

        [Switch]
        $Stop,

        [Switch]
        $Start,

        [Switch]
        $Restart,

        [Switch]
        $SetAutomatic,

        [Switch]
        $SetDisabled,

        [Switch]
        $SetManual
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        foreach ($service in $ServiceName)
        {
            if($Stop)
            {
                try
                {
                    $message = "Stopping service $service"
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message $message
                    Stop-Service -Name $service -Force
                }
                catch
                {
                    Write-Verbose "$message failed!"
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message "$message failed!"
                }
            }
            if($Start)
            {
                try
                {
                    $message = "Starting service $service"
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message $message
                    Start-Service -Name $service -Force -Verbose
                }
                catch
                {
                    Write-Verbose "$message failed!"
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message "$message failed!"
                }
            }
            if($Restart)
            {
                try
                {
                    $message = "Restarting service $service"
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message $message
                    Restart-Service -Name $service -Force
                }
                catch
                {
                    Write-Verbose "$message failed!"
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message "$message failed!"
                }
            }
            if($SetAutomatic)
            {
                try
                {
                    $message = "Setting service $service : Automatic"
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message $message
                    Set-Service -Name $service -StartupType Automatic -Force
                }
                catch
                {
                    Write-Verbose "$message failed!"
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message "$message failed!"
                }
            }
            if($SetDisabled)
            {
                try
                {
                    $message = "Setting service $service : Disabled"
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message $message
                    Set-Service -Name $service -StartupType Disabled -Force
                }
                catch
                {
                    Write-Verbose "$message failed!"
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message "$message failed!"
                }
            }
            if($SetManual)
            {
                try
                {
                    $message = "Setting service $service : Manual"
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message $message
                    Set-Service -Name $service -StartupType Manual -Force
                }
                catch
                {
                    Write-Verbose "$message failed!"
                    Write-EventLog -LogName Application -EntryType Information -EventId 4 -Source $regsource -Message "$message failed!"
                }
            }
        }
    }
    End
    {
    }
}

function Process-Certificate
{
    <#

  .SYNOPSIS

  Import  a certificate from a local or remote system.

  .DESCRIPTION

  Import  a certificate from a local or remote system.

  .PARAMETER  Computername

  A  single or  list of computernames to  perform search against

  .PARAMETER  StoreName

  The  name of  the certificate store name that  you want to search

  .PARAMETER  StoreLocation

  The  location  of the certificate store.

  .NOTES

  Name:  Import-Certificate

  Author:  Boe  Prox

  Version  History:

  1.0  -  Initial Version

  .EXAMPLE

  $File =  "C:\temp\SomeRootCA.cer"

  $Computername = 'Server1','Server2','Client1','Client2'

  Import-Certificate -Certificate $File -StoreName Root -StoreLocation  LocalMachine -ComputerName $Computername

  

  Description

  -----------

  Adds  the SomeRootCA certificate to the Trusted Root Certificate Authority store on  the remote systems.

  #>
    [cmdletbinding(

        SupportsShouldProcess = $True

    )]

    Param (

        [parameter(ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]

        [Alias('PSComputername', '__Server', 'IPAddress')]

        [string[]]$Computername = $env:COMPUTERNAME,

  

        [parameter(Mandatory = $True)]

        [string]$Certificate,

        [System.Security.Cryptography.X509Certificates.StoreName]$StoreName = 'My',

        [System.Security.Cryptography.X509Certificates.StoreLocation]$StoreLocation = 'LocalMachine'

    )

    Begin {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue

        $CertificateObject = New-Object  System.Security.Cryptography.X509Certificates.X509Certificate2

        $CertificateObject.Import($Certificate)

    }

    Process {

        ForEach ($Computer in  $Computername) {

            Try {

                Write-Verbose  ("Connecting to {0}\{1}" -f "\\$($Computer)\$($StoreName)", $StoreLocation)

                $CertStore = New-Object   System.Security.Cryptography.X509Certificates.X509Store  -ArgumentList  "\\$($Computer)\$($StoreName)", $StoreLocation

                $CertStore.Open('ReadWrite')

                If ($PSCmdlet.ShouldProcess("$($StoreName)\$($StoreLocation)", "Add  $Certificate")) {

                    $CertStore.Add($CertificateObject)
                    Write-EventLog -LogName Application -EntryType Information -EventId 6 -Source $regsource -Message "Imported certificate $certificate in storename `'$($StoreName)`' for `'$($StoreLocation)`'"
                }

            }
            Catch {

                Write-Warning  "$($Computer): $_"
                Write-EventLog -LogName Application -EntryType Information -EventId 6 -Source $regsource -Message "Importing certificate $certificate in storename `'$($StoreName)`' for `'$($StoreLocation)`' failed!"

            }

        }

    }

}

function Copy-File
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $FilePath,

        # Param2 help description
        $Destination
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        try
        {
            $message="Copying `'$FilePath`' to `'$Destination`'"
            Write-Verbose $message
            Copy-Item -Path $FilePath -Destination $Destination -Force
            Write-EventLog -LogName Application -EntryType Information -EventId 7 -Source $regsource -Message $message
        }
        catch
        {
            Write-Verbose "$message failed"
            Write-EventLog -LogName Application -EntryType Information -EventId 7 -Source $regsource -Message "$message failed"
        }
    }
    End
    {
    }
}

function Remove-Directory
{
    [CmdletBinding(SupportsShouldProcess=$false)]
    [Alias()]
    [OutputType()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Path,

        # Param2 help description
        [Switch]
        $IncludeSubDirectories
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        if ([System.IO.Directory]::Exists($Path))
        {
            if ($IncludeSubDirectories)
            {
                try
                {
                    $message="Removing directory `'$Path`' including subdirectories"
                    [System.IO.Directory]::Delete($Path,$true)
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 8 -Source $regsource -Message $message
                }
                catch
                {
                    Write-Verbose "$message failed"
                }
                
            }
            else
            {
                try
                {
                    $message="Removing directory `'$Path`'"
                    [System.IO.Directory]::Delete($Path)
                    Write-Verbose $message
                    Write-EventLog -LogName Application -EntryType Information -EventId 8 -Source $regsource -Message $message
                }
                catch
                {
                    Write-Verbose "$message failed"
                }
            }
        }
        else
        {
            $message = "Cannot delete $Path because it does not exist"
            Write-Verbose $message
            Write-EventLog -LogName Application -EntryType Information -EventId 8 -Source $regsource -Message $message
        }
    }
    End
    {
    }
}

function Remove-File
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $FilePath
    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        try
        {
            $message="Removing `'$FilePath`'"
            Write-Verbose $message
            Remove-Item -Path $FilePath -Force
            Write-EventLog -LogName Application -EntryType Information -EventId 8 -Source $regsource -Message $message
        }
        catch
        {
            Write-Verbose "$message failed"
            Write-EventLog -LogName Application -EntryType Information -EventId 8 -Source $regsource -Message "$message failed"
        }
    }
    End
    {
    }
}

function Replace-TextInFile
{
<#
.Synopsis
   Replaces text in files
.DESCRIPTION
   Replaces text in files
.EXAMPLE
   Replace-TextInFile -Path 'P:\Office_Standard_2K10\config - Copy.xml' -TextToReplace 'Completionnotice="no"' -ReplaceWith 'Completionnotice="yes"'
.EXAMPLE
   Another example  of how to use this cmdlet
#>
    [CmdletBinding(SupportsShouldProcess=$false)]
    [OutputType()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true)]
        $Path,

        # The text to be replaced
        [String]
        $TextToReplace,

        # Replace with $ReplaceWith
        [String]
        $ReplaceWith

    )

    Begin
    {
        # create new source for eventlog to write to if not existent
        $regsource = "PowerShell_SCCM_Process"
        Write-Verbose "Creating eventsource $regsource in de Application log"
        New-EventLog -LogName Application -Source $regsource -ErrorAction SilentlyContinue
    }
    Process
    {
        # read filecontents
        try
        {
            $message="Replacing `'$TextToReplace`' with `'$ReplaceWith`' in `'$Path`'"
            Write-Verbose $message

            $reader = [System.IO.StreamReader]$Path
            $data = $reader.ReadToEnd()
            $reader.close()
        }
        catch
        {
            Write-Warning $Error[0].Exception.Message
            Write-EventLog -LogName Application -EntryType Information -EventId 11 -Source $regsource -Message "$message failed!"
            break
        }
        finally
        {
           if ($reader -ne $null)
           {
               $reader.dispose()
           }
        }

        # replace text
        $data = $data -replace $TextToReplace, $ReplaceWith

        # write content back to file
        try
        {
           $writer = [System.IO.StreamWriter]$Path
           $writer.write($data)
           $writer.close()
           Write-EventLog -LogName Application -EntryType Information -EventId 11 -Source $regsource -Message $message
        }
        catch
        {
            Write-EventLog -LogName Application -EntryType Information -EventId 11 -Source $regsource -Message "$message failed!"
            Write-Warning $Error[0].Exception.Message
        }
        finally
        {
           if ($writer -ne $null)
           {
               $writer.dispose()
           }
        }
    }
    End
    {
    }
}

#endregion