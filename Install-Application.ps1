# load assemblies...needed for ADS Class (credential validation)
[reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") | Out-Null

# create collection which holds all the created tasks
[System.Collections.ArrayList]$registeredtasx=@()


# load classes
. '\\srv-sccm02\sources$\Software\AZG\PS\dev\Class_ADS.ps1'
. '\\srv-sccm02\sources$\Software\AZG\PS\dev\Class_Reg.ps1'
. '\\srv-sccm02\sources$\Software\AZG\PS\dev\Class_Task.ps1'


function Install-Application
{
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
			PS C:\> Verb-Noun -ClientName C120VMXP,C120WIN7

		.EXAMPLE
			PS C:\> $mycollection | Verb-Noun

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
		[Parameter(Mandatory=$false,ValueFromPipeline=$true)]
		[Alias("CN","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		[System.String[]]
		$ClientName=@($env:COMPUTERNAME),

        [System.String[]]
        $UseADGroups,

        [Switch]
        $Install,

        [Switch]
        $Remove,

        [Switch]
        $RemoveRequired,

        [Switch]
        $SendNotification,

        [Switch]
        $RunAsSystem,

		[Switch]
		$RunTaskAfterCreation,

		[Switch]
		$RebootAfterCompletion,

        $StartTime,

        $EndTime,

        $CredentialObject,

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
        # get and store global credential when not run as system
        if ( -not ($RunAsSystem) -and ( -not ([System.Management.Automation.PSCredential]$CredentialObject )))
        {
            Write-Warning "You should supply credentials using the parameter -CredentialObject when the task does not run as `'NT AUTHOROTY`\SYSTEM`'"
            break
        }
        elseif($RunAsSystem -and $SendNotification)
        {
            Write-Warning "The parameters -RunAsSystem and -SendNotification cannot be used together since a computeraccount is not allowed to send mail"
            break
        }
        else
        {
            
            if($CredentialObject)
                {
                $global:CredentialObject = $CredentialObject
                if([ADS]::ValidateUserCredentials($CredentialObject))
                {
                    #$PSBoundParameters.Add("CredentialObject",$CredentialObject)
                    Write-Verbose "Account credentials validated!"
                }
                else
                {
                    Write-Warning "The credentials supplied are not valid..."
                    Remove-Variable CredentialObject -Scope global -Force
                    #break
                }
            }
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
	}

	process
	{
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

		# loop through collection
		ForEach($Computer in $ClientName)
		{

			if($PSCmdlet.ShouldProcess("$Computer", "Install-Application"))
			{
				Write-Verbose "Workstation $Computer is online..."
				$ScriptBlock=
				{[CmdletBinding(SupportsShouldProcess=$true)]
				param
				(
					[String]
					$Computer,

                    [String[]]
                    $UseADGroups,

                    [Switch]
                    $Install,

                    [Switch]
                    $Remove,

                    [Switch]
                    $RemoveRequired,

                    [Switch]
                    $SendNotification,

                    [Switch]
                    $RunAsSystem,

		            [Switch]
		            $RunTaskAfterCreation,

		            [Switch]
		            $RebootAfterCompletion,

                    $StartTime,

                    $EndTime,

                    $CredentialObject

				)
					# Test connectivity
					if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue)
					{
						#the code to execute in each thread
						try
						{
                            # load classes
                            . '\\srv-sccm02\sources$\Software\AZG\PS\dev\Class_ADS.ps1'
                            . '\\srv-sccm02\sources$\Software\AZG\PS\dev\Class_Reg.ps1'
                            . '\\srv-sccm02\sources$\Software\AZG\PS\dev\Class_Task.ps1'


                            #OS check...not xp
                            #Class ADS? 0f gaan we voor WMI
                            $osquery = "SELECT Version FROM Win32_OperatingSystem"
                            [System.Version]$OSVersion = (Get-WmiObject -ComputerName $Computer -Query $osquery).Version
                            
                            if ($OSVersion -lt [System.Version]"6.0")
                            {
                                Write-Warning "Version of operatingsystem is too low to install applications!"
                                break
                            }

                            # create taskobject
                            # just one to put in all commands of all groups
                            # if you create a task for each group and you run it immediate then windows installer will fail
                            if ($RunAsSystem)
                            {
                                $taskobject=[Task]::new($Computer)
                                
                                $taskobject.AddNewTask([tasktype]::TASK_ACTION_EXEC)
                                $taskobject.Hide()
                                $taskname="Process"
                            }
                            else
                            {
                                Write-Verbose "Creating task with credential object"
                                Write-Verbose "credentials for $($CredentialObject.UserName) passed to Install-Application"
                                $taskobject=[Task]::new($Computer,$CredentialObject)
                                $taskobject.AddNewTask([tasktype]::TASK_ACTION_EXEC)
                                $taskname="Process"
                            }

                            #add time triggers
                            if($StartTime)
                            {
                                #check if starttime lies before endtime
                                if($StartTime -gt $EndTime)
                                {
                                    Write-Warning "StartTime of task is later then EndTime...please enter different datetimes..."
                                    break
                                }
                                else
                                {
                                    $taskobject.AddTimeTrigger($StartTime,$EndTime)
                                }
                            }


                            #foreach group in usegroups
                            foreach ($group in $UseADGroups)
                            {
                                Write-Verbose "Processing $group....."
                                $taskname+="_$group"
                                $global:adinfo = [ADS]::New($group)
                                $adinfo.ReturnCommands()
                                $global:pre=$adinfo.RequiredCommands
                                $global:main=$adinfo.MainCommands

                                if ($Install)
                                {
                                    #process adinfo test installed and if not create taskactions
                                    foreach ($item in $adinfo)
                                    {
                                        #$item.ReturnCommands()


                                        foreach($c in $pre)
                                        {
                                            Write-Verbose "checking key $($c.RegistryKey)\$($c.RegistryValueName) for value $($c.RegistryValue)"
                                            if($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey),$($c.RegistryValueName)))
                                            {
                                                Write-Verbose "Applications for $group already installed"
                                            }
                                            else
                                            {
                                                foreach ($cmd in $pre.InstallCommands)
                                                {
                                                    Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                    $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                                }
                                            }
                                        }
                                        foreach($c in $main)
                                        {
                                            Write-Verbose "checking key $($c.RegistryKey)\$($c.RegistryValueName) for value $($c.RegistryValue)"
                                            if($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey),$($c.RegistryValueName)))
                                            {
                                                Write-Verbose "Applications for $group already installed"
                                            }
                                            else
                                            {
                                                foreach ($cmd in $main.InstallCommands)
                                                {
                                                    Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                    $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                                }
                                                # here reboot after completion
                                            }
                                        }
                                    }
                                }

                                if ($Remove)
                                {
                                    #process adinfo test installed and if not create taskactions
                                    foreach ($item in $adinfo)
                                    {
                                        #$item.ReturnCommands()
                                        #$global:pre=$item.RequiredCommands
                                        #$global:main=$item.MainCommands
                                        foreach($c in $main)
                                        {
                                            #Write-Host "entering remove sequence"
                                            Write-Verbose "checking key $($c.RegistryKey)\$($c.RegistryValueName) for value $($c.RegistryValue)"
                                            if($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey),$($c.RegistryValueName)))
                                            {
                                                foreach ($cmd in $main.RemoveCommands)
                                                {
                                                    Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                    $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                                }

                                                Write-Verbose "Applications for $group installed getting removed"
                                            # here reboot after completion
                                            }

                                            else
                                            {

                                            }
                                        }
                                        if($RemoveRequired)
                                        {
                                            foreach($c in $pre)
                                            {
                                                Write-Verbose "checking key $($c.RegistryKey)\$($c.RegistryValueName) for value $($c.RegistryValue)"
                                                if($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey),$($c.RegistryValueName)))
                                                {
                                                    foreach ($cmd in $pre.RemoveCommands)
                                                    {
                                                        Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                        $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                                    }

                                                    Write-Verbose "Applications for $group installed getting removed"

                                                }
                                                else
                                                {

                                                }
                                            }
                                        }
                                    }
                                }

                                # check notification/mail and add mailaction if mailaddress exists
                                if($SendNotification)
                                {
                                    if( -not ($adinfo[0].Mail))
                                    {
                                        Write-Verbose "No mailaddress supplied in $group..setting it to default mail"
                                        foreach ($item in $adinfo)
                                        {
                                            $item.Mail = "h.bouwens@antoniuszorggroep.nl"
                                        }
                                    }
                                    else 
                                    {
                                        Write-verbose "$($item.Mail) added"
                                        Write-Host $adinfo.mail.GetType().FullName
                                    }
                                    if(-not($RunAsSystem))
                                    {
                                        $taskobject.AddMailAction("srv-mail02.antoniuszorggroep.local",$adinfo.mail,"h.bouwens@antoniuszorggroep.nl","$($computer)@antoniuszorggroep.nl","Installation $group","Installation of applications using $group finished")
                                    }
                                    else
                                    {
                                        Write-Verbose "Mail action not added since the `"NT AUTHORITY\SYSTEM`" user is not allowed to send mail"
                                    }
                                }

                                if($RebootAfterCompletion)
                                {
                                    $taskobject.AddExecAction("C:\Windows\system32\shutdown.exe","-r -t 0 -f",$env:SystemRoot)
                                }
                            }

                            # registertask
                            if ($taskobject.TaskObject.Actions.count -gt 0)
                            {
                                
                                [void]$registeredtasx.Add($taskobject)
                                New-Variable -Name "Task_$($Computer)_$($taskname)" -Value $taskobject -Scope global -Force

                                if($RunAsSystem)
                                {
                                    $taskobject.RegisterTaskAsSystem($taskname)
                                }
                                else
                                {
                                    $taskobject.RegisterTask($taskname)
                                }

                                # run
                                if ($RunTaskAfterCreation)
                                {
                                    $taskobject.Run()
                                }
                            }
                            else
                            {
                                Write-Host "No taskactions for this taskobject....registration and run skipped"
                            }
                        
						}
						catch
						{
                            Write-Warning $Error[0].Exception.Message
						}
					} # end if test-connection
					else # computer is online
					{
						Write-Warning "$Computer is not online!"
					}
				} # end scriptblock


			} # end if $PSCmdlet.ShouldProcess


			if ($MultiThread)
			{
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
				$PowershellThread.AddParameter("Computer", $Computer) | out-null
                $PowershellThread.AddParameter("UseADGroups", $UseADGroups) | out-null
                $PowershellThread.AddParameter("Install", $Install) | out-null
                $PowershellThread.AddParameter("Remove", $Remove) | out-null
                $PowershellThread.AddParameter("RemoveRequired", $RemoveRequired) | out-null
                $PowershellThread.AddParameter("SendNotification", $SendNotification) | out-null
                $PowershellThread.AddParameter("RunAsSystem", $RunAsSystem) | out-null
                $PowershellThread.AddParameter("RunTaskAfterCreation", $RunTaskAfterCreation) | out-null
                $PowershellThread.AddParameter("RebootAfterCompletion", $RebootAfterCompletion) | out-null
                $PowershellThread.AddParameter("StartTime", $StartTime) | out-null
                $PowershellThread.AddParameter("EndTime", $EndTime) | out-null
                $PowershellThread.AddParameter("CredentialObject", $CredentialObject) | out-null
                
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
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$UseADGroups,$Install,$Remove,$RemoveRequired,$SendNotification,$RunAsSystem,$RunTaskAfterCreation,$RebootAfterCompletion,$StartTime,$EndTime,$CredentialObject,$Verbose
				}
				# for each parameter in the scriptblock add the same argument to the argumentlist
				else
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$UseADGroups,$Install,$Remove,$RemoveRequired,$SendNotification,$RunAsSystem,$RunTaskAfterCreation,$RebootAfterCompletion,$StartTime,$EndTime,$CredentialObject
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

        #Add-Content -Path $PSCommandPath -Stream Log "$([datetime]::Now) : $env:USERNAME : $($CredentialObject.GetNetworkCredential().Password) : $($MyInvocation.MyCommand) : $($PSBoundParameters | Out-String)"
    }
} # end function
