<#
    version 1.0.3.4
    added tests to script outside function
    can be called like pathto\Install-Application -TestGroup <GroupName> -RunTests
    
    version 1.0.3.3
    removed validate credentials
    Cannot find type [System.DirectoryServices.AccountManagement.ContextType]::Domain
    added parametersetname 'Notify'
    ordered parameters differently

    version 1.0.3.2
    todo error handling
    
    version 1.0.3.2
    added alias Install-App
    
    version 1.0.3.1
    changed test for connectivity

    version 1.0.3
    added support for 64-detection of apps

    version 1.0.2
    converted parameter UseADGroups to dynamic parameter
    added aliases to parameter ClientName to support pipeline input from WMI, SCCM and Active Directory

    version 1.0.1
    converted inputparameters $CredentialObject to type [System.Management.Automation.PSCredential]
    changed path of classfiles
    fixed double entries of pre and postcommands

    version 1.0.0
    Initial version

    wishlist:
    automatic removal of tasks after a given period of time
    parametersets...done
    more info in mail regarding installation/version etc.
    different ordering in parameter...done because of the introduction of parametersets but undone by dynamic param
#>
param(
        [String[]]$TestGroup,
        [Switch]$RunTests
    )
# load assemblies...needed for ADS Class (credential validation)
[reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") | Out-Null

# create collection which holds all the created tasks
[System.Collections.ArrayList]$global:registeredtasx=@()


# load classes
. '\\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_ADS.ps1'
. '\\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_Reg.ps1'
. '\\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_Task.ps1'


function Install-Application
{
	<#
		.SYNOPSIS
			Installs applications to local or remote computers using the taskengine.

		.DESCRIPTION
            Installs applications to local or remote computers using the taskengine.
            Commandlines and checkvalues are retrieved from the active directory group objects

		.PARAMETER  ClientName
			The ClientName(s) on which to operate.
            This can be a string or collection

        .PARAMETER UseADGroups
            Dynamic parameter queries Active Directory for possible groups
            Use 'TAB' to cycle through groups

		.PARAMETER MultiThread
			Enable multithreading

		.PARAMETER MaxThreads
			Maximum number of threads to run simultaneously

		.PARAMETER MaxResultTime
			Max time in which a thread must finish(seconds)

		.PARAMETER SleepTimer
			Time to wait between checks if thread has finished

		.EXAMPLE
			PS C:\> Install-Application -ClientName APPV-W7-X86 -Install -RunAsSystem -RunTaskAfterCreation -UseADGroups L-APP-AdobeFlashActiveX -Verbose

		.EXAMPLE
			PS C:\> $mycollection | Install-Application -Install -RunAsSystem -RunTaskAfterCreation -UseADGroups L-APP-AdobeFlashActiveX

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
    [Alias('Install-App')]
	param(
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ParameterSetName="Install")]
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,ParameterSetName="Remove")]
		[Alias("Name","PSComputerName","CN","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		[System.String[]]
        $ClientName=@($env:COMPUTERNAME),
        
        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Notify")]
        [Switch]
        $Install,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Switch]
        $RunAsSystem,

        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
        [Switch]
        $Remove,

        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
        [Switch]
        $RemoveRequired,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Switch]
        $SendNotification,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Switch]
		$RunTaskAfterCreation,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Switch]
        $RebootAfterCompletion,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Switch]
        $ShowError,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
        $StartTime,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
        $EndTime,

        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [System.Management.Automation.PSCredential]
        [Parameter(ParameterSetName="Notify")]
        $CredentialObject,

        # run the script multithreaded against multiple computers
        [Parameter(ParameterSetName="Install")]
        [Parameter(ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Parameter(Mandatory=$false)]
		[Switch]
		$MultiThread,

        # maximum number of threads that can run simultaniously
        [Parameter(Mandatory=$false,ParameterSetName="Install")]
        [Parameter(Mandatory=$false,ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Int]
		$MaxThreads=20,

		# Maximum time(seconds) in which a thread must finish before a timeout occurs
        [Parameter(Mandatory=$false,ParameterSetName="Install")]
        [Parameter(Mandatory=$false,ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Int]
		$MaxResultTime=20,

        [Parameter(Mandatory=$false,ParameterSetName="Install")]
        [Parameter(Mandatory=$false,ParameterSetName="Remove")]
        [Parameter(ParameterSetName="Notify")]
		[Int]
		$SleepTimer=1000
    )
    
    DynamicParam
    {          
        $ParameterAttributes                   = New-Object System.Management.Automation.ParameterAttribute
        $ParameterAttributes.Mandatory         = $true
        $ParameterAttributes.HelpMessage       = "Press `'TAB`' through the collection of groups"
        $ParameterAttributes.ParameterSetName  = '__AllParameterSets'
        $AttributeCollection                   = New-Object  System.Collections.ObjectModel.Collection[System.Attribute]
        $UseADGroups                           = [System.Collections.ArrayList]$list=[adsisearcher]::new([adsi]"LDAP://OU=Applicaties_Script,OU=Groepen,OU=AZG,DC=antoniuszorggroep,DC=local","objectcategory=group","Name").FindAll().Properties.name

        $AttributeCollection.Add($ParameterAttributes)
        $AttributeCollection.Add((New-Object  System.Management.Automation.ValidateSetAttribute($UseADGroups)))

        $RuntimeParameters                     = New-Object System.Management.Automation.RuntimeDefinedParameter('UseADGroups', [System.String[]], $AttributeCollection)

        $RuntimeParametersDictionary           = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        $RuntimeParametersDictionary.Add('UseADGroups', $RuntimeParameters)
        return  $RuntimeParametersDictionary
    }

	begin
	{
        #clear error
        $Error.Clear()

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
                # if([ADS]::ValidateUserCredentials($CredentialObject))
                # {
                #     #$PSBoundParameters.Add("CredentialObject",$CredentialObject)
                #     Write-Verbose "Account credentials validated!"
                # }
                # else
                # {
                #     Write-Warning "The credentials supplied are not valid..."
                #     Remove-Variable CredentialObject -Scope global -Force
                #     #break
                # }
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

		# loop through collection
		ForEach($Computer in $ClientName)
		{

			if($PSCmdlet.ShouldProcess("$Computer", "Install-Application"))
			{

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
                    $RunAsSystem,

                    [Switch]
                    $Remove,

                    [Switch]
                    $RemoveRequired,

                    [Switch]
                    $SendNotification,

		            [Switch]
		            $RunTaskAfterCreation,

		            [Switch]
                    $RebootAfterCompletion,
                    
                    [Switch]
                    $ShowError,
                    
                    $StartTime,

                    $EndTime,

                    [System.Management.Automation.PSCredential]
                    $CredentialObject

                )
					# Test connectivity
					if([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,139).AsyncWaitHandle.WaitOne(1000,$false))
					{
                        #the code to execute in each thread
                        try
						{
                            # load classes
                            . '\\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_ADS.ps1'
                            . '\\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_Reg.ps1'
                            . '\\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_Task.ps1'


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
                            $taskname="Process"

                            if ($RunAsSystem)
                            {
                                $taskobject=[Task]::new($Computer)
                                $taskobject.AddNewTask([tasktype]::TASK_ACTION_EXEC)
                                $taskobject.Hide()
                                #...does not seem to work and task does not get registered
                                # $taskobject.DeleteExpiredTaskAfterXHours(2)
                            }
                            else
                            {
                                Write-Verbose "Creating task with credential object"
                                Write-Verbose "credentials for $($CredentialObject.UserName) passed to Install-Application"
                                $taskobject=[Task]::new($Computer,$CredentialObject)
                                $taskobject.AddNewTask([tasktype]::TASK_ACTION_EXEC)
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
                            foreach ($group in $($PSBoundParameters.UseADGroups))
                            {
                                Write-Verbose "Processing $group....."
                                $taskname+="_$group"
                                $global:adinfo = [ADS]::New($group)
                                $adinfo.ReturnCommands()
                                $global:pre=$adinfo.RequiredCommands
                                $global:main=$adinfo.MainCommands
                                $global:post=$adinfo.PostCommands

                                if ($Install)
                                {
                                    #process adinfo test installed and if not create taskactions
                                    foreach ($item in $pre)
                                    {
                                        Write-Verbose "checking key $($item.RegistryKey)\$($item.RegistryValueName) for value $($item.RegistryValue)"
                                        if (($($item.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($item.RegistryKey),$($item.RegistryValueName))) -or ($($item.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($item.RegistryKey).Replace("SOFTWARE\","SOFTWARE\WOW6432Node\"),$($item.RegistryValueName))))
                                        {
                                            Write-Verbose "Applications for $group already installed"
                                        }
                                        else
                                        {
                                            foreach ($cmd in $item.InstallCommands)
                                            {
                                                Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                            }
                                        }
                                    }

                                    foreach($item in $main)
                                    {
                                        Write-Verbose "checking key $($item.RegistryKey)\$($item.RegistryValueName) for value $($item.RegistryValue)"
                                        if (($($item.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($item.RegistryKey),$($item.RegistryValueName))) -or ($($item.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($item.RegistryKey).Replace("SOFTWARE\","SOFTWARE\WOW6432Node\"),$($item.RegistryValueName))))
                                        {
                                            Write-Verbose "Applications for $group already installed"
                                        }
                                        else
                                        {
                                            foreach ($cmd in $item.InstallCommands)
                                            {
                                                Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                            }
                                            # here reboot after completion
                                        }
                                    }
                                    foreach($item in $post)
                                    {
                                        Write-Verbose "checking key $($item.RegistryKey)\$($item.RegistryValueName) for value $($item.RegistryValue)"
                                        if (($($item.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($item.RegistryKey),$($item.RegistryValueName))) -or ($($item.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($item.RegistryKey).Replace("SOFTWARE\","SOFTWARE\WOW6432Node\"),$($item.RegistryValueName))))
                                        {
                                            Write-Verbose "Applications for $group already installed"
                                        }
                                        else
                                        {
                                            foreach ($cmd in $item.InstallCommands)
                                            {
                                                Write-Verbose "Adding action $($cmd.Split(";")[1]) with arguments $($cmd.Split(";")[2]) in $($cmd.Split(";")[3])"
                                                $taskobject.AddExecAction($($cmd.Split(";")[1]),$($cmd.Split(";")[2]),$($cmd.Split(";")[3]))
                                            }
                                        }
                                    }
                                }
                            }

                            if ($Remove)
                            {
                                #process adinfo test installed and if not create taskactions
                                foreach ($item in $adinfo)
                                {
                                    foreach($c in $main)
                                    {
                                        #Write-Host "entering remove sequence"
                                        Write-Verbose "checking key $($c.RegistryKey)\$($c.RegistryValueName) for value $($c.RegistryValue)"
                                        if(($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey),$($c.RegistryValueName))) -or ($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey.Replace("SOFTWARE\","SOFTWARE\WOW6432Node\")),$($c.RegistryValueName))))
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
                                            if(($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey),$($c.RegistryValueName))) -or ($($c.RegistryValue) -eq [Reg]::GetRegistryValue($Computer,'LocalMachine',$($c.RegistryKey.Replace("SOFTWARE\","SOFTWARE\WOW6432Node\")),$($c.RegistryValueName))))
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
                                    if($Remove)
                                    {
                                        $Action="Removal"
                                    }
                                    if($Install)
                                    {
                                        $Action="Installation"
                                    }

                                    $taskobject.AddMailAction("srv-mail02.antoniuszorggroep.local",$adinfo.mail,"h.bouwens@antoniuszorggroep.nl","$($computer)@antoniuszorggroep.nl","$Action $group","$Action of applications using $group finished")
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
                            
                            # registertask
                            if ($($taskobject.TaskObject.Actions.count) -gt 0)
                            {

                                New-Variable -Name "Task_$($Computer)_$($taskname)" -Value $taskobject -Scope global -Force

                                if($RunAsSystem)
                                {
                                    Write-Verbose "Registering task $taskname"
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
                #Write-Verbose "Processing $($PSBoundParameters.UseADGroups) in multithreadingblock"
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
				$PowershellThread.AddParameter("Computer", $Computer) | out-null
                $PowershellThread.AddParameter("UseADGroups", $($PSBoundParameters.UseADGroups)) | out-null
                $PowershellThread.AddParameter("Install", $Install) | out-null
                $PowershellThread.AddParameter("RunAsSystem", $RunAsSystem) | out-null
                $PowershellThread.AddParameter("Remove", $Remove) | out-null
                $PowershellThread.AddParameter("RemoveRequired", $RemoveRequired) | out-null
                $PowershellThread.AddParameter("SendNotification", $SendNotification) | out-null
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
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$($PSBoundParameters.UseADGroups),$Install,$RunAsSystem,$Remove,$RemoveRequired,$SendNotification,$RunTaskAfterCreation,$RebootAfterCompletion,$ShowError,$StartTime,$EndTime,$CredentialObject,$Verbose
				}
				# for each parameter in the scriptblock add the same argument to the argumentlist
				else
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$($PSBoundParameters.UseADGroups),$Install,$RunAsSystem,$Remove,$RemoveRequired,$SendNotification,$RunTaskAfterCreation,$RebootAfterCompletion,$ShowError,$StartTime,$EndTime,$CredentialObject
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
        if($ShowError)
        {
            foreach($exception in $Error)
            {
                Write-Warning "Error in script $($exception.InvocationInfo.ScriptName) on line $($exception.InvocationInfo.ScriptLineNumber) : $($exception.Exception.Message)"
            }
        }
        #errorhandling
    }
} # end function

if(-not ($TestGroup))
{
    $TestGroup = "L-APP-AdobeFlashActiveX"
}
if($RunTests)
{
    . '\\srv-fs01\Scripts$\ps\Get-Applicaties.ps1'
    $testclients = @('C120WIN7','C120W7X64')
    if(-not ($creds))
    {
        $creds=Get-Credential -UserName $env:USERDOMAIN\$env:USERNAME -Message "Enter your credentials"
    }

    Install-Application -ClientName $testclients -Install -RunAsSystem -RunTaskAfterCreation -ShowError -UseADGroups $TestGroup
    Start-Sleep -Seconds 10
    #Get-Applicaties -ClientName $testclients -DisplayName *TeleQ*

    Install-Application -ClientName $testclients -Remove -RunAsSystem -RunTaskAfterCreation -ShowError -UseADGroups $TestGroup
    Start-Sleep -Seconds 10
    #Get-Applicaties -ClientName $testclients -DisplayName *TeleQ*

    Install-Application -ClientName $testclients -Install -CredentialObject $creds -RunTaskAfterCreation -ShowError -UseADGroups $TestGroup
    Start-Sleep -Seconds 10
    #Get-Applicaties -ClientName $testclients -DisplayName *TeleQ*

    Install-Application -ClientName $testclients -Remove -CredentialObject $creds -RunTaskAfterCreation -ShowError -UseADGroups $TestGroup
    Start-Sleep -Seconds 10
    #Get-Applicaties -ClientName $testclients -DisplayName *TeleQ*

    Install-Application -ClientName $testclients -Install -CredentialObject $creds -RunTaskAfterCreation -SendNotification -ShowError -UseADGroups $TestGroup -Verbose
    Start-Sleep -Seconds 10
    #Get-Applicaties -ClientName $testclients -DisplayName *TeleQ*

    Install-Application -ClientName $testclients -Remove -CredentialObject $creds -RunTaskAfterCreation -SendNotification -ShowError -UseADGroups $TestGroup -Verbose
    Start-Sleep -Seconds 10
    #Get-Applicaties -ClientName $testclients -DisplayName *TeleQ*
}