<#
	version 1.0.1.2
	aliases not working as expected when using pipeline and piping different types of objects
	added if($Computer.Name){$Computer=$Computer.Name} in processblock

	version 1.0.1.1
	test connectivity now with wmi

	version 1.0.1
	todo!! change required scripts in scriptblock

	version 1.0.0
	Initial upload
	added aliases to clientname property to support pipeline input from WMI, SCCM and  Active Directory
	removed unused variables
#>

function Update-Hixclient{
	<#
		.SYNOPSIS
			Updates HiX program on servers and workstations.

		.DESCRIPTION
			Updates HiX program on servers and workstations.
			Supports multithreading now.

		.PARAMETER  ClientName
			The ClientName(s) on which to operate.
			This can be a string or collection of computers

		.PARAMETER  Produktie
			Tells the function to update the Produktie files on the client

		.PARAMETER  Acceptatie
			Tells the function to update the Acceptatie files on the client

		.PARAMETER  Update
			Tells the function to update the Update files on the client.
			Usually this one is used for terminal servers where users still have the produktion files in use.
			An administrator can assign the production environment to the update environment which has a different version.
			The day after the production\acceptation files have been updated

		.PARAMETER  Ontwikkel
			Tells the function to update the Ontwikkel/Test files on the client

		.PARAMETER  Support
			Tells the function to update the Support files on the client

		.PARAMETER  Sedatie
			Tells the function to update the Sedatie files on the client

		.PARAMETER  DWH
			Tells the function to update the DWH files on the client

		.PARAMETER  DWH_ACC
			Tells the function to update the DWH Acceptatie files on the client

		.PARAMETER  Kill
			Kills EZIS/HiX on the client when doing an update 

		.PARAMETER  MaxThreads
			[Int]Tells the function how many threads can run simultaneously

		.PARAMETER  MaxResultTime
			[Int]Tells the function how long it should wait before timing out in seconds

		.PARAMETER  SleepTimer
			[Int]Tells the function to check if a thread has finished every $SleepTimer milliseconds

		.EXAMPLE
			PS C:\> Update-Hixclient_MT -ClientName C120VMXP,C120WIN7

		.EXAMPLE
			PS C:\> $mycollection | Update-Hixclient_MT

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
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

        [Parameter(Mandatory=$false)]
        [System.Version]
        $Version,

        [Switch]
        $Produktie,

        [Switch]
        $Acceptatie,
 
        [Switch]
        $Update,

        [Switch]
        $Ontwikkel,

        [Switch]
        $Support,

        [Switch]
        $DWH,

        [Switch]
        $DWH_ACC,

        [Switch]
        $Kill,

        [Switch]
        $MultiThread,

		[int]
		$MaxThreads=20,
		
		[int]
		$MaxResultTime=20000,
		
		[int]
		$SleepTimer=3000
	)
	
	begin
	{
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
	        #$PSBoundParameters
		}
		
        # define environments
        [System.Collections.ArrayList]$Environments=@("Produktie","Acceptatie","Update","Ontwikkel","Support","Sedatie","ZCD","DWH","DWH_ACC")

	}
	process{

		# Loop through collection
		ForEach($Computer in $ClientName){

			if($Computer.Name){$Computer=$Computer.Name}
			
			# Test connectivity
			if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0)
			{

###

                foreach ($Environment in $Environments)
                {

                    if($PSBoundParameters.ContainsKey($Environment)){

                        if($PSCmdlet.ShouldProcess("$Computer", "Update-HixClient"))
                        {
                            Write-Verbose "Workstation $Computer is online..."

                            $ScriptBlock=
                            {
                            [CmdletBinding(SupportsShouldProcess=$true)]
                            param
                            (
                                [String]
                                $Computer,

                                [System.Version]
                                $Version,

                                [System.String]
                                $Environment,
                                
                                [Switch]
                                $Kill

                            )
                                #the code to execute in each thread
                                $sccmserver="srv-sccm02"
                                
                                try
                                {
                                    . \\srv-fs01\Scripts$\ps\Get-RegistryValue.ps1
                                    . \\srv-fs01\Scripts$\ps\Set-RegistryValue.ps1
									. \\srv-fs01\Scripts$\ps\Get-HiXVersion.ps1
                                }
                                catch
                                {
                                    $Error[0].Exception.Message
                                }

                                try{
                                    $objComputerSystem=Get-WmiObject -Class Win32_ComputerSystem -Property SystemType,DomainRole -ComputerName $Computer -ErrorAction Stop

                                    #test server/workstation
                                    if($objComputerSystem.DomainRole -ge 2)
                                    {
                                        if((Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer).Caption -match "2016")
                                        {
                                            $ProgramDriveHidden="C$"
                                            $ProgramDrive="C:"
                                            $HiXSpecific="WS_$Environment"
                                        }
                                        else
                                        {
                                            $ProgramDriveHidden="E$"
                                            $ProgramDrive="E:"
                                            $HiXSpecific="TS_$Environment"
                                        }

                                    }
                                    else 
                                    {
                                        $ProgramDriveHidden="C$"
                                        $ProgramDrive="C:"
                                        $HiXSpecific="WS_$Environment"
                                    }

                                    #test 32-/64-bits
                                    if ($objComputerSystem.SystemType -eq "x64-based PC")
                                    {
                                        $ZISKey="SOFTWARE\Wow6432Node\ChipSoft\ZIS2000"
                                    }
                                    else
                                    {
                                        $ZISKey="SOFTWARE\ChipSoft\ZIS2000"
                                    }

                                } # end try

                                # start catch specific
                                catch [System.Runtime.InteropServices.COMException]
                                {
                                    Write-Warning "Cannot connect to $Computer through WMI"
                                    $Error[0].Exception.Message
                                } # end catch specific error

                                # catch rest of errors
                                catch {
                                    $Error[0].Exception.Message
                                } # end catch rest of errors


                                if(-not ($Version))
                                {
                                    $hix_version=[ADSI]"LDAP://CN=L-APP-HiX_$Environment,OU=Domain Local,OU=Groepen,OU=AZG,DC=antoniuszorggroep,DC=local"
                                    #[System.Version]$Version=$hix_version.extensionattribute3
                                }

                                $Version=[System.Version]$hix_version.extensionattribute3.ToString()
                                Write-Verbose "Trying to update to version $Version"
                                try{
                                    $CurrentVersion=[System.Version]([System.Diagnostics.FileVersionInfo]::GetVersionInfo("\\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment\ChipSoft.FCL.ClassRegistry.dll").FileVersion)
                                }
                                catch
                                {
                                    
                                }
                                if($CurrentVersion -eq $Version)
                                {
                                    Write-Verbose "$Computer already has version $Version"
                                }
                                else
                                {
                                    Write-Verbose "$Computer does not have version $Version"
                                    $ExecutablePath="$ProgramDrive\Chipsoft\HiX_$Environment\Chipsoft.HiX.exe"
                                    Write-Verbose $ExecutablePath
                                    try {
                                        $ezisprocess=Get-WmiObject -Class Win32_Process -ComputerName $Computer -Filter "Name='ChipSoft.HiX.exe'" -ErrorAction SilentlyContinue | Where-Object {$_.Path -eq $ExecutablePath}
										if($ezisprocess)
										{
											$owners=$($ezisprocess.GetOwner().User)
										}
                                    }
                                    catch [System.Exception] {
									$Error[0]
                                        $Error[0].Exception.Message
                                    }
                                    #$ezisprocess

                                    # catch running processes and kill
                                    if($ezisprocess -and $Kill)
                                    {
                                        Write-Verbose "Killing $ExecutablePath in use by $owners"

										if($ezisprocess.Terminate().ReturnValue -eq 0)
										{
											Write-Verbose "Succesfully killed $ExecutablePath in use by $owners"
										}
                                    }

                                    # catch running process but do not kill
                                    if($ezisprocess -and (-not ($Kill)))
                                    {
                                        Write-Verbose "$ExecutablePath in use by $owners"
                                    }
                                    
                                    # start the copying/regvalues/startmenu
                                    else
                                    {
                                        # regvalues
                                        Set-RegistryValue -ComputerName $Computer -Key $ZISKey -ValueName AllowMultipleApps -Value True -Type String
							            Set-RegistryValue -ComputerName $Computer -Key $ZISKey\PacsConnection -ValueName PacsAssemblyName -Value "ChipSoft.Ezis.Rontgen.PacsConnectors, Version=5.2.0.0, Culture=neutral, PublicKeyToken=e44af0478e02a927" -Type String
							            Set-RegistryValue -ComputerName $Computer -Key $ZISKey\PacsConnection -ValueName PacsTypeName -Value "ChipSoft.Ezis.Rontgen.PacsConnectors.Sectra.SectraConnection" -Type String
                                        $PublicPrograms=(Get-RegistryValue -ComputerName $Computer -Key "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -ValueName "Common Programs").Value
                                        $PublicDesktop=(Get-RegistryValue -ComputerName $Computer -Key "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -ValueName "Common Desktop").Value
                                        $PublicDesktop=$PublicDesktop.Replace(":","$")
                                        $PublicPrograms=$PublicPrograms.Replace(":","$")
                                        Write-Verbose "resolved remote public desktop to the following directory: $PublicDesktop"
                                        Write-Verbose "resolved remote common public programs to the following directory: $PublicPrograms"
										try
										{
	                                        $hixsource="\\$sccmserver\sources$\Software\Chipsoft\$Version"
                                            Write-Verbose "HiX source is $hixsource"
                                            Write-Verbose "Creating directory \\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment"
	                                        New-Item -Path "\\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment" -ItemType Directory -Force -ErrorAction SilentlyContinue
                                            
                                            Write-Verbose "Copying $hixsource\PFiles\Chipsoft\HiX 6.0\* to \\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment"
	                                        Copy-Item -Path "$hixsource\PFiles\Chipsoft\HiX 6.0\*" -Destination "\\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment" -Force -Recurse -Container
	                                        Copy-Item -Path "$hixsource\$HiXSpecific\HiX\*" -Destination "\\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment" -Force -Recurse
										}
										catch
										{
											#$Error
										}
                                        #ws specific
                                        if($objComputerSystem.DomainRole -eq 1)
                                        {
                                            New-Item -Path "\\$Computer\$PublicPrograms\ZIS" -ItemType Directory -Force -ErrorAction SilentlyContinue
                                            Write-Verbose "copying $hixsource\LNKC\HiX.lnk to \\$Computer\$PublicDesktop\HiX.lnk"
                                            Copy-Item -Path "$hixsource\LNKC\HiX.lnk" -Destination "\\$Computer\$PublicDesktop\HiX.lnk" -Force -Recurse
                                            Write-Verbose "copying $hixsource\LNKC\HiX.Nood.lnk to \\$Computer\$PublicPrograms\ZIS\HiX.Nood.lnk"
                                            Copy-Item -Path "$hixsource\LNKC\HiX.Nood.lnk" -Destination "\\$Computer\$PublicPrograms\ZIS\HiX.Nood.lnk" -Force -Recurse
                                            Write-Verbose "copying $hixsource\LNKC\HiX_$Environment.lnk to \\$Computer\$PublicPrograms\ZIS\HiX_$Environment.lnk"
                                            Copy-Item -Path "$hixsource\LNKC\HiX_$Environment.lnk" -Destination "\\$Computer\$PublicPrograms\ZIS\HiX_$Environment.lnk" -Force -Recurse
                                        }

                                    }

                                }
								
								# show what has been done
								$file=Get-Item "\\$Computer\$ProgramDriveHidden\Chipsoft\HiX_$Environment\ChipSoft.FCL.ClassRegistry.dll" -Force | Select-Object @{Expression={[System.Version]$_.VersionInfo.FileVersion.Replace(',','.').Split(' _')[0]};Label="FileVersion"},LastWriteTime
								$output=New-Object PSObject | Select-Object ComputerName,Environment,FileVersion,LastWriteTime
								# write info to $output
								$output.ComputerName=$Computer
								$output.Environment=$Environment
								$output.FileVersion=$file.FileVersion
								$output.LastWriteTime=$file.LastWriteTime
								Write-Output $output
                            } # end scriptblock

                        } # end if $PSCmdlet.ShouldProcess

						if($MultiThread)
						{
		                    $PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
		                    $PowershellThread.AddParameter("Computer", $Computer) | out-null
		                    $PowershellThread.AddParameter("Version", $Version) | out-null
		                    $PowershellThread.AddParameter("Environment",$Environment) | out-null
		                    $PowershellThread.AddParameter("Kill",$Kill) | out-null
							if($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
							{
								$PowershellThread.AddParameter("Verbose") | out-null
							}
		                    $PowershellThread.RunspacePool = $RunspacePool
		                    $Handle = $PowershellThread.BeginInvoke()
		                    $Job = "" | Select-Object Handle, Thread, object
		                    $Job.Handle = $Handle
		                    $Job.Thread = $PowershellThread
		                    $Job.Object = "$Computer($Environment)"
		                    $Jobs += $Job
						}
						else
						{
							if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
							{
								Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Version,$Environment,$Kill,$Verbose
							}
							# for each parameter in the scriptblock add the same argument to the argumentlist
							else
							{
								Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Version,$Environment,$Kill
							}
						}
                    } # end if psboundparameter ContainsKey

                } # end foreach environment.

			} # end if test-connection

			else
			{
				Write-Warning "$Computer is not online!"
			}
		} # end foreach $computer

	} # end processblock

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
						Write-Error "Child script appears to be frozen, try increasing MaxResultTime"
						break
					}
					Start-Sleep -Milliseconds $SleepTimer
					
				} 
			$RunspacePool.Close() | Out-Null
			$RunspacePool.Dispose() | Out-Null
		}
	} # end endblock
}