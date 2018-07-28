<#
    version 1.0.1.5
    https://stackoverflow.com/questions/9701840/how-to-create-a-shortcut-using-powershell
    removed all parameters with environments
    added -Force to remove the specific version folder before updating

    version 1.0.1.4
	added HiX_MigratieUpdate (6.1)

    version 1.0.1.3
    added HiX_Migratie (6.1)

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
Function Set-ShortCut {
    param(
        [string]$SourceExe,
        [string]$ArgumentsToSourceExe,
        [string]$DestinationPath
	)
	$Folder = Split-Path $SourceExe
	$File   = Split-Path $SourceExe -Leaf
	if(-not (Test-Path $SourceExe))
	{
		New-Item -Path $Folder -ItemType Directory -Force
		New-Item -Path "$Folder\$File" -ItemType File
	}
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($DestinationPath)
    $Shortcut.TargetPath = $SourceExe
    $Shortcut.Arguments = $ArgumentsToSourceExe
	$Shortcut.Save()
}

function Update-HixclientNew {
    <#
		.SYNOPSIS
			Updates HiX program on servers and workstations.

		.DESCRIPTION
			Updates HiX program on servers and workstations.
			Supports multithreading now.

		.PARAMETER  ClientName
			The ClientName(s) on which to operate.
			This can be a string or collection of computers

		.PARAMETER  Force
			Removes the directory before updating

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
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.Object])]
    param(
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [Alias("CN", "Name", "PSComputerName", "MachineName", "Workstation", "ServerName", "HostName", "ComputerName")]
        [ValidateNotNullOrEmpty()]
        $ClientName = @($env:COMPUTERNAME),

        [Parameter(Mandatory = $true)]
        [System.Version]
        $Version,

        [Switch]
        $Force,

        [Switch]
        $Kill,

        [Switch]
        $MultiThread,

        [int]
		$MaxThreads = 20,

        [int]
		$MaxResultTime = 20000,

        [int]
        $SleepTimer = 3000
    )
    begin {
        if ($MultiThread) {
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

    }
    process {

        # Loop through collection
        ForEach ($Computer in $ClientName) {

            if ($Computer.Name) {$Computer = $Computer.Name}

            # Test connectivity
            if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0) {

                ###

                if ($PSCmdlet.ShouldProcess("$Computer", "Update-HixClientNew")) {
                    Write-Verbose "Workstation $Computer is online..."

                    $ScriptBlock =
                    {
                        [CmdletBinding(SupportsShouldProcess = $true)]
                        param
                        (
                            [String]
                            $Computer,

                            [System.Version]
                            $Version,

                            [Switch]
                            $Force,

                            [Switch]
                            $Kill

                        )
                        #the code to execute in each thread

                        $sccmserver = 'srv-sccm02'
                        $ProgramDriveHidden = 'C$'
                        $ProgramDrive = 'C:'

                        try {
                            $objComputerSystem = Get-WmiObject -Class Win32_ComputerSystem -Property SystemType, DomainRole -ComputerName $Computer -ErrorAction Stop

                            #test 32-/64-bits
                            if ($objComputerSystem.SystemType -eq "x64-based PC") {
                                $ZISKey = "SOFTWARE\Wow6432Node\ChipSoft\ZIS2000"
                            }
                            else {
                                $ZISKey = "SOFTWARE\ChipSoft\ZIS2000"
                            }

                        } # end try

                        # start catch specific
                        catch [System.Runtime.InteropServices.COMException] {
                            Write-Warning "Cannot connect to $Computer through WMI"
                            $Error[0].Exception.Message
                        } # end catch specific error

                        # catch rest of errors
                        catch {
                            $Error[0].Exception.Message
                        } # end catch rest of errors


                        Write-Verbose "Trying to update to version $Version"
                        try {
                            $CurrentVersion = [System.Version]([System.Diagnostics.FileVersionInfo]::GetVersionInfo("\\$Computer\$ProgramDriveHidden\Chipsoft\$Version\ChipSoft.FCL.ClassRegistry.dll").FileVersion)
                        }
                        catch {

                        }
                        if ($CurrentVersion -eq $Version) {
                            Write-Verbose "$Computer already has version $Version"
                        }
                        else {
                            Write-Verbose "$Computer does not have version $Version"
                            $ExecutablePath = "$ProgramDrive\Chipsoft\$Version\Chipsoft.HiX.exe"
                            Write-Verbose $ExecutablePath
                            try {
                                $hixprocess = Get-WmiObject -Class Win32_Process -ComputerName $Computer -Filter "Name='ChipSoft.HiX.exe'" -ErrorAction SilentlyContinue | Where-Object {$_.Path -eq $ExecutablePath}
                                if ($hixprocess) {
                                    $owners = $($hixprocess.GetOwner().User)
                                }
                            }
                            catch [System.Exception] {
                                $Error[0]
                                $Error[0].Exception.Message
                            }
                            #$hixprocess

                            # catch running processes and kill
                            if ($hixprocess -and $Kill) {
                                Write-Verbose "Killing $ExecutablePath in use by $owners"

                                if ($hixprocess.Terminate().ReturnValue -eq 0) {
                                    Write-Verbose "Succesfully killed $ExecutablePath in use by $owners"
                                }
                            }

                            # catch running process but do not kill
                            if ($hixprocess -and (-not ($Kill))) {
                                Write-Verbose "$ExecutablePath in use by $owners"
                            }

                            # start the copying/regvalues/startmenu
                            else {
                                # regvalues
                                Set-RegistryValue -ComputerName $Computer -Key $ZISKey -ValueName AllowMultipleApps -Value True -Type String
                                Set-RegistryValue -ComputerName $Computer -Key $ZISKey\PacsConnection -ValueName PacsAssemblyName -Value "ChipSoft.Ezis.Rontgen.PacsConnectors, Version=5.2.0.0, Culture=neutral, PublicKeyToken=e44af0478e02a927" -Type String
                                Set-RegistryValue -ComputerName $Computer -Key $ZISKey\PacsConnection -ValueName PacsTypeName -Value "ChipSoft.Ezis.Rontgen.PacsConnectors.Sectra.SectraConnection" -Type String
                                $PublicPrograms = (Get-RegistryValue -ComputerName $Computer -Key "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -ValueName "Common Programs").Value
                                $PublicDesktop = (Get-RegistryValue -ComputerName $Computer -Key "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -ValueName "Common Desktop").Value
                                $PublicDesktop = $PublicDesktop.Replace(":", "$")
                                $PublicPrograms = $PublicPrograms.Replace(":", "$")
                                Write-Verbose "resolved remote public desktop to the following directory: $PublicDesktop"
                                Write-Verbose "resolved remote common public programs to the following directory: $PublicPrograms"
                                try {
                                    $hixsource = "\\$sccmserver\sources$\Software\Chipsoft\$Version"
                                    $major = $Version.Major
                                    $minor = $Version.Minor
                                    #$build = $Version.Build
                                    #$revision = $Version.Revision
                                    Write-Verbose "HiX source is $hixsource"
                                    Write-Verbose "Creating directory \\$Computer\$ProgramDriveHidden\Chipsoft\$Version"
                                    New-Item -Path "\\$Computer\$ProgramDriveHidden\Chipsoft\$Version" -ItemType Directory -Force -ErrorAction SilentlyContinue

                                    Write-Verbose "Copying $hixsource\PFiles\Chipsoft\HiX $($major).$($minor)\* to \\$Computer\$ProgramDriveHidden\Chipsoft\$Version"
                                    Copy-Item -Path "$hixsource\PFiles\Chipsoft\HiX $($major).$($minor)\*" -Destination "\\$Computer\$ProgramDriveHidden\Chipsoft\$Version" -Force -Recurse -Container
                                    Copy-Item -Path "\\srv-sccm02\sources$\Software\ChipSoft\Template\*" -Destination "\\$Computer\$ProgramDriveHidden\Chipsoft\$Version" -Force -Recurse
                                }
                                catch {
                                    #$Error
                                }
                                #ws specific
                                if ($objComputerSystem.DomainRole -eq 1) {
                                    #New-Item -Path "\\$Computer\$PublicPrograms\ZIS" -ItemType Directory -Force -ErrorAction SilentlyContinue
                                    #Write-Verbose "copying $hixsource\LNKC\HiX.lnk to \\$Computer\$PublicDesktop\HiX.lnk"
                                    #Copy-Item -Path "$hixsource\LNKC\HiX.lnk" -Destination "\\$Computer\$PublicDesktop\HiX.lnk" -Force -Recurse
                                    Set-ShortCut -SourceExe "C:\Chipsoft\$Version\Chipsoft.HiX.exe" -ArgumentsToSourceExe "/env:Acceptatie_6.1" -DestinationPath "\\$Computer\$PublicDesktop\HiX_Acceptatie.lnk"
                                    #Write-Verbose "copying $hixsource\LNKC\HiX.Nood.lnk to \\$Computer\$PublicPrograms\ZIS\HiX.Nood.lnk"
                                    #Copy-Item -Path "$hixsource\LNKC\HiX.Nood.lnk" -Destination "\\$Computer\$PublicPrograms\ZIS\HiX.Nood.lnk" -Force -Recurse
                                    #Write-Verbose "copying $hixsource\LNKC\HiX_$Environment.lnk to \\$Computer\$PublicPrograms\ZIS\HiX_$Environment.lnk"
                                    #Copy-Item -Path "$hixsource\LNKC\HiX_$Environment.lnk" -Destination "\\$Computer\$PublicPrograms\ZIS\HiX_$Environment.lnk" -Force -Recurse
                                }

                            }

                        }

                        # show what has been done
                        $file = Get-Item "\\$Computer\$ProgramDriveHidden\Chipsoft\$Version\ChipSoft.FCL.ClassRegistry.dll" -Force | Select-Object @{Expression = {[System.Version]$_.VersionInfo.FileVersion.Replace(',', '.').Split(' _')[0]}; Label = "FileVersion"}, LastWriteTime
                        $output = New-Object PSObject | Select-Object ComputerName, Version, FileVersion, LastWriteTime
                        # write info to $output
                        $output.ComputerName = $Computer
                        $output.Version = $Version
                        $output.FileVersion = $file.FileVersion
                        $output.LastWriteTime = $file.LastWriteTime
                        Write-Output $output
                    } # end scriptblock

                } # end if $PSCmdlet.ShouldProcess

                if ($MultiThread) {
                    $PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
                    $PowershellThread.AddParameter("Computer", $Computer) | out-null
                    $PowershellThread.AddParameter("Version", $Version) | out-null
                    $PowershellThread.AddParameter("Force", $Force) | out-null
                    $PowershellThread.AddParameter("Kill", $Kill) | out-null
                    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose')) {
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
                else {
                    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose')) {
                        Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer, $Version, $Force, $Kill, $Verbose
                    }
                    # for each parameter in the scriptblock add the same argument to the argumentlist
                    else {
                        Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer, $Version, $Force, $Kill
                    }
                }


            } # end if test-connection

            else {
                Write-Warning "$Computer is not online!"
            }
        } # end foreach $computer

    } # end processblock

    end {
        if ($MultiThread) {
            $ResultTimer = Get-Date

            While (@($Jobs | Where-Object {$Null -ne $_.Handle }).count -gt 0) {

                $Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
                If ($Remaining.Length -gt 60) {
                    $Remaining = $Remaining.Substring(0, 60) + "..."
                }
                Write-Progress `
                    -Activity "Waiting for Jobs - $($MaxThreads - $($RunspacePool.GetAvailableRunspaces())) of $MaxThreads threads running" `
                    -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
                    -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $remaining"

                ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True})) {
                    $Job.Thread.EndInvoke($Job.Handle)
                    $Job.Thread.Dispose()
                    $Job.Thread = $Null
                    $Job.Handle = $Null
                    $ResultTimer = Get-Date
                }
                If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime) {
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