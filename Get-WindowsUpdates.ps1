<#
	version 1.2.6.4
	fixed some issues with constructing the commandline in Start-ProcessingPatch

	version 1.2.6.3
	removed some comment with code
	fixed line 355 with subject mail(added space)
	
	version 1.2.6.2
	fixed little bug with retrieving filepaths to patch
	sometimes more filepaths are returned which are actually the same path(e.g. a patch is for windows7 but also for windows2008)
	now the last item is picked

	version 1.2.6.1
	build try-catch block in Start-ProcessingPatchFiles
	
	version 1.2.6.0
	filter out SQL updates when run from taskengine
	invoke sql updates when run from explorer

	version 1.2.5.3
	added scantime to output
	
	version 1.2.5.2
	updated Synopsis in Get-WindowsUpdates

	version 1.2.5.1
	aliases not working as expected when using pipeline and piping different types of objects
	added if($Computer.Name){$Computer=$Computer.Name} in processblock
	
	version 1.2.5.0
	test connectivity now with wmi in Get-Memory and Get-WindowsUpdates
	removed test-online function

	version 1.2.4
	added Test-Online function
	[System.Net.Sockets.TcpClient]::new() relies on Windows Management Framework 5.1 which is not installed everywhere

	version 1.2.3
	added aliases to clientname in function Get-Memory
	added some verbosing...more to follow

	version 1.2.2
	add updatetype[] parameter with defaultvalues and remove the includeoffice and includesqlupdates switch
	remove search through the different updatefolders and make it a search through the root
	this is slower but the code smaller and less complicated...adapt resolve-path to return only one value(first?)
	made status parameter multivalued so we can use it like -Status Missing,Installed
	remove outputfilter for updaterollups..filtering will be applied by inputparameter
	remove test-pipeline input..added aliases to clientname parameter to support WMI,AD and SCCM
	made bulletinparameter multivalued and renamed it to Bulletins
	
	version 1.2.1
	output filter applied...do not show updaterollups

	version 1.2.0
	added parameter SkipFileCheck
	this will return all advertised updates

	version 1.1.5
	Article parameter multivalued like 29334545,12335211,24556222

	version 1.1.4
	added counter to mailsubject and operation to mailbody 

	initial version 1.1.3
	ported from Get-MissingWindowsUpdates script to make it more generic
	implemented switch Status with two options
	implemented removal of patches

	version 1.1.2
	created a fancier counter in function Start-ProcessingPatchFiles

	version 1.1.1
	Fixed a bug in the multithreadingpart
	SetlabeledUri and ClearLabeledUriBefore were still in the scriptblock...thes are removed
	Added some verbosing to the Get-MissingWindowsUpdates

	version 1.1.0
	Added systeembeheer to Cc in function Start-ProcessingPatchFiles

	version 1.0.9
	added statements to catchblock around line 199

	version 1.0.8
	rewritten synopsis for Invoke-PatchProcess
	added -erroraction stop to get-wmiobject

	version 1.0.7
	removed unapproved verbs and implemented SupportsShouldProcess
	removed SetlabeledUri and ClearLabeledUriBefore from function Get-MissingWidowsUpdates and the scriptblock
	(is obsolete now because updates are checked from the local store instead of the labeledUri object in Active Directory)
	added try catch block to address elevation issue

	version 1.0.6
	set $subject properly in 

	version 1.0.5
	on break(CTRL+C) still send the updatereport for the updates installed so far
	move processblock of process-patchfiles to endblock to support counting number of piped in objects

	version 1.0.4
	fixed Out-HTMLDataTable not to output empty tablerows
	fixed creating computername directory under temp to extract patches

	version 1.0.3
	bug resolved $result+= op_addition failed
	added scriptvariable $mailresult

	version 1.0.2
	return mailadress added from computerobject
	out-htmldatatable function added
	switch sendnotification added in process-patchfiles

	version 1.0.1
	.cab files are now processed with dism.exe instead of pkgmgr
	bug solved with the switches -IncludeOfficeUpdates and -IncludeSQLUpdates


	wishlist
	mail report when ready...done
	on break send report of patches installed so far...done
	counter on processpatchfiles...done
	make more generic(filter on status)(remove patchfiles)...done
	make Article parameter multivalued like 29334545,12335211,24556222...done
	implement classifications translationtable...see https://msdn.microsoft.com/en-us/library/windows/desktop/ff357803(v=vs.85).aspx
	investigate hidden property
	a more speedy test-connection...done
	add default set of update types which we use most?
#> 

function Start-ProcessingPatchFiles
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.String])]
	param(
		[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String[]]
		$FilePath,
        
        [ValidateSet('Install','Remove')]
		[String]$Operation = 'Install',
		
		[Parameter(Mandatory=$false)]
		[System.Int16]
        $MinFreeMemPercentage = 10,
        
        [Switch]
        $SendNotification
	)
    begin
    {
        [System.Collections.Arraylist]$script:mailresult = @()
        [System.Collections.Arraylist]$bucket = @()
        $subject = "UpdateReport $env:COMPUTERNAME"
    }
    
    process
    {
        [void]$bucket.Add($FilePath)
    }
    
    end
    {
        Write-Host "Processing $($bucket.Count) updates" -ForegroundColor Green
        try{
            
            foreach($File in $bucket){

                #region Setting ProcessPath and Argumentlist
                $extension=([System.IO.Path]::GetExtension($($File)))
                Write-Verbose "extension is $extension."
                if ($extension -eq ".cab"){
                    if ($Operation -eq 'Install') {
                        $patch_arguments="/Online","/Add-Package","/PackagePath:`"$($File)`"","/quiet","/norestart"
                    }
                    if ($Operation -eq 'Remove') {
                        $patch_arguments="/Online","/Remove-Package","/PackagePath:`"$($File)`"","/quiet","/norestart"
                    }
                    $procname="C:\Windows\System32\Dism.exe"
                    Write-Verbose "argument for start-process is $patch_arguments"
                }
                elseif ($extension -eq ".msu")
                {
                    if ($Operation -eq 'Install')
                    {
                        $patch_arguments="`"$($File)`"","/quiet","/norestart"
                    }
                    if ($Operation -eq 'Remove')
                    {
                        $patch_arguments="/uninstall","`"$($File)`"","/quiet","/norestart"
                    }
                    $procname="C:\Windows\System32\wusa.exe"
                    Write-Verbose "argument for start-process is $patch_arguments"
                    }
                elseif ($extension -eq ".exe")
                {
                    $procname="`"$File`""
                    if(($File -match "rvkroots") -or ($File -match "KB890830") -or ($File -match "MSIPatchRegFix"))
                    {
                        if ($Operation -eq 'Install')
                        {
                            $patch_arguments="/Q"
                        }
                        
                    }
                    elseif ($File -match "ndp")
                    {
                        if ($Operation -eq 'Install') {
                            $patch_arguments="/q","/norestart"
                        }
                        if ($Operation -eq 'Remove') {
                            $patch_arguments="/uninstall","/q","/norestart"
                        }
                        
					}
					elseif ($File -match "SQL")
					{
						# find out if we are running from explorer or taskengine
						$ParentProcessName = (Get-WmiObject -Class Win32_Process -Filter "ProcessID='$PID'").ParentProcessID | ForEach-Object {Get-Process -Id $_ | Select-Object -ExpandProperty Name}
						if(($ParentProcessName -eq  'explorer') -or ($ParentProcessName -eq 'winpty-agent'))
						{
							$procname=$File
							$patch_arguments='noarguments'
							Write-Host "Starting $File manually"
						}
						if($ParentProcessName -eq  "taskeng")
						{
							Write-Host "Skipping $File"
							$procname='donothing'
							$patch_arguments='noarguments'
						}
					}
                    else
                    {
                        if ($Operation -eq 'Install') {
                            $patch_arguments="/quiet","/norestart"
                        }
                        if ($Operation -eq 'Remove') {
                            $patch_arguments="/uninstall","/quiet","/norestart"
                        }
                    }
                    
                    Write-Verbose "argument for start-process is $patch_arguments"
                }
                
                else{
                    Write-Warning "$extension for `"$($File)`" is not a recognized extension!"
                    $patch_arguments=$null
                    $procname=$null
                }
                #endregion

                #region Start and evaluating patchprocess
				try
				{
					$processinfo=Invoke-PatchProcess -ProcessName $procname -Argument $patch_arguments

                    if($processinfo[1] -eq 0){
                        $message="`'$File`' succesfully processed!"
                    }
                    elseif(($processinfo[1] -eq 3010) -or ($processinfo -eq 1641)){
                        $message="Installation of `'$File`' succesfully processed but reboot is required."
                    }
                    elseif($processinfo[1] -eq 1642){
                        $message="Product to patch not found...perhaps another version installed"
                    }
                    elseif($processinfo[1] -eq [int]-196608){
                        $message="Probably not sufficient rights!"
					}
					elseif($processinfo[1] -eq [int]1223){
                        $message="Canceled by user"
                    }
                    elseif($processinfo[1] -eq 2){
                        if($File -match ".exe")
                        {
                            Write-Verbose "$($File) matches .exe...we do not want to extract this"}
                        else
                        {
                            Write-Verbose "Installation of $($File) failed (exitcode $($processinfo[1]))....this is a known issue.(expanding cab ourselves and process again.)"
                            New-Item -Path "\\srv-sccm02\sources$\Software\AZG\PS\Temp\$($env:COMPUTERNAME)" -ItemType Directory -Force
                            ExpandCab -Path $File -Destination "\\srv-sccm02\sources$\Software\AZG\PS\Temp\$($env:COMPUTERNAME)\"
                            $msp=Get-ChildItem "\\srv-sccm02\sources$\Software\AZG\PS\Temp\$($env:COMPUTERNAME)" -Filter *.msp | Select-Object -ExpandProperty FullName
                            Write-Verbose "mspfilepath is $($msp)"
                            $processinfo=Invoke-PatchProcess -ProcessName C:\Windows\System32\msiexec.exe -Arguments "/update", "$($msp)", "/norestart", "/quiet"
                            #Start-Sleep -Seconds 3
                            Remove-Item -Path "\\srv-sccm02\sources$\Software\AZG\PS\Temp\*" -Force -Recurse
                            $message="Installed $msp (exitcode $($processinfo[1]))"
                        }
                    }

                    elseif($processinfo[1] -eq 3){
                        $message="Installation of `'$File`' failed (exitcode $($processinfo[1]))....folderpath not found"
                    }

                    elseif($processinfo[1] -eq -2145124329){
                        $message="Installation of `'$File`' failed (exitcode $($processinfo[1]))....Operation was not performed because there are no applicable updates."
                    }

                    else{
                        $message="Installation of `'$File`' failed with exitcode $($processinfo[1])"
                        #$exitcode | Select ExitCode
					}

					Write-Verbose $message

				}
				catch
				{

				}
                    
                #endregion
                
                #region creating custom output object
                $output             = New-Object PSObject | Select-Object ProcessName,PatchFile,Parameters,ExitCode,Message,FreeVirtualMemoryPercentage,IsDeployed,Operation
                $output.ProcessName = $procname
                $output.PatchFile   = "`'$File`'"
                $output.Parameters  = $patch_arguments | Out-String
                $output.ExitCode    = $processinfo[1]
                $output.Message     = $message
                $output.Operation   = $Operation
                $output.FreeVirtualMemoryPercentage=Get-Memory | Select-Object -ExpandProperty FreeVirtualPercentage
				if(($output.exitcode -eq 0) -or ($output.exitcode -eq 3010) -or ($output.exitcode -eq 1642) -or ($output.exitcode -eq 1641) -or ($output.ExitCode -eq -2145124329))
				{
                    $output.IsDeployed=$true
                }
				else
				{
                    $output.IsDeployed=$false
                }
                #endregion
                
                $counter+=1
                Write-Host "Number of patches processed: $counter of $($bucket.Count)" -ForegroundColor Green
                if ($output.FreeVirtualMemoryPercentage -lt $MinFreeMemPercentage)
                    {
                        $memwarning="...Memory less than $MinFreeMemPercentage percent"
                        Write-Warning $memwarning
                        $subject+="($memwarning)"
                        $breakprocess = $true
                    }
                [void]$mailresult.Add($output)
                $output

                # test break 
                if ($breakprocess) 
                {
                    break
                }
            }
        }
		catch
		{
			Write-Warning -Message $Error[0].Exception.Message
		}
        finally
        {
            if($SendNotification)
            {
                $subject += " $($Operation) ($($counter) updates)"
                $mailbody = $mailresult | Select-Object PatchFile,Parameters,ExitCode,Message,IsDeployed,ProcessName,Operation | Out-HTMLDataTable -AsString
                Send-MailMessage -From $env:COMPUTERNAME@antoniuszorggroep.nl -Subject $subject -SmtpServer "srv-mail02.antoniuszorggroep.local" -Cc @("h.bouwens@antoniuszorggroep.nl","systeembeheer@antoniuszorggroep.nl") -To (Get-MailAdresBeheerder) -Body $mailbody -BodyAsHtml
            }
            Remove-Item -Path "\\srv-sccm02\sources$\Software\AZG\PS\Temp\*" -Force -Recurse
        }
	}

}

function Get-MailAdresBeheerder
{
	try{
		$sName=$env:COMPUTERNAME
		$adsisearcher = [ADSISearcher]"(sAMAccountName=$($sName)$)"
		$objpath_adsi = $adsisearcher.FindOne()
		[ADSI]$script:obj_adsi=$objpath_adsi.Path
	}
	catch{
        Write-Error $Error[0].Exception.Message
        return "h.bouwens@antoniuszorggroep.nl"
	}
    if( -not ($obj_adsi.mail))
    {
        return "h.bouwens@antoniuszorggroep.nl"
    }
    else
    {
		Write-Verbose "Setting mailaddress to $($obj_adsi.mail)"
		return $obj_adsi.mail
	}

}

function Out-HTMLDataTable
{
    <#
        .Synopsis
        Short description
        .DESCRIPTION
        Long description
        .EXAMPLE
        Example of how to use this cmdlet
        .EXAMPLE
        Another example of how to use this cmdlet
        .INPUTS
        Inputs to this cmdlet (if any)
        .OUTPUTS
        Output from this cmdlet (if any)
        .NOTES
        General notes
        .COMPONENT
        The component this cmdlet belongs to
        .ROLE
        The role this cmdlet belongs to
        .FUNCTIONALITY
        The functionality that best describes this cmdlet
    #>
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([System.Object])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, ValueFromRemainingArguments=$false)]
        $InputObject,
		
		[Switch]
		$AsString
    )

    Begin
    {
    #styling voor de tabel in de emailbody
    $style = "<style>BODY{font-family: Arial; font-size: 8pt;} "
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;} "
    $style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; } "
    $style = $style + "TD{border: 1px solid black; padding: 5px; } "
    $style = $style + "</style>"
    [System.Collections.ArrayList]$temptable=@()
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("$InputObject", "Out-HTMLDataTable"))
        {
            
			foreach($item in $InputObject)
			{
				
				$addrow = $false

				foreach($prop in $item)
				{
                    if ( -not ([String]::IsNullOrEmpty($prop)))
                    {
						$addrow = $true
					}
					else
					{
						#$addrow = $true
					}
                }
                if ($addrow)
                {
					[void]$temptable.Add($item)
					$addrow = $false
                }
            }
        }
        
        $output=$temptable  | ConvertTo-Html -Head $style
        
    }
    End
    {
		if($AsString){
			$output | Out-String
		}
		else{
        	$output
		}
    }
}

function ExpandCab
{
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([String])]
	param(
		$Path,
		$Destination
	)
    if( -not (Test-Path "$Home\PatchTemp")){
        New-Item -Path "$Home\PatchTemp" -ItemType Directory
    }

    C:\Windows\System32\expand.exe $Path -F:* $Destination
    Start-Sleep -Seconds 5

}

function Invoke-PatchProcess
{
	<#
		.SYNOPSIS
			Starts the patchprocess of given patchfile.

		.DESCRIPTION
			Starts process with arguments and returns exitcode.

		.PARAMETER  ProcessName
			The description of the ProcessName parameter.

		.PARAMETER  Arguments
			The description of the Arguments parameter.

		.EXAMPLE
			Invoke-PatchProcess -ProcessName C:\Windows\System32\msiexec.exe -Arguments "/update", "<path to .msp file>", "/norestart", "/quiet"

		.INPUTS
			System.String

		.OUTPUTS
			System.Int32

		.NOTES
			Additional information about the function go here.

		.LINK
			about_functions_advanced

		.LINK
			about_comment_based_help

	#>
	[CmdletBinding()]
	[OutputType([System.Int32])]
	param(
		[Parameter(Mandatory=$false)]
		$ProcessName,

		[Parameter(Mandatory=$false)]
		$Arguments
	)
	begin
	{

	}
	process
	{
		try
		{
			Write-Verbose "Starting process $ProcessName with arguments $Arguments"
			if($ProcessName -eq 'donothing')
			{
				#canceled by user
				return 1223
			}
			if($Arguments -eq "noarguments")
			{
				Start-Process "$ProcessName" -Passthru -Wait -OutVariable var_result
			}
			else
			{
				Start-Process "$ProcessName" -ArgumentList $Arguments -Passthru -Wait -OutVariable var_result
			}
			
			Write-Verbose "returning exitcode $($var_result.Item(0).ExitCode) from function Invoke-PatchProcess"

			return $($var_result.Item(0).ExitCode)
		}
		catch {
			Write-Error $Error[0].Exception.Message
		}
	}
	end
	{

	}
}

Function Get-Memory
{
	<#
		.SYNOPSIS
			Returns memory info of remote and local computer.

		.DESCRIPTION
			Returns memory info of remote and local computer.

		.PARAMETER ClientName
			The ComputerName(s) on which to operate.(Accepts value from pipeline)

		.EXAMPLE
			Get-Memory -ClientName 'C120VMXP','C120WIN7'

		.EXAMPLE
			'C120VMXP','C120WIN7' | Get-Memory

		.EXAMPLE
			Get-Memory -ClientName (Get-Content C:\computers.txt)

		.EXAMPLE
			(Get-Content C:\computers.txt) | Get-Memory

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
		[Parameter(Position=0, Mandatory=$false,ValueFromPipeline=$true)]
		[Alias("Name","PSComputerName","ComputerName","CN","MachineName","Workstation","ServerName","HostName")]
		$ClientName=@($env:COMPUTERNAME)
	)

	# set initial values in the begin block (populate variables, check dependent modules etc.)
	begin {

	} # end beginblock

	# processblock
	process {

		# add -Whatif and -Confirm support to the CmdLet
		if($PSCmdlet.ShouldProcess("$ClientName", "Get-Memory")){

			# loop through collection $ClientName
			ForEach($Computer in $ClientName){

				# test connection to each $Computer
				if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0) 
				{

					Write-Verbose "$Computer is online..."

					# start try
					try{
						$os=Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ErrorAction Stop
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


					# do the things you want to do after this line
					$output=New-Object -TypeName PSObject | Select-Object ComputerName,FreePhysicalMemory,FreeVirtualMemory,TotalVirtualMemorySize,TotalVisibleMemorySize,MaxProcessMemorySize,FreePhysicalPercentage,FreeVirtualPercentage

					# fill the $output objects properties with the proper values
					$output.ComputerName=$Computer
					$output.FreePhysicalMemory=$os.FreePhysicalMemory
					$output.FreeVirtualMemory=$os.FreeVirtualMemory
					$output.TotalVirtualMemorySize=$os.TotalVirtualMemorySize
					$output.TotalVisibleMemorySize=$os.TotalVisibleMemorySize
					$output.MaxProcessMemorySize=$os.MaxProcessMemorySize
					$output.FreePhysicalPercentage=($os.FreePhysicalMemory/$os.TotalVisibleMemorySize)*100
					$output.FreeVirtualPercentage=($os.FreeVirtualMemory/$os.TotalVirtualMemorySize)*100
					Write-Output $output
				}

			# else $Computer not online
			else {
				Write-Warning "$Computer is not online!"
				} # end test connection to each $Computer

			} # end for each $Computer

		} # end if $pscmdlet.ShouldProcess

	} # end processblock

	# remove variables  in the endblock
	end {
		try{

		}
		catch [System.Management.Automation.ItemNotFoundException] {
			Write-Warning $Error[0].Exception.Message
		}
	} # end endblock

}

function Get-WindowsUpdates
{
	<#
		.SYNOPSIS
			Retrieves list of missing updates.

		.DESCRIPTION
			Retrieves list of missing updates.
			The function uses the class 'CCM_UpdateStatus' in the namespace 'ROOT\ccm\SoftwareUpdates\UpdatesStore'

		.PARAMETER  ClientName
			The ClientName(s) on which to operate.
			This can be a string or collection

        .Parameter Status
			Filter on Installed or Missing updates
			
		.Parameter SkipFileCheck
			By using this parameter even the not downloaded updates will be displayed

        .Parameter Article
            The kb article(s) to query

        .Parameter Bulletin
            The bulletin to query

		.PARAMETER MultiThread
			Enable multithreading

		.PARAMETER MaxThreads
			Maximum number of threads to run simultaneously

		.PARAMETER MaxResultTime
			Max time in which a thread must finish(seconds)

		.PARAMETER SleepTimer
			Time to wait between checks if thread has finished

		.EXAMPLE
			PS C:\> Get-WindowsUpdates -ClientName C120VMXP,C120WIN7 -IncludeOfficeUpdates

		.EXAMPLE
			PS C:\> Get-WindowsUpdates -IncludeOfficeUpdates | Process-PatchFiles -Verbose
			Installs security and officeupdates on the current host

		.EXAMPLE
			PS C:\> $mycollection | Get-WindowsUpdates

		.INPUTS
			System.String,System.String[],Switch

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
		[Parameter(Mandatory = $false,ValueFromPipeline = $true)]
		[Alias("Name","PSComputerName","CN","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		#[System.String[]]
		$ClientName=@($env:COMPUTERNAME),

        [String[]]
		[Parameter(Mandatory = $false)]
		[ValidateSet('Missing','Installed')]
        $Status = 'Missing',

		[String[]]
		[Parameter(Mandatory = $false)]
		[ValidateSet('Application','Connectors','CriticalUpdates','DefinitionUpdates','DeveloperKits','Guidance','OfficeUpdates','SecurityUpdates','ServicePacks','Tools','UpdateRollups','Updates','SCOM2012R2','SQL2016_SecurityUpdates')]
		#maybe application in default??
		$UpdateType = @('CriticalUpdates','OfficeUpdates','SecurityUpdates','ServicePacks','Updates','SCOM2012R2','SQL2016_SecurityUpdates'),

		[String[]]
		$Articles,

		[String[]]
		$Bulletins,

		[Switch]
		$SkipFileCheck,

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
			if($PSCmdlet.ShouldProcess("$Computer", "Get-WindowsUpdates"))
			{
				#Write-Verbose "Workstation $Computer is online..."

				$ScriptBlock=
				{[CmdletBinding(SupportsShouldProcess=$true)]
				param
				(
					[String]
					$Computer,

					$Status,
					
					[System.String[]]
					$UpdateType,

					[String[]]
					$Articles,

					[String[]]
					$Bulletins,

					[Switch]
					$SkipFileCheck
			
				)
					# Test connectivity
					#if([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,139).AsyncWaitHandle.WaitOne(1000,$false))
					if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$Computer') and timeout=1000").StatusCode -eq 0) 
					{

						# hashtable to translate ClassificationTypes to GUID
						$ClassificationTypes2GUID = @{
							Application             = "5C9376AB-8CE6-464A-B136-22113DD69801"
							Connectors              = "434DE588-ED14-48F5-8EED-A15E09A991F6"
							CriticalUpdates         = "E6CF1350-C01B-414D-A61F-263D14D133B4"
							DefinitionUpdates       = "E0789628-CE08-4437-BE74-2495B842F43B"
							DeveloperKits           = "E140075D-8433-45C3-AD87-E72345B36078"
							Guidance                = "9511D615-35B2-47BB-927F-F73D8E9260BB"
							OfficeUpdates           = "84F5F325-30D7-41C4-81D1-87A0E6535B66"
							SecurityUpdates         = "0FA1201D-4330-4FA8-8AE9-B877473B6441"
							ServicePacks            = "68C5B0A3-D1A6-4553-AE49-01D3A7827828"
							Tools                   = "B4832BD8-E735-4761-8DAF-37F882276DAB"
							UpdateRollups           = "28BC880E-0592-4CBF-8F95-C79B17911D5F"
							Updates                 = "CD5FFD1E-E932-4E3A-BF74-18BF0B1BBD83"
							SCOM2012R2              = "2A9170D5-3434-4820-885C-61A4F3FC6F84"
							SQL2016_SecurityUpdates = "93F0B0BC-9C20-4CA5-B630-06EB4706A447"
						}

						# reverse the $ClassificationTypes2GUID hashtable to make reverse lookup possible
						# hashtable to translate GUID to ClassificationType
						$ClassificationGUIDs2Type=@{}
						foreach ($key  in $ClassificationTypes2GUID.Keys) {
							$ClassificationGUIDs2Type.Add($ClassificationTypes2GUID[$key],$key)
						}

						$UpdateClassification=$ClassificationTypes2GUID[$UpdateType]

						# scantime
						$scantime = @{
							Name = 'ScanTime'
							Expression = {[System.Management.ManagementDateTimeConverter]::ToDateTime($_.ScanTime)}
						}

						#the code to execute in each thread
						try
						{
							$global:qry="Select * from CCM_UpdateStatus"
                            #$article.Replace("KB","")

							if ($Status.Count -eq 1)
							{
								$qry="$qry WHERE Status = '$($Status)'"
							}
							elseif($Status.Count -eq 2) 
							{
								$qry="$qry WHERE Status = 'Installed' or Status='Missing'"
							}
							
							if($UpdateType)
							{
								$tempqry=$UpdateClassification -join ("`' or UpdateClassification = `'") -join ("`' or UpdateClassification = `'")
								$qry="$qry And (UpdateClassification = `'$tempqry`')"
								#Write-Host $qry
							}

							if($Articles)
							{
								#remove KB from article-string...just in case someone entered them on the commandline
								foreach($Article in $Articles)
								{
									$Article = $Article.Replace("KB","")
								}
                                $tempqry=$Articles -join ("`' or Article = `'")
                                #("Article Like `'%$tempqry%`'")
                                	$qry="$qry And (Article = `'$tempqry`')"
                            }

                            if($Bulletins)
                            {
								$tempqry=$ArticleBulletinss -join ("`' or Bulletin Like `'")
                            	#("Article Like `'%$tempqry%`'")
                            	$qry="$qry And (Bulletin = `'$tempqry`')"
							}

                            try
                            {
                            	#$computer=$env:COMPUTERNAME
								$global:result=Get-WmiObject -ComputerName $($Computer) -Namespace ROOT\ccm\SoftwareUpdates\UpdatesStore -Query $qry -ErrorAction Stop
								foreach ($item in $result)
								{
									if($ClassificationGUIDs2Type.$($item.ProductID))
									{
										$item.ProductID = $ClassificationGUIDs2Type.$($item.ProductID)
									}
								}
							
							}
							catch [System.Management.ManagementException]
							{
								Write-Warning "$($Error[0].Exception)"
								break
							}
							catch
							{
								Write-Warning -Message $Error[0].Exception.Message
								Write-Warning -Message "exiting script"
								break
							}
								
							if (-not ($SkipFileCheck))
							{
								Write-Verbose "Searching in directories \\srv-sccm02\Packages$\Updates\*\*"
								foreach($item in $result)
								{
									Write-Verbose "resolving path \\srv-sccm02\Packages$\Updates\*\$($item.UniqueId)"
									$ppath=(Resolve-Path "\\srv-sccm02\Packages$\Updates\*\$($item.UniqueId)\*.cab","\\srv-sccm02\Packages$\Updates\*\$($item.UniqueId)\*.exe" -ErrorAction 0).ProviderPath #| Select-Object -First
									if($ppath)
									{
										# sometimes more than one item(same kb) is returned..pick the last one
										if($ppath.count -gt 1)
										{
											$ppath = $ppath[-1]
										}
										Write-Verbose "Adding path $ppath to resultitem"
										Add-Member -InputObject $item -MemberType NoteProperty -Name FilePath -Value $ppath -Force
									}
								}

								#$output=$result | Where-Object {(-not ([System.String]::IsNullOrEmpty($_.FilePath))) -and ($_.ProductID -ne "UpdateRollUps")} | Select-Object PSComputerName,Article,Bulletin,ProductID,Title,UniqueId,Status,FilePath
								$output=$result | Where-Object {(-not ([System.String]::IsNullOrEmpty($_.FilePath)))} | Select-Object PSComputerName,Article,Bulletin,ProductID,Title,UniqueId,Status,FilePath,$scantime

							}
                                #productid zie.... https://msdn.microsoft.com/en-us/library/windows/desktop/ff357803(v=vs.85).aspx
								#$result
							else
							{
								#$output=$result | Where-Object {$_.ProductID -ne "UpdateRollUps"} | Select-Object PSComputerName,Article,Bulletin,ProductID,Title,UniqueId,Status

								$output=$result | Select-Object PSComputerName,Article,Bulletin,ProductID,Title,UniqueId,Status,$scantime
							}
 
							$output
							
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
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
                $PowershellThread.AddParameter("Computer", $Computer) | out-null
				$PowershellThread.AddParameter("Status", $Status) | out-null
				$PowershellThread.AddParameter("UpdateType", $UpdateType) | out-null
                $PowershellThread.AddParameter("Articles", $Articles) | out-null
                $PowershellThread.AddParameter("Bulletins", $Bulletins) | out-null
				$PowershellThread.AddParameter("IncludeSQLUpdates", $IncludeSQLUpdates) | out-null
				$PowershellThread.AddParameter("SkipFileCheck", $SkipFileCheck) | out-null
                
                
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
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Status,$UpdateType,$Articles,$Bulletins,$SkipFileCheck,$Verbose
				}
				# for each parameter in the scriptblock add the same argument to the argumentlist
				else
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Status,$UpdateType,$Articles,$Bulletins,$SkipFileCheck
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
	
}