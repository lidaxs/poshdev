<#
version 1.1.0
Added systeembeheer to Cc in function Start-ProcessingPatchFiles

version 1.0.9
filled catchblock around line 199

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
on break send report of patches installed sofar...done
counter on processpatchfiles...done
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
		
		[Parameter(Mandatory=$false)]
		[System.Int16]
        $MinFreeMemPercentage=10,
        
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
                    $patch_arguments="/Online","/Add-Package","/PackagePath:`"$($File)`"","/quiet","/norestart"
                    $procname="C:\Windows\System32\Dism.exe"
                    Write-Verbose "argument for start-process is $patch_arguments"
                }
                elseif ($extension -eq ".msu"){
                    $patch_arguments="`"$($File)`"","/quiet","/norestart"
                    $procname="C:\Windows\System32\wusa.exe"
                    Write-Verbose "argument for start-process is $patch_arguments"
                    }
                elseif ($extension -eq ".exe"){
                    $procname="`"$File`""
                    if($FilePath -match "rvkroots"){$patch_arguments="/Q"}
                    else{$patch_arguments="/quiet","/norestart"}
                    
                    Write-Verbose "argument for start-process is $patch_arguments"
                }
                
                else{
                    Write-Warning "$extension for `"$($File)`" is not a recognized extension!"
                    $patch_arguments=$null
                    $procname=$null
                }
                #endregion

                #region Start and evaluating patchprocess

                    $processinfo=Invoke-PatchProcess -ProcessName $procname -Argument $patch_arguments
                    #if($exitcode.GetType().Name -eq "Object[]"){Write-Host "Object";$exitcode=$exitcode | Select -expa ExitCode}
                    if($processinfo[1] -eq 0){
                        $message="`'$File`' succesfully installed!"
                    }
                    elseif(($processinfo[1] -eq 3010) -or ($processinfo -eq 1641)){
                        $message="Installation of `'$File`' succesfully finished but reboot is required."
                    }
                    elseif($processinfo[1] -eq 1642){
                        $message="Product to patch not found...perhaps another version installed"
                    }
                    elseif($processinfo[1] -eq [int]-196608){
                        $message="Probably not sufficient rights!"
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
                    
                #endregion
                
                #region Updating labeledUri in AD
                # creating custom output object
                $output = New-Object PSObject | Select-Object ProcessName,PatchFile,Parameters,ExitCode,Message,FreeVirtualMemoryPercentage,IsDeployed
                $output.ProcessName=$procname
                $output.PatchFile="`'$File`'"
                $output.Parameters=$patch_arguments | Out-String
                $output.ExitCode=$processinfo[1]
                $output.Message=$message
                $output.FreeVirtualMemoryPercentage=Get-Memory | Select-Object -ExpandProperty FreeVirtualPercentage
				if(($output.exitcode -eq 0) -or ($output.exitcode -eq 3010)-or ($output.exitcode -eq 1642) -or ($output.exitcode -eq 1641) -or ($output.ExitCode -eq -2145124329))
				{
                    $output.IsDeployed=$true

                }
                else{
                    $output.IsDeployed=$false
                }
                #endregion
                
                $counter+=1
                Write-Host "Number of patches processed: $counter" -ForegroundColor Green
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
                $mailbody = $mailresult | Select-Object PatchFile,Parameters,ExitCode,Message,IsDeployed,ProcessName | Out-HTMLDataTable -AsString
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
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Int32])]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[System.String]
		$ProcessName,

		[Parameter(Mandatory=$false)]
		$Arguments
	)
	begin{}
	process
	{
		if ($PSCmdlet.ShouldProcess("$ProcessName $Arguments","Invoke-PatchProcess")) 
		{
			try
			{
				Write-Verbose "Starting process $ProcessName with arguments $Arguments"
				Start-Process $ProcessName -ArgumentList $Arguments -Passthru -Wait -OutVariable var_result

				Write-Verbose "returning exitcode $($var_result.Item(0).ExitCode) from function Invoke-PatchProcess"

				return $($var_result.Item(0).ExitCode)
			}
			catch {
				Write-Error $Error[0].Exception.Message
			}
		}
	}
	end{}
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
		[Alias("ComputerName","CN","MachineName","Workstation","ServerName","HostName")]
		$ClientName=@($env:COMPUTERNAME)
	)

	# set initial values in the begin block (populate variables, check dependent modules etc.)
	begin {

	} # end beginblock

	# processblock
	process {


		#
		# test pipeline input and pick the right attributes from the incoming objects
		if($ClientName.__NAMESPACE -like 'root\sms\site_*'){
			Write-Verbose "Object received from sccm."
			$ClientName=$ClientName.Name
		}
		elseif($ClientName.objectclass -eq 'computer'){
			Write-Verbose "Object received from Active Directory module."
			$ClientName=$ClientName.Name
		}
		elseif($ClientName.__NAMESPACE -like 'root\cimv2*'){
			Write-Verbose "Object received from WMI"
			$ClientName=$ClientName.PSComputerName
		}
		elseif($ClientName.ComputerName){
			Write-Verbose "Object received from pscustom"
			$ClientName=$ClientName.ComputerName
		}
		else{
			Write-Verbose "No pipeline or no specified attribute from inputobject"
		}
		# end test pipeline input and pick the right attributes from the incoming objects
		#


		# add -Whatif and -Confirm support to the CmdLet
		if($PSCmdlet.ShouldProcess("$ClientName", "Get-Memory")){

			# loop through collection $ClientName
			ForEach($Computer in $ClientName){

				# test connection to each $Computer
				if ( Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue) {

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

function Get-MissingWindowsUpdates
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

		.PARAMETER MultiThread
			Enable multithreading

		.PARAMETER IncludeOfficeUpdates
			Includes officeupdates in result

		.PARAMETER IncludeSQLUpdates
			Includes SQLupdates in result

		.PARAMETER MaxThreads
			Maximum number of threads to run simultaneously

		.PARAMETER MaxResultTime
			Max time in which a thread must finish(seconds)

		.PARAMETER SleepTimer
			Time to wait between checks if thread has finished

		.EXAMPLE
			PS C:\> Get-MissingWindowsUpdates -ClientName C120VMXP,C120WIN7 -IncludeOfficeUpdates

		.EXAMPLE
			PS C:\> Get-MissingWindowsUpdates -IncludeOfficeUpdates | Process-PatchFiles -Verbose
			Installs security and officeupdates on the current host

		.EXAMPLE
			PS C:\> $mycollection | Get-MissingWindowsUpdates

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
		[Parameter(Mandatory=$false,ValueFromPipeline=$true)]
		[Alias("CN","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		#[System.String[]]
		$ClientName=@($env:COMPUTERNAME),

        [String]$Article,

        [String]$Bulletin,

        [Switch]$IncludeOfficeUpdates,

        [Switch]$IncludeSQLUpdates,

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
		# test pipeline input and pick the right attributes from the incoming objects
		if($ClientName.__NAMESPACE -like 'root\sms\site_*')
		{
			Write-Verbose "Object received from sccm."
			$ClientName=$ClientName.Name
		}
		elseif($ClientName.classname -eq 'computer')
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
			$ErrorActionPreference='SilentlyContinue'
			#if([Quest.ActiveRoles.ArsPowerShellSnapIn.Data.ArsComputerObject]$Computer){$Computer=$Computer.Name}
			$ErrorActionPreference='Continue'

			if($PSCmdlet.ShouldProcess("$Computer", "Get-MissingWindowsUpdates"))
			{
				Write-Verbose "Workstation $Computer is online..."

				$ScriptBlock=
				{[CmdletBinding(SupportsShouldProcess=$true)]
				param
				(
					[String]
					$Computer,

                    [String]$Article,

                    [String]$Bulletin,

                    [Switch]$IncludeOfficeUpdates,

                    [Switch]$IncludeSQLUpdates
				)
					# Test connectivity
					if (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue)
					{
						#the code to execute in each thread
						try
						{
							[System.Collections.ArrayList]$OSDirectories=@()
                            $OSCaption=(Get-WmiObject -Class Win32_OperatingSystem -ComputerName $($Computer)).Caption
                            if($OSCaption -match "Server 2008 R2")
                            {
                                [void]$OSDirectories.Add("Windows 2008 R2")
                            }
                            elseif($OSCaption -match "Server 2012 R2")
                            {
								[void]$OSDirectories.Add("Windows 2012 R2")
                            }
                            elseif($OSCaption -match "Windows Server 2016")
                            {
								[void]$OSDirectories.Add("Windows 2016")
                            }
                            elseif($OSCaption -match "Windows 7")
                            {
								[void]$OSDirectories.Add("Windows7")
                            }
                            elseif($OSCaption -match "Windows 10")
                            {
								[void]$OSDirectories.Add("Windows 10")
                            }

                                #$article.Replace("KB","")
                                #$qry="Select * from CCM_UpdateStatus WHERE Status = 'Missing' And Article Like '$Article' And Bulletin LIKE '$Bulletin'"
								$qry="Select * from CCM_UpdateStatus WHERE Status = 'Missing'"
								if($Article)
								{
									$Article = $Article.Replace("KB","")
									$qry="$qry And Article Like '%$Article%'"
									[void]$OSDirectories.Add("Office 2010 32-Bit")
									[void]$OSDirectories.Add("SQL 2008 R2")
									[void]$OSDirectories.Add("SQL 2016")
								}
								if($Bulletin)
								{
									$qry="$qry And Bulletin LIKE '%$Bulletin%'"
									[void]$OSDirectories.Add("Office 2010 32-Bit")
									[void]$OSDirectories.Add("SQL 2008 R2")
									[void]$OSDirectories.Add("SQL 2016")
								}

								try
								{
									$result=Get-WmiObject -ComputerName $($Computer) -Namespace ROOT\ccm\SoftwareUpdates\UpdatesStore -Query $qry -ErrorAction Stop
									#$result=Get-WmiObject -ComputerName $env:COMPUTERNAME -Namespace ROOT\ccm\SoftwareUpdates\UpdatesStore -Query $qry -ErrorAction Stop
								}
								catch [System.Management.ManagementException]
								{
									Write-Warning "You should run this script with elevated rights..exiting!"
									break
								}
								catch
								{
									Write-Warning -Message $Error[0].Exception.Message
									Write-Warning -Message "exiting script"
									break
								}
                                
								#$result
								if($IncludeOfficeUpdates)
								{
									[void]$OSDirectories.Add("Office 2010 32-Bit")
								}
								if($IncludeSQLUpdates)
								{
									[void]$OSDirectories.Add("SQL 2008 R2")
									[void]$OSDirectories.Add("SQL 2016")
                                }
                                
                                foreach($OSDirectory in $OSDirectories)
                                {
                                    foreach($item in $result)
                                    {
                                        $ppath=(Resolve-Path "\\srv-sccm02\Packages$\Updates\$($OSDirectory)\$($item.UniqueId)\*.cab","\\srv-sccm02\Packages$\Updates\$($OSDirectory)\$($item.UniqueId)\*.exe" -ErrorAction 0).ProviderPath
                                        if($ppath)
                                        {
                                            Add-Member -InputObject $item -MemberType NoteProperty -Name FilePath -Value $ppath -Force
                                        }
                                    }

                                }

                                #productid zie.... https://msdn.microsoft.com/en-us/library/windows/desktop/ff357803(v=vs.85).aspx
                                #$result
                                $output=$result | Where-Object {(-not ([System.String]::IsNullOrEmpty($_.FilePath))) -and ($_.ProductID -ne "28bc880e-0592-4cbf-8f95-c79b17911d5f")} | Select-Object PSComputerName,Article,Bulletin,ProductID,Title,UniqueId,FilePath
                                #$output = $output -notmatch '%'
 
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

			####
			# you can add other parameters and they should correspond with the parameters defined in the $ScriptBlock
			#$PowershellThread.AddParameter("AnotherParameter", $AnotherParameter) | out-null
			#param
			#	(
			#		[String]
			#		$Computer,
			#		$AnotherParameter
			#	)
			####

			if ($MultiThread)
			{
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
				$PowershellThread.AddParameter("Computer", $Computer) | out-null
                $PowershellThread.AddParameter("Article", $Article) | out-null
                $PowershellThread.AddParameter("Bulletin", $Bulletin) | out-null
                $PowershellThread.AddParameter("IncludeOfficeUpdates", $IncludeOfficeUpdates) | out-null
                $PowershellThread.AddParameter("IncludeSQLUpdates", $IncludeSQLUpdates) | out-null
                
                
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
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Article,$Bulletin,$SetLabeledUri,$ClearLabeledUriBefore,$IncludeOfficeUpdates,$IncludeSQLUpdates,$Verbose
				}
				# for each parameter in the scriptblock add the same argument to the argumentlist
				else
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Article,$Bulletin,$SetLabeledUri,$ClearLabeledUriBefore,$IncludeOfficeUpdates,$IncludeSQLUpdates
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