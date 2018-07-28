<#
	version 1.0.0.3
	removed statement remove-variable
	was not necessary after all

	version 1.0.0.2
	removing variables in endblock

	version 1.0.0.1
	added some predefined exclusion

	version 1.0.0.0
	initial upload

    wishlist foldersize/acl's
#>
function Get-FolderHash {
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
			PS C:\> Get-FolderHash -ClientName C120VMXP,C120WIN7

		.EXAMPLE
			PS C:\> $mycollection | Get-FolderHash

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
		[Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[Alias("Directory","Map","DirectoryPath","Folder","FolderPath")]
		[ValidateNotNullOrEmpty()]
		$Path,

        [String[]]
        $Exclude = @('*.xml','*.log','*dd.ini','GenerateEzisStructure*'),

		# run the script multithreaded against multiple targets
		[Parameter(Mandatory=$false)]
		[Switch]
		$MultiThread,

		# maximum number of threads that can run simultaniously
		[Parameter(Mandatory=$false)]
		[Int]
		$MaxThreads=4,

		# Maximum time(seconds) in which a thread must finish before a timeout occurs
		[Parameter(Mandatory=$false)]
		[Int]
		$MaxResultTime=120,

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
		ForEach($iPath in $Path)
		{

			if($PSCmdlet.ShouldProcess("$iPath", "Get-FolderHash"))
			{

				$ScriptBlock=
				{[CmdletBinding(SupportsShouldProcess=$true)]
				param
				(
					$iPath,

                    $Exclude
				)
					# Test path
					if ( [System.IO.Directory]::Exists("$iPath"))
					{
						#the code to execute in each thread
						try
						{
                            $files = Get-ChildItem $iPath -Exclude $Exclude -Recurse | Where-Object { -not $_.psiscontainer }

                            $allBytes = new-object System.Collections.Generic.List[byte]
                            foreach ($file in $files)
                            {
                                $allBytes.AddRange([System.IO.File]::ReadAllBytes($file.FullName))
                                $allBytes.AddRange([System.Text.Encoding]::UTF8.GetBytes($file.Name))
                            }
                            $hasher         = [System.Security.Cryptography.MD5]::Create()
                            $calculatedHash = [string]::Join("",$($hasher.ComputeHash($allBytes.ToArray()) | ForEach-Object {"{0:x2}" -f $_}))

                            [PSCustomObject]$output = "" | Select-Object Path,Hash,Exclusions
                            $output.Path     = $iPath
                            $output.Hash     = $calculatedHash
                            $output.Exclusions = $Exclude
                            $output
						}
						catch
						{
                            $Error[0].Exception.Message
						}
					} # end if test-path
					else # path exists
					{
						Write-Warning "$iPath does not exist or path not reachable!"
					}
				} # end scriptblock


			} # end if $PSCmdlet.ShouldProcess


			if ($MultiThread)
			{
				$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
				$PowershellThread.AddParameter("iPath", $iPath) | out-null
                $PowershellThread.AddParameter("Exclude", $Exclude) | out-null
				if($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
				{
					$PowershellThread.AddParameter("Verbose") | out-null
				}
				$PowershellThread.RunspacePool = $RunspacePool

				$Handle = $PowershellThread.BeginInvoke()
				$Job = "" | Select-Object Handle, Thread, object
				$Job.Handle = $Handle
				$Job.Thread = $PowershellThread
				$Job.Object = $iPath.ToString()
				$Jobs += $Job
			}
			else # $MultiThread
			{
				if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('verbose'))
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $iPath,$Exclude,$Verbose
				}
				# for each parameter in the scriptblock add the same argument to the argumentlist
				else
				{
					Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $iPath,$Exclude
				}
			}

		} # end foreach $computer

	} # end processblock

	end
		{
			if($MultiThread)
			{

			$ResultTimer = Get-Date
			While (@($Jobs | Where-Object {$Null -ne $_.Handle}).count -gt 0)
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

} # end function