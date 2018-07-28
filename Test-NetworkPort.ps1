<#
    version 1.0.0.0
    initial commit
#>
function Test-NetworkPort {
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Object])]
	param(
		[Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[ValidateNotNullOrEmpty()]
		[Alias("ComputerName","Host","Server")]
		[System.String[]]
		$ClientName = $env:COMPUTERNAME,
		
		[String[]]
		$Ports=("21","23","25","80","443","3389"),
		
		[Switch]
		$MultiThread,
		
		$MaxThreads=20,
		
		$MaxResultTime=500,
		
		$SleepTimer=1000
	)
	begin
	{
		#turn off error pipeline 
		$ErrorActionPreference = "SilentlyContinue"
		
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
		}
	}
	process 
	{
		foreach($Computer in $ClientName)
		{
			Write-Verbose "Scanning host $Computer"
			ForEach ( $Port in $Ports )
			{
				Write-Verbose "Scanning Port $Port on $Computer"
				$Scriptblock=
				{
					param($Computer,$Port)
					try{
						$tcp=New-Object System.Net.Sockets.TcpClient
						#($Computer, $Port)
						$connect = $tcp.BeginConnect($Computer,$Port,$null,$null)
						$Wait = $connect.AsyncWaitHandle.WaitOne(1000,$false)
						
						If (-Not $Wait)
						{
						   Write-Verbose "Timeout connecting to $($Computer) : $($Port)"
						}
						Else
						{
						    $Error.clear()
						    $tcp.EndConnect($connect)
						}

					}
					catch [System.Management.Automation.RuntimeException]{
						Write-Warning "RuntimeException connecting to $($Computer):$($Port)"
					}
					catch [System.TimeoutException]{
						Write-Warning "TimeOutException connecting to $($Computer):$($Port)"
					}
					catch 
					{
						Write-Warning "Could not connect to $($Computer):$($Port)"
					}
					$note=New-Object PSObject | Select-Object ComputerName,Port,PortOpen,TTL,RemoteIP
					$note.ComputerName=$Computer
					$note.Port=$Port
					
					if ($tcp.client.connected) {
						$note.PortOpen=$True
						$note.TTL=$($tcp.client.ttl)
						[string]$rep=$tcp.client.RemoteEndPoint
						[string]$ip=$rep.substring(0,$rep.indexof(":"))
		           		$note.RemoteIP=$ip
					}
					else{
						$note.PortOpen=$False
						$note.TTL="No TTL"
						$note.RemoteIP="Unknown"
					}	
					$note
					#dispose and disconnect
					$tcp.close()
				}

				if ($MultiThread)
				{
					$PowershellThread = [powershell]::Create().AddScript($ScriptBlock)
					$PowershellThread.AddParameter("Computer", $Computer) | out-null
					$PowershellThread.AddParameter("Port", $Port) | out-null
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
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Port,$Verbose
					}
					# for each parameter in the scriptblock add the same argument to the argumentlist
					else
					{
						Invoke-Command -ScriptBlock $ScriptBlock -ArgumentList $Computer,$Port
					}
				}
			}
		}
	}
	end
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
               Write-Warning "Child script appears to be frozen, try increasing MaxResultTime...CTRL + C to abort operation"
	        }
	        Start-Sleep -Milliseconds $SleepTimer
	    } 
	    $RunspacePool.Close() | Out-Null
	    $RunspacePool.Dispose() | Out-Null
	} # end endblock
}