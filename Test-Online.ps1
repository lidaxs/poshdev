<#
    version 1.0.6
    Renamed function scan-port to test-ports

    version 1.0.5
    added function Test-Port
    added function Scan-Port
    with scanport capabilities

    version 1.0.4
    replaced PrimaryAddressResolutionStatus with StatusCode
    added verbosing in pipeline output

    version 1.0.3
    When used in pipeline only returns $true instead of pscustom object

    version 1.0.2
    changed ipaddress -like to a regex pattern

    version 1.0.1
    replaced StatusCode with StatusCode
    Changed parameter -Filter to -Query and removed -Class parameter

    version 1.0.0 initial staging

    wishlist
    Only return true when using the pipeline..done
#>

function Test-Online
{
    param
    (
        # make parameter pipeline-aware
        [Parameter(Mandatory,ValueFromPipeline)]
        [string[]]
        $ComputerName,
 
        $TimeoutMillisec = 1000
    )
 
    begin
    {
        # use this to collect computer names that were sent via pipeline
        [Collections.ArrayList]$bucket = $input
        # hash table with error code to text translation
        $StatusCode_ReturnValue =
        @{
            0='Success'
            11001='Buffer Too Small'
            11002='Destination Net Unreachable'
            11003='Destination Host Unreachable'
            11004='Destination Protocol Unreachable'
            11005='Destination Port Unreachable'
            11006='No Resources'
            11007='Bad Option'
            11008='Hardware Error'
            11009='Packet Too Big'
            11010='Request Timed Out'
            11011='Bad Request'
            11012='Bad Route'
            11013='TimeToLive Expired Transit'
            11014='TimeToLive Expired Reassembly'
            11015='Parameter Problem'
            11016='Source Quench'
            11017='Option Too Big'
            11018='Bad Destination'
            11032='Negotiating IPSEC'
            11050='General Failure'
        }
    
    
        # hash table with calculated property that translates
        # numeric return value into friendly text
 
        $statusFriendlyText = @{
            # name of column
            Name = 'Status'
            # code to calculate content of column
            Expression = { 
                # take status code and use it as index into
                # the hash table with friendly names
                # make sure the key is of same data type (int)
                $StatusCode_ReturnValue[([int]$_.StatusCode)]
            }
        }
 
        # calculated property that returns $true when status -eq 0
        $IsOnline = @{
            Name = 'Online'
            Expression = { $_.StatusCode -eq 0 }
        }
 
        # do DNS resolution when system responds to ping
        $ippattern = "\A(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\z"
        $DNSName = @{
            Name = 'DNSName'
            Expression = { if ($_.StatusCode -eq 0) { 
                    if ([regex]::IsMatch($($_.Address),$ippattern))
                    {                    
                        Write-Host "resolving $($_.Address)"
                         [Net.DNS]::GetHostByAddress($_.Address).HostName
                    } 
                    else
                    {
                        [Net.DNS]::GetHostByName($_.Address).HostName
                    } 
                }
            }
        }
    }
    
    process
    {
        if($PSCmdlet.MyInvocation.ExpectingInput){
            if ((Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$_') and timeout=$TimeoutMillisec").StatusCode -eq 0) {
                return $true
            }
            else {
                Write-Verbose "$($_) not online!"
            }
        }
        else {
            $query = $ComputerName -join "' or Address='"
            Get-WmiObject -Query "Select * From Win32_PingStatus Where (Address='$query') and timeout=$TimeoutMillisec" |
            Select-Object -Property Address, $IsOnline, $DNSName, $statusFriendlyText
        }
    }
    
    end
    {
    }
    
} 

function Test-Port
{
        # make parameter pipeline-aware
    param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

        [Int]
        $TimeoutMS = 1000,

        [Int]
        $portNumber=139,

        [Int]
        $Count = 0
    )

    begin
    {
        $Script:exitContext = $false
    }

    process
    {
        
        ForEach($Computer in $ClientName)
        {
            if( -not ([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,$portNumber).AsyncWaitHandle.WaitOne($TimeoutMS,$exitContext)))
            {
                #retry
                for ($i = 0; $i -lt $Count; $i++)
                { 
                    Write-Host "retrying tcpconnection to $Computer on $portNumber"
                    if([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,$portNumber).AsyncWaitHandle.WaitOne($TimeoutMS,$exitContext))
                    {
                        Write-Output $true
                    }
        
                }
        
                Write-Output $false
        
            }
            else
            {
                Write-Output $true
            }
        }
    }

    end
    {

    }
}


function Test-Ports
{
	[CmdletBinding()]
	[OutputType([System.Object])]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),
		
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

	}
	process 
	{
		foreach($Computer in $ClientName)
		{$note=New-Object PSObject | Select-Object ComputerName,RemoteIP
			Write-Verbose "Scanning host $Computer"
			ForEach ( $Port in $Ports )
            {
                #$note=New-Object PSObject | Select-Object ComputerName,TTL,RemoteIP
                Add-Member -InputObject $note -MemberType NoteProperty -Name $Port -Value "NA"
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
                           $note.$Port = "Timeout"
                           #$note.$Port = $false
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
					
					$note.ComputerName=$Computer
					#$note.Port=$Port
					
                    if ($tcp.client.connected)
                    {
                        #$note.PortOpen=$True
                        $note.$Port = $true
						#$note.TTL=$($tcp.client.ttl)
						[string]$rep=$tcp.client.RemoteEndPoint
						[string]$ip=$rep.substring(0,$rep.indexof(":"))
		           		$note.RemoteIP=$ip
					}
					else{
                        #$note.PortOpen=$False
                        
						#$note.TTL="No TTL"
						$note.RemoteIP="Unknown"
					}	
					#$note
					#dispose and disconnect
					$tcp.close()
                }
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

            $note

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