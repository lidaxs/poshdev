function Test-Online
{
<#
    version 1.0.4
    replaced PrimaryAddressResolutionStatus with StatusCode
    added verbosing in pipeline output

    version 1.0.3
    When used in pipeline only returns $true instead of pscutom object

    version 1.0.2
    changed ipaddress -like to a regex pattern

    version 1.0.1
    replaced StatusCode with StatusCode
    Changed parameter -Filter to -Query and removed -Class parameter

    version 1.0.0 initial staging

    wishlist
    Only return true when using the pipeline..done
#>
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
