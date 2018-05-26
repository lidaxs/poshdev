<#
version 1.0.0.1
added processinfo

version 1.0.0.0
initial creation
#>
class NetworkConfig {
    #region properties
    $ArpAlwaysSourceRoute
    $ArpUseEtherSNAP
    $Caption
    $DatabasePath
    $DeadGWDetectEnabled
    $DefaultIPGateway
    $DefaultTOS
    $DefaultTTL
    $Description
    $DHCPEnabled
    $DHCPLeaseExpires
    $DHCPLeaseObtained
    $DHCPServer
    $DNSDomain
    $DNSDomainSuffixSearchOrder
    $DNSEnabledForWINSResolution
    $DNSHostName
    $DNSServerSearchOrder
    $DomainDNSRegistrationEnabled
    $ForwardBufferMemory
    $FullDNSRegistrationEnabled
    $GatewayCostMetric
    $IGMPLevel
    $Index
    $InterfaceIndex
    $IPAddress
    $IPConnectionMetric
    $IPEnabled
    $IPFilterSecurityEnabled
    $IPPortSecurityEnabled
    $IPSecPermitIPProtocols
    $IPSecPermitTCPPorts
    $IPSecPermitUDPPorts
    $IPSubnet
    $IPUseZeroBroadcast
    $IPXAddress
    $IPXEnabled
    $IPXFrameType
    $IPXMediaType
    $IPXNetworkNumber
    $IPXVirtualNetNumber
    $KeepAliveInterval
    $KeepAliveTime
    $MACAddress
    $MTU
    $NumForwardPackets
    $PMTUBHDetectEnabled
    $PMTUDiscoveryEnabled
    $ServiceName
    $SettingID
    $TcpipNetbiosOptions
    $TcpMaxConnectRetransmissions
    $TcpMaxDataRetransmissions
    $TcpNumConnections
    $TcpUseRFC1122UrgentPointer
    $TcpWindowSize
    $WINSEnableLMHostsLookup
    $WINSHostLookupFile
    $WINSPrimaryServer
    $WINSScopeID
    $WINSSecondaryServer
    $ComputerName
    #endregion properties

    NetworkConfig([System.Management.ManagementObject]$wmi_NetworkAdapterConfiguration)
    {
        foreach($item in $wmi_NetworkAdapterConfiguration.Properties)
        {
            $this."$($item.Name)"=$item.Value
        }

        try
        {
            if ($this.DHCPLeaseExpires) {
                $this.DHCPLeaseExpires  = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.DHCPLeaseExpires)
                $this.DHCPLeaseObtained = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.DHCPLeaseObtained)
            }
            else {
                $this.DHCPLeaseExpires  = 'Unknown'
                $this.DHCPLeaseObtained = 'Unknown'
            }
        }
        catch
        {
            
        }
    }
}

class NetworkConfigInfo {
    #region properties
    $WMIObject
    $ComputerName
    [System.String]$VisioIP
    [System.Int16]$NumberOfInterfaces
    $ErrorObject = 'no errors'
    #endregion properties

    NetworkConfigInfo() : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $env:COMPUTERNAME
        }
        catch {
            $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        }
    
        $this.ProcessInfo()

    }
    NetworkConfigInfo([System.Boolean]$OnlyActiveAdapters) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $env:COMPUTERNAME -Filter "IPEnabled=True and DNSDomain='antoniuszorggroep.local'"
        }
        catch {
            $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        }
    
        $this.ProcessInfo()

    }
    NetworkConfigInfo([String]$ComputerName) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName
        }
        catch {
            $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        }
    
        $this.ProcessInfo()

    }
    NetworkConfigInfo([String]$ComputerName,[System.Boolean]$OnlyActiveAdapters) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled=True and DNSDomain='antoniuszorggroep.local'"
        }
        catch {
            $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        }
    
        $this.ProcessInfo()

    }
    Hidden ProcessInfo()
    {
        foreach ($item in $this.wmiobject) {
            Add-Member -InputObject $this -MemberType NoteProperty -Name $item.Caption -Value ([NetworkConfig]::new($item))
            $this.NumberOfInterfaces+=1
            $this.ComputerName = $item.DNSHostName
            if($item.DNSDomain -eq 'antoniuszorggroep.local')
            {
                foreach ($prop in $item.properties)
                {
                    Add-Member -InputObject $this -MemberType NoteProperty -Name $prop.Name -Value $prop.value
                }
            }
        }

        $this.TransformIPToVisio()
        $this.DHCPLeaseExpires = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.DHCPLeaseExpires)
        $this.DHCPLeaseObtained = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.DHCPLeaseObtained)
    }

    TransformToVisioList($inputlist)
    {
        [System.Collections.ArrayList]$visiolist = $inputlist
        
        if($visiolist)
        {
            #$visiolist.Sort()
            $this.VisioIP = $visiolist -join ";"
        }
        else
        {
            $this.VisioIP = $null
        }
        
    }

    [String]TransformIPToVisio()
    {
        $this.TransformToVisioList($this.IPAddress)
        return $this.VisioIP
    }
}