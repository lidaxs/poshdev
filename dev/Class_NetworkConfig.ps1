<#
    version 1.0.0.2
    added MACToOID
    added Class ErrorObject
    added ProcessError method
    
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
    [System.String]$MACToOID
    [System.Int16]$NumberOfInterfaces
    Hidden [System.Int16]$ecounter
    #endregion properties

    NetworkConfigInfo() : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $env:COMPUTERNAME -ErrorAction Stop
            $this.ProcessInfo()
        }
        catch {
            $this.ProcessError()
        }
    }

    NetworkConfigInfo([System.Boolean]$OnlyActiveAdapters) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $env:COMPUTERNAME -Filter "IPEnabled=True and DNSDomain='antoniuszorggroep.local'"  -ErrorAction Stop
            $this.ProcessInfo()
        }
        catch {
            $this.ProcessError()
        }
    }

    NetworkConfigInfo([String]$ComputerName) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -ErrorAction Stop
            $this.ProcessInfo()
        }
        catch {
            $this.ProcessError()
        }
    }

    NetworkConfigInfo([String]$ComputerName,[System.Boolean]$OnlyActiveAdapters) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $ComputerName -Filter "IPEnabled=True and DNSDomain='antoniuszorggroep.local'" -ErrorAction Stop
            $this.ProcessInfo()
        }
        catch {
            $this.ProcessError()
        }
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
        $this.MACAddressToOID()
        try{
            $this.DHCPLeaseExpires = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.DHCPLeaseExpires)
        }
        catch{
            $this.ProcessError()
        }
        try {
            $this.DHCPLeaseObtained = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.DHCPLeaseObtained)
        }
        catch{
            $this.ProcessError()
        }
        
    }

    ProcessError()
    {
        $this.ecounter+=1
        Add-Member -InputObject $this -MemberType NoteProperty -Name ErrorObject$($this.ecounter) -Value ([ErrorObject]::New($Error[0]))
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

    [String]MACAddressToOID()
    {
        try {
            $MAC=$this.MACAddress.Replace("-",":")
            $arrHexToInt=$MAC.Split(":") | ForEach-Object {[int]"0x$_"}
            $this.MACToOID = $arrHexToInt -join "."
        }
        catch {
            $this.ProcessError()
        }

        return $this.MACToOID

    }
}

Class ErrorObject{
    $Exception
    $Reason
    $Target
    $Script
    $Line
    $Column
 
    ErrorObject([Management.Automation.ErrorRecord]$objError) : Base()
    {
        $this.Exception = $objError.Exception.Message
        $this.Reason    = $objError.CategoryInfo.Reason
        $this.Target    = $objError.CategoryInfo.TargetName
        $this.Script    = $objError.InvocationInfo.ScriptName
        $this.Line      = $objError.InvocationInfo.ScriptLineNumber
        $this.Column    = $objError.InvocationInfo.OffsetInLine
    }
 }