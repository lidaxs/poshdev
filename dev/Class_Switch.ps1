# only managementdesktop
. \\srv-fs01\users\adm-bouweh01\Appz\GIT\Class_NetworkConfig.ps1

# list of switches
$switches=Get-Content \\srv-fs01\scripts$\switch_name2.txt
$WS="C0300180"
$WS="C0300008"
$WS="C1203200"
$WS="C0300011"
$mac="00:23:7D:BF:25:A0"
$arrSwitches=@("10.20.5.1","10.20.5.2","10.20.5.3","10.20.4.1","10.20.4.2","10.20.4.3","10.20.4.5")
[System.Collections.ArrayList]$arrSwitches=Get-Content -Path '\\srv-fs01\Scripts$\switch_name2.txt' | Foreach-Object {[System.Net.DNS]::GetHostByName($_).addresslist.ipaddresstostring}

#$arrSwitches = $arrSwitches[2..15]

# bpn
# oid to mac_reverse lookup mac in sccm

# load sharpsnmplib.dll
[void][reflection.assembly]::LoadFrom( "\\srv-fs01\scripts$\BD_SNMP\SharpSnmpLib.dll" )


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

Class ErrorObject {
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



Class Switch
{

}

Class SwitchInfo
{

}

Class SwitchComputer
{
    #class for found computerconnection on switch
    $AdminStatus
    [Lextm.SharpSnmpLib.ObjectIdentifier]$BaseOID = ".1.3.6.1.2.1.17.4.3.1.2"
    [Lextm.SharpSnmpLib.ObjectIdentifier]$MACOID
    [System.Collections.Generic.List[Lextm.SharpSnmpLib.Variable]]$OIDList=@()
    [System.Int32]$BPN
    $Community="public"
    [System.Net.IPEndPoint]$EndPoint
    $NetConfig
    $SNMPData
    [System.Net.IPAddress]$SysIP
    $SysName
    $SysUptime
    $SysDescription
    $SysLocation
    [int]$TimeOut = 1000
    [int]$Port = 161
    $IFDescription
    $IFPort
    $IFOperationalStatus
    $SpeedInMB
    $ComputerName
    $ComputerMACAddress
    $ComputerIPAddress
    $InError
    $OutError
    $LastChange
    $SNMPVersion = [Lextm.SharpSnmpLib.VersionCode]::V2

    SwitchComputer() : Base()
    {
        $this.NetConfig = [NetworkConfigInfo]::new()
        $this.MACOID=[Lextm.SharpSnmpLib.ObjectIdentifier]::new("$($this.BaseOID).$($this.NetConfig.MACToOID)")
        $this.OIDList.Add($this.MACOID)
        $this.ComputerName = $env:COMPUTERNAME
        $this.ProcessInfo()

    }
    SwitchComputer([System.String]$ComputerName) : Base()
    {
        $this.NetConfig = [NetworkConfigInfo]::new("$ComputerName")
        $this.MACOID=[Lextm.SharpSnmpLib.ObjectIdentifier]::new("$($this.BaseOID).$($this.NetConfig.MACToOID)")
        $this.OIDList.Add($this.MACOID)
        $this.ComputerName = $ComputerName
        $this.ProcessInfo()
        
    }
    SwitchComputer([System.String]$ComputerName,$SwitchList) : Base()
    {
        
        $this.NetConfig = [NetworkConfigInfo]::new("$ComputerName")
        $this.MACOID=[Lextm.SharpSnmpLib.ObjectIdentifier]::new("$($this.BaseOID).$($this.NetConfig.MACToOID)")
        $this.OIDList.Add($this.MACOID)
        $this.ComputerName = $ComputerName
        
        foreach($Switch in $SwitchList)
        {
            $this.SysIP = $Switch
            #Write-Host "processing switchip $Switch"
            $this.ProcessInfo()
        }
        
    }
    GetBPN()
    {
        try
        {
            $this.BPN = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get($this.SNMPVersion, $this.EndPoint, $this.Community, $this.OIDList, $this.TimeOut).Data.ToInt32()
        }
        catch
        {

        }
    }
    GetSNMPData($OIDList)
    {
        $this.SNMPData = [Lextm.SharpSnmpLib.Messaging.Messenger]::Get($this.SNMPVersion, $this.EndPoint, $this.Community, $OIDList, $this.TimeOut)
    }
    ProcessInfo()
    {
        
        $this.ComputerIPAddress = $this.NetConfig.IPAddress
        $this.ComputerMACAddress = $this.NetConfig.MACAddress
        $this.EndPoint = [System.Net.IPEndPoint]::new($this.SysIP,$this.Port)
        
        $this.GetBPN()
        
        if($this.BPN -in 1..4096)
        {
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.1.5.0")                    #sysname
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.1.3.0")                    #sysuptime
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.1.1.0")                    #sysdesc
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.1.6.0")                    #sysloc
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.31.1.1.1.1.$($this.BPN)")  #ifport
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.31.1.1.1.18.$($this.BPN)") #devdesc?
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.5.$($this.BPN)")     #speed
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.8.$($this.BPN)")     #ifoperationalstatus (0=off;1=on)
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.14.$($this.BPN)")    #inerr
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.20.$($this.BPN)")    #outerr
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.9.$($this.BPN)")     #lastchange in timeticks
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.2.$($this.BPN)")     #ifdesc...convert??
            $this.OIDList.Add([Lextm.SharpSnmpLib.ObjectIdentifier]".1.3.6.1.2.1.2.2.1.7.$($this.BPN)")     #ifadminstatus
            $this.GetSNMPData($this.OIDList)
            $this.SysName   = $this.SNMPData[1].Data.ToString()
            $this.SysUptime = $this.SNMPData[2].Data.ToTimeSpan()
            $this.SysDescription = $this.SNMPData[3].Data.ToString()
            $this.SysLocation = $this.SNMPData[4].Data.ToString()
            $this.IFPort = $this.SNMPData[5].Data.ToString()
            $this.IFDescription = $this.SNMPData[6].Data.ToString()
            $this.SpeedInMB = [Math]::Round(($this.SNMPData[7].Data.ToUInt32())/1MB)
            $this.IFOperationalStatus = $this.SNMPData[8].Data.ToInt32()
            $this.InError = $this.SNMPData[9].Data.ToUInt32()
            $this.OutError = $this.SNMPData[10].Data.ToUInt32()
            $this.LastChange = $this.SNMPData[11].Data.ToTimeSpan()
            $this.AdminStatus = $this.SNMPData[13].Data.ToInt32()
            #$this.OIDList.Clear()
            break
        }
        else
        {
            $this=$null
        }
    }
}