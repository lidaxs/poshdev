<#
    version 1.0.0.2
    added class errorobject
    added processerror method
    version 1.0.0.1
    added processinfo method

    version 1.0.0.0
    initial creation
#>
Class CPUSingle
{
    # Properties
    $AddressWidth
    $Architecture
    $AssetTag
    $Availability
    $Caption
    $Characteristics
    $ConfigManagerErrorCode
    $ConfigManagerUserConfig
    $CpuStatus
    $CreationClassName
    $CurrentClockSpeed
    $CurrentVoltage
    $DataWidth
    $Description
    $DeviceID
    $ErrorCleared
    $ErrorDescription
    $ExtClock
    $Family
    $InstallDate
    $L2CacheSize
    $L2CacheSpeed
    $L3CacheSize
    $L3CacheSpeed
    $LastErrorCode
    $Level
    $LoadPercentage
    $Manufacturer
    $MaxClockSpeed
    $Name
    $NumberOfCores
    $NumberOfEnabledCore
    $NumberOfLogicalProcessors
    $OtherFamilyDescription
    $PartNumber
    $PNPDeviceID
    $PowerManagementCapabilities
    $PowerManagementSupported
    $ProcessorId
    $ProcessorType
    $Revision
    $Role
    $SecondLevelAddressTranslationExtensions
    $SerialNumber
    $SocketDesignation
    $Status
    $StatusInfo
    $Stepping
    $SystemCreationClassName
    $SystemName
    $ThreadCount
    $UniqueId
    $UpgradeMethod
    $Version
    $VirtualizationFirmwareEnabled
    $VMMonitorModeExtensions
    $VoltageCaps

    CPUSingle([System.Management.ManagementObject]$cpuwmiobject) : Base(){
        foreach($item in $cpuwmiobject.Properties)
        {
            $this."$($item.Name)"=$item.Value
        }

        try
        {
            if ($this.InstallDate) {
                $this.InstallDate = [System.Management.ManagementDateTimeConverter]::ToDateTime($this.InstallDate)
            }
            
        }
        catch
        {
            
        }
    }
}

Class CPUInfo
{
    $WMIObject
    [System.String]$SystemName
    [System.Int16]$NumberOfCPU
    [System.Int16]$TotalL2CacheSize
    [System.Int16]$TotalL3CacheSize
    [System.Int16]$TotalNumberOfCores
    [System.Int16]$AverageClockSpeed
    [System.String]$LoadPercentages
    Hidden [Int]$ecounter

    CPUInfo() : Base(){
        # map values to the class properties
        $this.wmiobject = Get-WmiObject -Class Win32_Processor -ComputerName $env:COMPUTERNAME -ErrorAction Stop
        $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        $this.ProcessInfo()
    }
    CPUInfo([String]$ComputerName) : Base(){
        # map values to the class properties
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_Processor -ComputerName $ComputerName -ErrorAction Stop
            $this.ProcessInfo()
        }
        catch {
            $this.ProcessError()
            #[CPUInfo]::ProcessErrorEx()
        }
    }

    Hidden ProcessInfo()
    {
        foreach ($item in $this.wmiobject)
        {
            Add-Member -InputObject $this -MemberType NoteProperty -Name $item.DeviceID -Value ([CPUSingle]::new($item))
            $totalclockspeed         += $item.CurrentClockSpeed
            $load+="($($item.DeviceID) : $($item.LoadPercentage)%)"
            $this.NumberOfCPU        += 1
            $this.AverageClockSpeed   = $totalclockspeed/$this.NumberOfCPU
            $this.SystemName          = $item.SystemName
            $this.TotalNumberOfCores += $item.NumberOfCores
            $this.TotalL2CacheSize   += $item.L2CacheSize
            $this.TotalL3CacheSize   += $item.L3CacheSize
            $this.LoadPercentages     = $load
        }
    }

    ProcessError()
    {
        $this.ecounter+=1
        Add-Member -InputObject $this -MemberType NoteProperty -Name ErrorObject$($this.ecounter) -Value ([ErrorObject]::New($Error[0]))
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