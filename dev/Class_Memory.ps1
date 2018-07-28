<#
    version 1.0.0.0
    initial creation
#>

Class Memory{
    #region properties
    $PSComputerName
    $Attributes
    $BankLabel
    $Capacity
    $Caption
    $ConfiguredClockSpeed
    $ConfiguredVoltage
    $CreationClassName
    $DataWidth
    $Description
    $DeviceLocator
    $FormFactor
    $HotSwappable
    $InstallDate
    $InterleaveDataDepth
    $InterleavePosition
    $Manufacturer
    $MaxVoltage
    $MemoryInGB
    $MemoryType
    $MinVoltage
    $Model
    $Name
    $OtherIdentifyingInfo
    $PartNumber
    $PositionInRow
    $PoweredOn
    $Removable
    $Replaceable
    $SerialNumber
    $SKU
    $SMBIOSMemoryType
    $Speed
    $Status
    $Tag
    $TotalWidth
    $TypeDetail
    $Version
    #endregion properties

    Memory([System.Management.ManagementObject]$memwmiobject) : Base()
    {
        foreach($item in $memwmiobject.Properties)
        {
            $this."$($item.Name)"=$item.Value
        }

        $this.MemoryInGB = $this.Capacity/1GB

    }
}

Class MemoryInfo{

    #region properties
    $ComputerName
    $WMIObject
    [System.Decimal]$TotalMemory = 0
    [System.Int16]$TotalMemoryInGB = 0
    $ErrorObject = 'no errors'
    #endregion properties

    MemoryInfo() : Base()
    {
        # map values to the class properties
        try {
            $this.ComputerName = $env:COMPUTERNAME
            $this.wmiobject = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $env:COMPUTERNAME -ErrorAction Stop
        }
        catch {
            $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        }

        $this.ProcessInfo()
    }

    MemoryInfo([String]$ComputerName) : Base()
    {
        # map values to the class properties
        $this.ComputerName = $ComputerName
        try {
            $this.wmiobject = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $ComputerName -ErrorAction  Stop
        }
        catch {
            $this.ErrorObject = ConvertFrom-ErrorRecord -Record $error[0]
        }

        $this.ProcessInfo()
    }

    Hidden ProcessInfo()
    {
        foreach ($item in $this.wmiobject)
        {
            Add-Member -InputObject $this -MemberType NoteProperty -Name $item.DeviceLocator -Value ([Memory]::new($item))
            $this.TotalMemory += $item.Capacity
        }

        $this.TotalMemoryInGB = $this.TotalMemory/1GB

    }
}