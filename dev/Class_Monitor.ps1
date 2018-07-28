<#
    version 1.0.0.1
    added ProcessError method
    added ErrorObject class
    
    version 1.0.0.0
    Initial creation
#>
Class Monitor{
    [System.String]$Manufacturer
    [System.String]$ProductCode
    [System.String]$SerialNumber
    [System.String]$DeviceFriendlyName
    [System.String]$Age
    [System.Management.ManagementObject]$wmiobject
 
    Monitor($wmimonitorobject) : Base()
    {
        $this.wmiobject = $wmimonitorobject
        foreach($property in $wmimonitorobject.Properties)
        {
            Add-Member -InputObject $this -MemberType NoteProperty -Name $property.Name -Value $property.Value
            if($property.Value -is [UInt16[]]){
                Add-Member -InputObject $this.$($property.Name) -MemberType ScriptMethod -Name ByteArrayToString -Value {$temp = [System.Text.Encoding]::ASCII.GetString($($this));$temp.Replace("`0","")} -Force
            }
        }
        $this.GetAge()
        $this.Manufacturer = $this.ManufacturerName.ByteArrayToString()
        $this.ProductCode = $this.ProductCodeID.ByteArrayToString()
        $this.DeviceFriendlyName = $this.UserFriendlyName.ByteArrayToString()
        $this.SerialNumber = $this.SerialNumberID.ByteArrayToString()
    }
    GetAge()
    {
        $manufacturerdate=(Get-Date -Date 1-1-$($this.YearOfManufacture)).AddDays(7*$($this.WeekOfManufacture))
        $today = [datetime]::Today
        $this.Age = $this.TimeSpanToYMD($today - $manufacturerdate)
    }
    [String]TimeSpanToYMD([System.TimeSpan]$TimeSpan)
    {   $totaldays = $TimeSpan.TotalDays
        $years = [System.Math]::Floor($($totaldays)/365)
        $totaldays%=365
        $months=[System.Math]::Floor($($totaldays)/30)
        $totaldays%=30
        return "$years year(s), $($months) months and $totaldays days"
    }
}

Class MonitorInfo
{
    $wmiobject
    Hidden [Int]$ecounter
    MonitorInfo() : Base()
    {
        $this.wmiobject = Get-WmiObject -Class WMIMonitorID -Namespace root\wmi -ComputerName $env:COMPUTERNAME
        $this.ProcessInfo()
    }
    MonitorInfo([String]$ComputerName) : Base()
    {
        $this.wmiobject = Get-WmiObject -Class WMIMonitorID -Namespace root\wmi -ComputerName $ComputerName
        $this.ProcessInfo()
    }
    ProcessInfo()
    {
        foreach ($item in $this.wmiobject) {
            $counter+=1
            Add-Member -InputObject $this -MemberType NoteProperty -Name "Monitor$counter" -Value ([Monitor]::new($item))
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