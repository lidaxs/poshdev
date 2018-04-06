Class Reg
{
#region properties

$ComputerName

[ValidateSet('ClassesRoot','CurrentConfig','CurrentUser','DynData','LocalMachine','PerformanceData','Users')]
$Hive='LocalMachine'

$RemoteBaseKey

$RegValueHashTable

$Value

[ValidateSet('Unknown','String','ExpandString','Binary','DWord','MultiString','QWord','None')]
$ValueKind

#endregion properties

#region constructors
    Reg() : Base(){}

    Reg([String]$ComputerName) : Base()
    {

        if(-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)){
            break
        }
        elseif((Get-Service -Name RemoteRegistry -ComputerName $ComputerName).Status -ne 'Running'){
            try
            {
                Get-Service -Name RemoteRegistry -ComputerName $ComputerName | Start-Service
            }
            catch
            {
                Write-Host $Error[0].Exception.Message
                break
            }
        }
        $this.Hive          = 'localmachine'
        #$this.ComputerName  = $ComputerName
        $this.RemoteBaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]'localmachine',$ComputerName,[Microsoft.Win32.RegistryView]::Registry64)
    }

    Reg([String]$ComputerName,[String]$Hive) : Base()
    {
    $this.ComputerName  = $ComputerName
    Write-Verbose "Connecting to registry of $ComputerName"
        if(-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)){
            break
        }
        elseif((Get-Service -Name RemoteRegistry -ComputerName $ComputerName).Status -ne 'Running'){
            try
            {
                Get-Service -Name RemoteRegistry -ComputerName $ComputerName | Start-Service
            }
            catch
            {
                Write-Host $Error[0].Exception.Message
                break
            }
        }
        $this.Hive          = $Hive
        
        $this.RemoteBaseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$Hive,$ComputerName,[Microsoft.Win32.RegistryView]::Registry64)
    }

#endregion constructors

#region methods
    [Byte[]]ConvertStringToBinary([System.String]$String,[System.Boolean]$NullsInserted)
    {
        [System.Collections.Arraylist]$byteArray=@()
        $String.ToCharArray() | ForEach-Object -Process {
            [void]$byteArray.Add([byte][char]$_)
            if($NullsInserted)
            {
                [void]$byteArray.Add([byte][char]0)
            }
        }
        return $byteArray
    }

    [String]GetRegistryValue([String]$Key,[String]$ValueName)
    {
        try
        {
            $SubKey=$this.RemoteBaseKey.OpenSubKey($Key)
            $this.Value=$SubKey.GetValue($ValueName)
            $this.RegValueHashTable = @{
                Key       = $Key
                ValueName = $ValueName
                Value     = $this.Value
            }
        }
        catch
        {
            Write-Verbose "Could not open key $Key on $($this.ComputerName)([String]GetRegistryValue([String]`$Key,[String]`$ValueName))"
        }

        return $this.Value
    }

    [HashTable]ReturnRegAsHashTable([String]$Hive,[String]$Key,[String]$ValueName)
    {
        [Reg]::ConnectHive($Hive)
        try
        {
            $this.Value=$this.RemoteBaseKey.OpenSubKey($Key).GetValue($ValueName)
            $this.RegValueHashTable = @{
                Computer  = $env:COMPUTERNAME
                Key       = $Key
                ValueName = $ValueName
                Value     = $this.Value
            }
        }
        catch
        {
            Write-Warning "Could not open key $Key on $($this.ComputerName)([Hashtable]ReturnRegAsHashTable([String]`$Hive,[String]`$Key,[String]`$ValueName))"
        }

        return $this.RegValueHashTable      
    }

    [HashTable]ReturnRegAsHashTable([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName)
    {
        [Reg]::ConnectHive($ComputerName,$Hive)
        try{
        $this.Value=$this.RemoteBaseKey.OpenSubKey($Key).GetValue($ValueName)
        $this.RegValueHashTable = @{
            Computer  = $ComputerName
            Key       = $Key
            ValueName = $ValueName
            Value     = $this.Value
        }
        }
        catch
        {
            Write-Warning "Could not open key $Key on $($this.ComputerName)([Hashtable]ReturnRegAsHashTable([String]`$ComputerName,[String]`$Hive,[String]`$Key,[String]`$ValueName))"
        }

        return $this.RegValueHashTable      
    }

    SetValue([String]$Key,[String]$ValueName,$Value)
    {
        $this.SetValue($this.ComputerName,$this.Hive,$Key,$ValueName,$Value,'String')
    }

    SetValue([String]$Key,[String]$ValueName,$Value,[String]$ValueKind)
    {
        $this.SetValue($this.ComputerName,$this.Hive,$Key,$ValueName,$Value,$ValueKind)
    }

    SetValue([String]$Hive,[String]$Key,[String]$ValueName,$Value,$ValueKind)
    {
        $this.SetValue($this.ComputerName,$Hive,$Key,$ValueName,$Value,$ValueKind)
    }

    SetValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,$Value,[String]$ValueKind)
    {

        #$reghive = [Reg]::New($ComputerName,$Hive)
        #$this.RemoteBaseKey.
        try
        {
            $this.CreateSubKey($Key)
            $this.RemoteBaseKey.OpenSubKey($Key,$true).SetValue($ValueName,$Value,$ValueKind)
        }
        catch
        {
            Write-Warning "could not set registryvalue `'$Value`' of type `'$ValueKind`' on `'$ValueName`' in `'$Hive\$Key`' on `'$ComputerName`' (static [void]SetRegistryValue([String]`$ComputerName,[String]`$Hive,[String]`$Key,[String]`$ValueName,[String]`$Value,[String]`$ValueKind))"
        }
    }
    hidden CreateSubKey([String]$Key)
    {
        $this.RemoteBaseKey.CreateSubKey($Key)
    }

#endregion methods

#region static methods

    static [System.Object]ConnectHive($Hive)
    {
        return [Reg]::New($env:COMPUTERNAME,$Hive)
    }

    static [System.Object]ConnectHive($ComputerName,$Hive)
    {
        return [Reg]::New($ComputerName,$Hive)
    }

    static [void]DeleteSubKey([String]$ComputerName,[String]$Hive,[String]$Key)
    {
        $reghive = [Reg]::ConnectHive($ComputerName,$Hive)
        $myvalue = $null
        try
        {
            $myvalue = $reghive.RemoteBaseKey.DeleteSubKey($Key,$false)
        }
        catch
        {
            Write-Warning "Could not delete subkey `'$Key`' in `'$Hive`' on `'$ComputerName`' (static [void]DeleteSubKey([String]`$ComputerName,[String]`$Hive,[String]`$Key))"
        }
    }

    static [void]DeleteSubKeyTree([String]$ComputerName,[String]$Hive,[String]$Key)
    {
        $reghive = [Reg]::ConnectHive($ComputerName,$Hive)
        $myvalue = $null
        try
        {
            $myvalue=$reghive.RemoteBaseKey.DeleteSubKeyTree($Key,$false)
        }
        catch
        {
            Write-Warning "Could not delete subkeytree `'$Key`' in `'$Hive`' on `'$ComputerName`' (static [void]DeleteSubKeyTree([String]`$ComputerName,[String]`$Hive,[String]`$Key)"
        }
    }

    static [void]DeleteValue([String]$ComputerName,[String]$Hive,[String]$Key,$ValueName)
    {
        $reghive = [Reg]::ConnectHive($ComputerName,$Hive)
        #$mykey   = $null
        try
        {
            #$myvalue=$reghive.RemoteBaseKey.OpenSubKey($Key,$true).DeleteValue($ValueName)
            $reghive.RemoteBaseKey.OpenSubKey($Key,$true).DeleteValue($ValueName)
        }
        catch
        {
            Write-Warning "Could not delete value `'$ValueName`' in `'$Key`' on `'$ComputerName`' (static [void]DeleteValue([String]`$ComputerName,[String]`$Hive,[String]`$Key,`$ValueName)"
        }
    }

    static [String]GetRegistryValue([String]$Hive,[String]$Key,[String]$ValueName)
    {
        $reghive = [Reg]::ConnectHive($Hive)
        $myvalue = $null

        try{
            $myvalue=$reghive.RemoteBaseKey.OpenSubKey($Key,$false).GetValue($ValueName)
        }
        catch
        {
            Write-Verbose "Could not open key `'$Key`' in hive `'$Hive`' ('static [String]GetRegistryValue([String]`$Hive,[String]`$Key,[String]`$ValueName)')"
        }

        return $myvalue
    }

    static [String]GetRegistryValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName)
    {

        $reghive = [Reg]::ConnectHive($ComputerName,$Hive)
        $myvalue = $null

        try{
            $myvalue=$reghive.RemoteBaseKey.OpenSubKey($Key,$false).GetValue($ValueName)
        }
        catch
        {
            Write-Verbose "Could not open key `'$Key`' in hive `'$Hive`' on `'$ComputerName`' ('static [String]GetRegistryValue([String]`$ComputerName,[String]`$Hive,[String]`$Key,[String]`$ValueName)')"
        }

        return $myvalue
   
    }

    static [void]SetRegistryValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,$Value,[String]$ValueKind)
    {

        $reghive = [Reg]::New($ComputerName,$Hive)

        try
        {
            [Reg]::CreateSubKey($ComputerName,$Hive,$Key)
            $reghive.RemoteBaseKey.OpenSubKey($Key,$true).SetValue($ValueName,$Value,$ValueKind)
        }
        catch
        {
            Write-Warning "could not set registryvalue `'$Value`' of type `'$ValueKind`' on `'$ValueName`' in `'$Hive\$Key`' on `'$ComputerName`' (static [void]SetRegistryValue([String]`$ComputerName,[String]`$Hive,[String]`$Key,[String]`$ValueName,[String]`$Value,[String]`$ValueKind))"
        }

    }

    static [void]SetStringValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[String]$Value)
    {

        [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'String')

    }

    static [void]SetDWordValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[Int]$Value)
    {

            [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'DWord')

    }

    static [void]SetUnknownValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[String]$Value)
    {

            [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'Unknown')

    }

    static [void]SetExpandStringValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[String]$Value)
    {

        [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'ExpandString')

    }

    static [void]SetBinaryValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[Byte[]]$Value)
    {

        [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'Binary')

    }

    static [void]SetMultiStringValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[String[]]$Value)
    {

        [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'MultiString')

    }

    static [void]SetQWordValue([String]$ComputerName,[String]$Hive,[String]$Key,[String]$ValueName,[Single]$Value)
    {

        [Reg]::SetRegistryValue($ComputerName,$Hive,$Key,$ValueName,$Value,'QWord')

    }

    static [void]CreateSubKey([String]$ComputerName,[String]$Hive,[String]$SubkeyPath)
    {

        $reghive = [Reg]::New($ComputerName,$Hive)

        try
        {
            $reghive.RemoteBaseKey.CreateSubKey($SubkeyPath)
        }
        catch
        {
            Write-Warning "could not create key `'$SubkeyPath`' in `'$Hive`' on `'$ComputerName`' (static [void]CreateSubKey([String]`$ComputerName,[String]`$Hive,[String]`$SubkeyPath))"
        }
    }
#endregion static methods
}

#$r=[Reg]::ConnectHive("C120WIN7","localmachine")
#$r.GetRegistryValue("SOFTWARE\Harry\Subkey2\Subkey3","MyValueName")
#$r.ReturnRegAsHashTable("localmachine","SOFTWARE\Harry","string1")
#[Reg]::GetRegistryValue("localmachine","SOFTWARE\Harry","string1")
#[Reg]::GetRegistryValue("C1204418","localmachine","SOFTWARE\Harry","string1")
#[Reg]::CreateSubKey("C120WIN7","localmachine","SOFTWARE\Harry\AnotherKey")

#[Reg]::SetStringValue("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2\Subkey3","string2","My New Value")
#[Reg]::SetDWordValue("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2\Subkey3","MyDword",1235)
#[Reg]::SetBinaryValue("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2\Subkey3","MyBytevalues",[Byte[]]@(12,23,34,23))
#[Reg]::SetMultiStringValue("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2\Subkey3","MyMultiStringvalues",[String[]]@("hallo mensen","daar","ik","ben","het"))
#[Reg]::SetQWordValue("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2\Subkey3","MyQWord",1233423562456736765)
#[Reg]::DeleteSubKeyTree("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2")
#[Reg]::DeleteValue("C120WIN7","localmachine","SOFTWARE\Harry\Subkey2\Subkey3","string2")

