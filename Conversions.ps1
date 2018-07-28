<#
    version 1.0.9.4
    added class errorobject
    
    version 1.0.9.3
    added  class conversion...just for test

    version 1.0.9.2
    added [CM_SYSTEM] to hash
    
    version 1.0.9.0
    added Base64ToString
    added StringToBase64

    version 1.0.8.0
    added RomanToInt
    
    version 1.0.7.0
    added IntToRoman

    version 1.0.6.0
    added IntToBinary
    added BinaryToInt
    renamed tohex --> IntToHex

    version 1.0.4.0
    added ToHex

    version 1.0.3.0
    added ConvertFrom-ErrorRecord

    version 1.0.2.0
    added ToPSCustom

    version 1.0.1.0
    added ToHashTable

    version 1.0.0.0
    first upload
#>

Class Conversion{
    $WhatGoesIn
    $WhatComesOut
    $EncodingTypes = @('ASCII', 'UTF7', 'UTF8', 'UTF32', 'UniCode')
    Conversion([Byte[]]$ByteArray) : Base()
    {
        $this.WhatGoesIn = $ByteArray
        $Encoding = 'ASCII'
        $this.WhatComesOut = [System.Text.Encoding]::$Encoding.GetString($ByteArray)
    }
    Conversion([Byte[]]$ByteArray,$Encoding) : Base()
    {
        $this.WhatGoesIn = $ByteArray
        if($Encoding -in $this.EncodingTypes){
            Write-Host "type `'$Encoding`' found...continuing"
            $this.WhatComesOut = [System.Text.Encoding]::$Encoding.GetString($ByteArray)
        }
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

Function ByteArrayToString {
    param(
        [Byte[]]$ByteArray,

        [ValidateSet('ASCII', 'UTF7', 'UTF8', 'UTF32', 'UniCode')]
        $Encoding = 'ASCII'
    )

    return [System.Text.Encoding]::$Encoding.GetString($ByteArray)
}

Function ByteArrayToHexString {

    [cmdletbinding()]
    
    param(
        [parameter(Mandatory = $true)]
        [Byte[]]
        $Bytes
    )
    
    $HexString = [System.Text.StringBuilder]::new($Bytes.Length * 2)
    
    ForEach ($byte in $Bytes) {
        $HexString.AppendFormat("{0:x2}", $byte) | Out-Null
    }
    
    $HexString.ToString()
}

Function HexToByteArray {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [String]
        $HexString
    )
    
    $Bytes = [byte[]]::new($HexString.Length / 2)
    
    For ($i = 0; $i -lt $HexString.Length; $i += 2) {
        $Bytes[$i / 2] = [convert]::ToByte($HexString.Substring($i, 2), 16)
    }
    
    $Bytes
}

Function StringToByteArray {
    [cmdletbinding()]
    param(
        [parameter(Mandatory = $true)]
        [String]
        $String
    )

    $byteArray = [System.Byte[]]@()
    $charArray = $String.ToCharArray()
    foreach ($char in $charArray) {$byteArray += [convert]::ToByte($char)}
    return $byteArray
}

Function ToHashTable
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([System.Collections.HashTable])]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $InputObject

    )

    Begin
    {
        #$InputObject.GetType()
    }
    Process
    {
        foreach($item in $inputobject)
        {
            #$item.GetType()
            if($item.GetType().Name -eq 'PSCustomObject')
            {
                [System.Collections.Hashtable]$out=@{}
                foreach($thing in $item.psobject.properties)
                {
                    $out.Add($($thing.Name),$($thing.Value))
                    #Write-Host "$($thing) : $($thing.Value)"
                }

                Write-Output $out
            }
            if($item.GetType().Name -eq 'CM_SYSTEM')
            {
                [System.Collections.Hashtable]$out=@{}
                foreach($thing in $item.psobject.properties)
                {
                    $out.Add($($thing.Name),$($thing.Value))
                    #Write-Host "$($thing) : $($thing.Value)"
                }

                Write-Output $out
            }

            if($item.GetType().Name -eq 'HashTable')
            {
                Write-Output $item
            }

            if($item.GetType().BaseType.Name -eq "ValueType")
            {
                $out=@{}
                $out.Add($item.GetType().Name,$item)
                Write-Output $out
            }

            if($item.GetType().Name -like "String*")
            {
                $out=@{}
                $out.Add($item.GetType().Name,$item)
                Write-Output $out
            }

            if($item.GetType().Name -eq "ManagementObject")
            {
                $out=@{}
                foreach($property in $item.properties){
                    $out.Add($property.Name,$property.Value)
                }
                Write-Output $out
            }
        }
    }
    End
    {
    }
}

function ToPSCustom {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,    
        ValueFromPipelineByPropertyName = $true,
        ValueFromPipeline = $true,
        Position=0
        )]
        $InputObject
    )
    
    begin {
        
    }
    
    process {
        
        foreach ($item in $InputObject) {
            if ($item.GetType().Name -eq 'HashTable') {
                $out=New-Object -TypeName psobject -Property $item
                $out 
            }
            if ($item.GetType().Name -eq 'DataRow') {
                [PSCustomObject]$out=@{}
                foreach ($column in $item.Table.Columns) {
                    Add-Member -InputObject $out -MemberType NoteProperty -Name $column.ColumnName -Value $item.$column.ColumnName
                }

                $out

            }
            if ($item.GetType().Name -eq 'PSCustomObject') {
                $item
            }
            if ($item.GetType().Name -eq 'ManagementObject')
            {
                $out=New-Object -TypeName PSCustomObject
                foreach ($prop in $item.properties)
                {
                    Add-Member -InputObject $out -MemberType NoteProperty -Name $($prop.Name) -Value $($prop.Value) -Force
                }

                Write-Output $out
            
            }
            if ($item.GetType().Name -eq 'CimInstance')
            {
                $out=New-Object -TypeName PSCustomObject
                foreach ($prop in $item.properties)
                {
                    Add-Member -InputObject $out -MemberType NoteProperty -Name $($prop.Name) -Value $($prop.Value) -Force
                }

                Write-Output $out
            
            }
        }
    }
    
    end {
    }
}

function ConvertFrom-ErrorRecord
{
  [CmdletBinding(DefaultParameterSetName="ErrorRecord")]
  param
  (
    [Management.Automation.ErrorRecord]
    [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="ErrorRecord", Position=0)]
    $Record,
    
    [Object]
    [Parameter(Mandatory,ValueFromPipeline,ParameterSetName="Unknown", Position=0)]
    $Alien
  )
  
  process
  {
    if ($PSCmdlet.ParameterSetName -eq 'ErrorRecord')
    {
      [PSCustomObject]@{
        Exception = $Record.Exception.Message
        Reason    = $Record.CategoryInfo.Reason
        Target    = $Record.CategoryInfo.TargetName
        Script    = $Record.InvocationInfo.ScriptName
        Line      = $Record.InvocationInfo.ScriptLineNumber
        Column    = $Record.InvocationInfo.OffsetInLine
      }
    }
    else
    {
      Write-Warning "$Alien"
    } 
  }
}

function IntToHex {
    [CmdletBinding()]
    param(
        # Parameter help description
        [Parameter(Mandatory=$true, ValueFromPipeline = $true)]
        [long[]]
        $InputObject,

        [Int]
        $Length = 8,

        [Switch]
        $UsePrefix
    )

    begin
    {
        $Prefix= ""
    }

    process
    {
        foreach ($item in $InputObject) {
            if($UsePrefix){$Prefix = "0x"}
            $hex ="$Prefix{0:X$Length}" -f $item
            $hex
        }
        
    }

    end
    {

    }
}

function IntToBinary {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $InputObject
    )
    
    begin {
    }
    
    process {
        foreach ($item in $InputObject) {
            [convert]::ToString($item,2)
        }
        
    }
    
    end {
    }
}

function BinaryToInt {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $InputObject
    )
    
    begin {
    }
    
    process {
        foreach ($item in $InputObject) {
            [convert]::ToInt32($item,2)
        }
        
    }
    
    end {
    }
}

function RomanToInt {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        Position=0)]
        [String[]]
        $InputObject
    )
    
    begin {
        [String[]]$roman1=@("MMM","MM","M")
        [String[]]$roman2=@("CM", "DCCC", "DCC", "DC", "D", "CD", "CCC", "CC", "C")
        [String[]]$roman3=@("XC", "LXXX", "LXX", "LX", "L", "XL", "XXX", "XX", "X")
        [String[]]$roman4=@("IX", "VIII", "VII", "VI", "V", "IV", "III", "II", "I")
        
    }
    
    process
    {
        foreach($item in $InputObject)
            {
            $value=0
            for ($i = 0; $i -lt 3; $i++) {
                if($item.StartsWith($roman1[$i]))
                {
                    $len = $roman1[$i].Length
                    $value += 1000 * (3 - $i)
                    break
            
                }
            }
            
            if($len -gt 0)
            {
                $item = $item.Substring($len)
                $len=0
            }

            for ($i = 0; $i -lt 9; $i++) {
                if($item.StartsWith($roman2[$i]))
                {
                    $value += 100 * (9 - $i)
                    $len = $roman2[$i].Length
                    break
                }
                
            }
        
            if($len -gt 0)
            {
                $item = $item.Substring($len)
                $len=0
            }

            for ($i = 0; $i -lt 9; $i++) {
                if($item.StartsWith($roman3[$i]))
                {
                    $value += 10 * (9 - $i)
                    $len = $roman3[$i].Length
                    break
                }
                
            }

            if($len -gt 0)
            {
                $item = $item.Substring($len)
                $len=0
            }

            for ($i = 0; $i -lt 9; $i++) {
                if($item.StartsWith($roman4[$i]))
                {
                    $value += 9 - $i
                    $len = $roman4[$i].Length
                    break
                }

            }

            $value

        }
    }
    
    end
    {
    }
}

function IntToRoman {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        Position=0)]
        [Int]
        [ValidateRange(1,3999)]
        $InputObject
    )
    
    begin {
        [String[]]$roman1=@("MMM","MM","M")
        [String[]]$roman2=@("CM", "DCCC", "DCC", "DC", "D", "CD", "CCC", "CC", "C")
        [String[]]$roman3=@("XC", "LXXX", "LXX", "LX", "L", "XL", "XXX", "XX", "X")
        [String[]]$roman4=@("IX", "VIII", "VII", "VI", "V", "IV", "III", "II", "I")
        
    }

    process
    {
        foreach($item in $InputObject)
        {
            $thousands=[Math]::Floor($item/1000)
            $item%=1000
            $hundreds=[Math]::Floor($item/100)
            $item%=100
            $tens=[Math]::Floor($item/10)
            $item%=10
            $units=$item%10

            $StringBuilder= [System.Text.StringBuilder]::new()
            if ($thousands -gt 0){ [void]$StringBuilder.Append($roman1[3 - $thousands])}
            if ($hundreds -gt 0){ [void]$StringBuilder.Append($roman2[9 - $hundreds])}
            if ($tens -gt 0){ [void]$StringBuilder.Append($roman3[9 - $tens])}
            if ($units -gt 0){ [void]$StringBuilder.Append($roman4[9 - $units])}
            $StringBuilder.ToString()
        }
    }

    end
    {

    }
}

function Base64ToString {
    [CmdletBinding()]
    param (
        [String[]]
        $Base64String,

        [String]
        [ValidateSet('ASCII','UTF7','UTF8','Unicode')]
        $CodeType = 'ASCII'

    )
    
    begin {
    }
    
    process {
        foreach($item in $Base64String)
        {
            [System.Text.Encoding]::$CodeType.GetString([System.Convert]::FromBase64String($item))
        }
        
    }
    
    end {
    }
}

function StringToBase64 {
    [CmdletBinding()]
    param (
        [String[]]
        $String,

        [String]
        [ValidateSet('ASCII','UTF7','UTF8','Unicode')]
        $CodeType = 'ASCII'

    )
    
    begin {
    }
    
    process {
        foreach($item in $String)
        {
            [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String($item))
        }
        
    }
    
    end {
    }
}
