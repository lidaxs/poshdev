<#
    version 1.0.2.0
    added ToPSCustom

    version 1.0.1.0
    added ToHashTable

    version 1.0.0.0
    first upload
#>
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
                   ValueFromPipelineByPropertyName=$true,
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
                #Write-Output "DataRow"
                #[PSCustomObject]$item
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
        }
    }
    
    end {
    }
}