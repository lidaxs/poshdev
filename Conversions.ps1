<#
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