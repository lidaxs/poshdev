﻿<#
    version 1.0.0.3
    added helper HS2PS
    comparing objects outputs strings in the format of hashtable
    this helperfunction converts them to objects
    there is a issue with properties and values with spaces...

    version 1.0.0.2
    added objects to compare

    version 1.0.0.1
    added some documentation
    fixed typo

    version 1.0.0.0
    Initial upload
#>
function Compare-Set
{
<#
.Synopsis
   Compares two arrays of the same type
.DESCRIPTION
   Compares two arrays of the same type and shows differences,similarities and combinations
.EXAMPLE
    $ReferenceObject = [string[]]@("aap","noot","mies","aap")
    $DifferenceObject = [string[]]@("noot","zus","jet")
    Compare-Set -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject

    ReferenceObject  : {aap, noot, mies}
    DifferenceObject : {noot, zus, jet}
    InBoth           : {noot}
    Combined         : {aap, noot, mies, zus...}
    OnlyInReference  : {aap, mies}
    OnlyInDifference : {zus, jet}
    ExclusiveInBoth  : {aap, zus, mies, jet}

.EXAMPLE
    $ReferenceObject = [int[]](1,2,3,4)
    $DifferenceObject = [int[]](2,4,6,8)
    Compare-Set -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject

    ReferenceObject  : {1, 2, 3, 4}
    DifferenceObject : {2, 4, 6, 8}
    InBoth           : {2, 4}
    Combined         : {1, 2, 3, 4...}
    OnlyInReference  : {1, 3}
    OnlyInDifference : {6, 8}
    ExclusiveInBoth  : {1, 8, 3, 6}

.EXAMPLE
    $ReferenceObject=Get-ADComputer -Filter 'Name -like "C12040*"' | Select -ExpandProperty Name
    $DifferenceObject=Get-ADComputer -Filter 'Name -like "C120403*"' | Select -ExpandProperty Name
    Compare-Set -ReferenceObject $ReferenceObject -DifferenceObject $DifferenceObject

    ReferenceObject  : {C1204000, C1204032, C1204038, C1204056...}
    DifferenceObject : {C1204032, C1204038}
    InBoth           : {C1204032, C1204038}
    Combined         : {C1204000, C1204032, C1204038, C1204056...}
    OnlyInReference  : {C1204000, C1204056, C1204058, C1204076...}
    OnlyInDifference : {}
    ExclusiveInBoth  : {C1204000, C1204056, C1204058, C1204076...}

.INPUTS
   Inputs to this cmdlet [int[]],[string[]]
.OUTPUTS
   Output from this cmdlet System.Collections.Generic.HashSet`1
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
    [CmdletBinding()]
    [Alias('Compare-Array')]
    [OutputType([System.Collections.Generic.HashSet`1])]
    Param
    (
        # ReferenceObject help description
        [Parameter(Mandatory=$true)]
        [ValidateScript({($_ -is 'String') -or ($_ -is 'Int') -or ($_ -is 'Object')})]
        $ReferenceObject,

        # DifferenceObject help description
        [Parameter(Mandatory=$true)]
        [ValidateScript({($_ -is 'String') -or ($_ -is 'Int') -or ($_ -is 'Object')})]
        $DifferenceObject
    )

    Begin
    {
        try
        {
            # cast inputobjects to [String[]]
            $ReferenceObject=[String[]]$ReferenceObject
            $DifferenceObject=[String[]]$DifferenceObject
        }
        catch
        {
            Write-Warning "Could not cast objects to [String[]]"
        }
    }

    Process
    {
        # convert arrays to hashsets
        $HashSet1 = New-Object System.Collections.Generic.HashSet[$($ReferenceObject.GetType().Name.Trim("[]"))](,$ReferenceObject)
        $HashSet2 = New-Object System.Collections.Generic.HashSet[$($DifferenceObject.GetType().Name.Trim("[]"))](,$DifferenceObject)

        # create customobject for returning results
        $result = New-Object PSCustomObject | Select-Object ReferenceObject,DifferenceObject,InBoth,Combined,OnlyInReference,OnlyInDifference,ExclusiveInBoth

        # create copy of object because they change
        $InBoth = New-Object "System.Collections.Generic.HashSet[$($ReferenceObject.GetType().Name.Trim("[]"))]" $HashSet1
        $InBoth.IntersectWith($HashSet2)
        $result.InBoth           = $InBoth

        $Combined = New-Object "System.Collections.Generic.HashSet[$($ReferenceObject.GetType().Name.Trim("[]"))]" $HashSet1
        $Combined.UnionWith($HashSet2)
        $result.Combined           = $Combined

        $OnlyRef = New-Object "System.Collections.Generic.HashSet[$($ReferenceObject.GetType().Name.Trim("[]"))]" $HashSet1
        $OnlyRef.ExceptWith($HashSet2)
        $result.OnlyInReference    = $OnlyRef

        $OnlyDif = New-Object "System.Collections.Generic.HashSet[$($ReferenceObject.GetType().Name.Trim("[]"))]" $HashSet2
        $OnlyDif.ExceptWith($HashSet1)
        $result.OnlyInDifference    = $OnlyDif

        $ExclInBoth = New-Object "System.Collections.Generic.HashSet[$($ReferenceObject.GetType().Name.Trim("[]"))]" $HashSet1
        $ExclInBoth.SymmetricExceptWith($HashSet2)
        $result.ExclusiveInBoth    = $ExclInBoth

        $result.ReferenceObject  = $HashSet1
        $result.DifferenceObject = $HashSet2

        $result

    }
    End
    {
    }
}
function HS2PS ($HashString)
{
    $states = $HashString.Replace("@{","").Replace("}","").Split("; ")

    $out    = New-Object -TypeName PSCustomObject
    for ($i = 0; $i -lt $states.Count+1; $i+=2)
    {
        [String]$prop  = $states[$i].Split("=")[0]
        [String]$value = $states[$i].Split("=")[1]
        $value
        Add-Member -InputObject $out -MemberType NoteProperty -Name $prop -Value $value -Force
    }
    $out
}