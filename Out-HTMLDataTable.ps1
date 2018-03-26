<#

v1.0.1
Empty rows were included...this is fixed

#>
function Out-HTMLDataTable
{
<#
    .Synopsis
    Short description
    .DESCRIPTION
    Long description
    .EXAMPLE
    Example of how to use this cmdlet
    .EXAMPLE
    Another example of how to use this cmdlet
    .INPUTS
    Inputs to this cmdlet (if any)
    .OUTPUTS
    Output from this cmdlet (if any)
    .NOTES
    General notes
    .COMPONENT
    The component this cmdlet belongs to
    .ROLE
    The role this cmdlet belongs to
    .FUNCTIONALITY
    The functionality that best describes this cmdlet
#>
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([System.Object])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, ValueFromRemainingArguments=$false)]
        $InputObject,
		
		[Switch]
		$AsString
    )

    Begin
    {
    #styling voor de tabel in de emailbody
    $style = "<style>BODY{font-family: Arial; font-size: 8pt;} "
    $style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;} "
    $style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; } "
    $style = $style + "TD{border: 1px solid black; padding: 5px; } "
    $style = $style + "</style>"
    $temptable=@()
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("$InputObject", "Out-HTMLDataTable"))
        {

            $addrow = $false

            foreach($item in $InputObject){
                foreach($prop in $item)
                {
                    if ( -not ([String]::IsNullOrEmpty($prop)))
                    {
                        $addrow=$true
                    }
                }
                if ($addrow)
                {
                    [void]$temptable.Add($item)
                    $addrow = $false
                }
            }
        }
        
        $output=$temptable  | ConvertTo-Html -Head $style
        
    }
    End
    {
		if($AsString){
			$output | Out-String
		}
		else{
        	$output
		}
    }
}