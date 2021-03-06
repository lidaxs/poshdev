<#
    version 1.0.1.1
    variable temptable set to System.Collections.ArrayList
    
    v1.0.1
    Empty rows were included...this is fixed

    v1.0.2
    Updated Synopsis

#>
function Out-HTMLDataTable
{
<#
    .Synopsis
    Converts objects to html datatable
    .DESCRIPTION
    Converts objects to html datatable
    Useful for outputting objects to a report and send this by mail
    .EXAMPLE
    Out-HTMLDataTable -InputObject $myobject -AsString
    .EXAMPLE
    $mailbody = Get-Service | Out-HTMLDataTable -AsString
    .INPUTS
    [System.Object[]]
    .OUTPUTS
    Text converted to HTML Datatable
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
    [System.Collections.ArrayList]$temptable=@()
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