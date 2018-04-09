<#
	version 1.0.0
	Initial upload
#>
function Get-RVCStudy {
	<#
		.SYNOPSIS
			Retrieves the studies from the RVC database by PatientID,Department or StudyType.

		.DESCRIPTION
			Retrieves the studies from the RVC database by PatientID,Department or StudyType.

		.PARAMETER  PatientID
			The Study to be retrieved by PatientID.

		.PARAMETER  Department
			The Study to be retrieved by Department.

		.PARAMETER  StudyType
			The Study to be retrieved by StudyType.

		.PARAMETER  StudyID
			The Study to be retrieved by StudyID.

		.PARAMETER  StudyDate
			The Study to be retrieved by StudyDate.

		.PARAMETER  ExportAsHTML
			Exports the reports as htmlfile.

		.PARAMETER  ExportPath
			Path where exportfiles will be saved.

		.PARAMETER  IncludeErrors
			Includes the remark field which contains the errors in the htmlfile.

		.PARAMETER  SQLInstance
			The SQL instance to operate on(default CLT-SQL02CLSQL2\CLTSQL02SH02)

		.PARAMETER  Database
			The database to query(default RVC)

		.EXAMPLE
			PS C:\> Get-RVCStudy -PatientID 9858680

		.EXAMPLE
 			PS C:\> 9858680,0000015 | Get-RVCStudy -Department INT

		.INPUTS
			System.String

		.NOTES
			Additional information about the function go here.

		.LINK
			http://www.microsoft.com/en-us/download/details.aspx?id=16978
			https://skydrive.live.com/?cid=ea42395138308430&id=EA42395138308430%21986&sc=documents
			http://sev17.com/2010/07/making-a-sqlps-module/

		.LINK
			about_comment_based_help

	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	[OutputType([System.Data.DataRow])]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Vul hier de 7 cijferige PatientID in!")]
		[ValidateScript({"{0:D7}" -f $_})]
		[System.String[]]
		$PatientID,

		[Parameter(Mandatory=$false)]
        [ValidateSet("COLOSCOPIE", "COLODDKS", "GASTROSCOP", "RUSTECG", "ERCP", "BRONCHO", "GASTRODUO", "OESOPHAGUS", "PEGSONDE", "SIGMOIDOSC", "VideoCap", "Echo", IgnoreCase = $true)]
		[System.String]
		$StudyType,

		[Parameter(Mandatory=$false)]
		[ValidateSet("INT", "CAR", "CHI", "LON", "URO", "MFO",IgnoreCase=$true)]
		[System.String]
		$Department,

		[Parameter(Mandatory=$false)]
		[System.String]
		$StudyDate,

		[Parameter(Mandatory=$false)]
		[System.String]
		$StudyID='%',

		[Parameter(Mandatory=$false)]
		[ValidateSet("HC_MSGFAIL", "HL7CDASEND", "HLCDADEMUR", "HL7SETOUT", "NonDICA", "%", IgnoreCase = $true)]
		[System.String]
		$State,

		[Parameter(Mandatory=$false)]
		[Switch]
		$ExportAsHTML,

		[Parameter(Mandatory=$false)]
		[Switch]
		$IncludeErrors,

		[Parameter(Mandatory=$false)]
        [System.String]
		$ExportPath="$($HOME)RVCExport",

		[Parameter(Mandatory=$false)]
		[System.String]
		$SQLInstance='CLT-SQL02CLSQL2\CLTSQL02SH02',

		[Parameter(Mandatory=$false, HelpMessage="Mogelijke waarden voor deze parameter zijn RVC, RVC_ACC")]
		[ValidateSet("RVC", "RVC_ACC", IgnoreCase = $true)]
		[System.String]
		$Database='RVC',

		[Parameter(Mandatory=$false)]
		[Switch]
		$ShowQuery
	)
begin
{

	if(!(Get-Module sqlps))
	{
		Write-Warning "Import Powershell SQL Module first!"
		Write-Verbose "For tooling download the correct installers (5 msi's) from http://www.microsoft.com/en-us/download/details.aspx?id=16978`nand the powershell SQLPS_Module from https://skydrive.live.com/?cid=ea42395138308430&id=EA42395138308430%21986&sc=documents"
		Write-Verbose "For more info....http://sev17.com/2010/07/making-a-sqlps-module/"
		break
	}

	if(($PatientID -eq '%') -and ($StudyType -eq '%') -and ($Department -eq '%'))
	{
		Write-Warning 'You should at least provide a PatientID,StudyType or Department'
		Write-Warning 'Exiting function.....'
		break
	}
}

process
    {
        foreach ($PatID in $PatientID)
        {

        $PatID="{0:D7}" -f [int]$PatID


# create query
$qry="
SELECT [Study_ID]
,[Patient_ID]
,[Department]
,[Study_Type]
,[Study_Date]
,[Accession]
,[Referrer]
,[Location]
,[Physician]
,[Diagnose]
,[Remarks]
,[Authorized]
,[Report]
,[State]
,[Properties]
FROM [$Database].[dbo].[All_Studies]
where Patient_ID like '$PatID'"

            if ($StudyType)
            {
                $qry+="`nand Study_Type = '$StudyType'"
            }

            if ($Department)
            {
                $qry+="`nand Department = '$Department'"
            }

		    if ($StudyID) {
			    $qry+="`nand Study_ID like '$StudyID'"
		    }

            if ($StudyDate)
            {
                $StudyDate=Get-Date $StudyDate -Format yyyy-MM-dd
                $qry+="`nand CONVERT(VARCHAR(25),Study_Date,126) like '$StudyDate%'"
            }

		    if ($State) {
			    $qry+="`nand State = '$State'"
		    }




            if($ShowQuery)
            {
                Write-Output $qry
                break
            }
		
		    # execute query....notice size of MaxCharLength
		    # this is because the report field contains lot of data.
            $output=Invoke-Sqlcmd -ServerInstance $SQLInstance -Database $Database -Query $qry -MaxCharLength 60000

            if($ExportAsHTML)
            {
                if(-not (Test-Path -Path "$ExportPath"))
                {
                    New-Item -Path $ExportPath -ItemType Directory -Force -ErrorAction SilentlyContinue
                }
                foreach($item in $output)
                {
                    #Write-Host $ExportPath
                
                
                    try{
                        if($IncludeErrors)
                        {
                            $content=$($item.Report) + "<BR><BR><BR><html><body><div>$($item.Remarks)</div></body></html>"
                            $exportfile="$ExportPath\$($item.Patient_ID)_$($item.Department)_$($item.Study_ID)_$($item.State)_ERR.html"
                        
                        }
                        else
                        {
                            $content=$($item.Report)
                            $exportfile="$ExportPath\$($item.Patient_ID)_$($item.Department)_$($item.Study_ID)_$($item.State).html"
                        }
                        New-Item $exportfile -ItemType File -Force
                    }
                    catch
                    {

                    }
                    Set-Content -Value $content -Path $exportfile -Force
                }
            }
            else
            {
                $output
            }
	    }
    }
}