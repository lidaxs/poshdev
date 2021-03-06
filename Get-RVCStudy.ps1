<#
	version 1.1.2.0
	added Get-RVCAuthorisedStudies

	version 1.1.1.1
	removed Import-RVCStudy
	find a way to import large textfiles first

	version 1.1.1.0
	added function Import-RVCStudy

	version 1.1.0.2
	changed query to support only Study_ID
	Now you have don't have to enter a patient_ID
	
	version 1.1.0.1
	added some formatting in script

	version 1.1.0
	Added function Move-RVCStudy

	version 1.0.1
	removed check parameter fill in beginblock
	removed report from the output except for export

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

		if(!(Get-Command Invoke-SqlCmd))
		{
			Write-Warning "Import Powershell SQL Module first!"
			Write-Verbose "For tooling download the correct installers (5 msi's) from http://www.microsoft.com/en-us/download/details.aspx?id=16978`nand the powershell SQLPS_Module from https://skydrive.live.com/?cid=ea42395138308430&id=EA42395138308430%21986&sc=documents"
			Write-Verbose "For more info....http://sev17.com/2010/07/making-a-sqlps-module/"
			break
		}

	}

	process
    {
        # create query
        $qry = "
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
			where Patient_ID like '%'"

        foreach ($PatID in $PatientID)
        {

			if ($PatID)
			{
				$PatID="{0:D7}" -f [int]$PatID
				if($qry -match "and Patient_ID")
				{
					$qry+="`nor Patient_ID = '$PatID'"
				}
				else
				{
					$qry+="`nand Patient_ID = '$PatID'"				
				}

			}
		}

        if ($StudyType) {
            $qry += "`nand Study_Type = '$StudyType'"
        }

        if ($Department) {
            $qry += "`nand Department = '$Department'"
        }

        if ($StudyID) {
            $qry += "`nand Study_ID like '$StudyID'"
        }

        if ($StudyDate) {
            $StudyDate = Get-Date $StudyDate -Format yyyy-MM-dd
            $qry += "`nand CONVERT(VARCHAR(25),Study_Date,126) like '$StudyDate%'"
        }

        if ($State) {
            $qry += "`nand State = '$State'"
        }

        if ($ShowQuery) {
            Write-Output $qry
            break
        }
		
        # execute query....notice size of MaxCharLength
        # this is because the report field contains lot of data.
        $output = Invoke-Sqlcmd -ServerInstance $SQLInstance -Database $Database -Query $qry -MaxCharLength 60000

        if ($ExportAsHTML) {
            if (-not (Test-Path -Path "$ExportPath")) {
                New-Item -Path $ExportPath -ItemType Directory -Force -ErrorAction SilentlyContinue
            }
            foreach ($item in $output) {
                #Write-Host $ExportPath
       
                try {
                    if ($IncludeErrors) {
                        $content = $($item.Report) + "<BR><BR><BR><html><body><div>$($item.Remarks)</div></body></html>"
                        $exportfile = "$ExportPath\$($item.Patient_ID)_$($item.Department)_$($item.Study_ID)_$($item.State)_ERR.html"
                        
                    }
                    else {
                        $content = $($item.Report)
                        $exportfile = "$ExportPath\$($item.Patient_ID)_$($item.Department)_$($item.Study_ID)_$($item.State).html"
                    }
                    New-Item $exportfile -ItemType File -Force
                }
                catch {

                }
                Set-Content -Value $content -Path $exportfile -Force
            }
        }
        else {
            # add Database property to out output so we can pass it to the pipeline
            foreach ($row in $output) {
                Add-Member -InputObject $row  -MemberType NoteProperty -Name Database -Value $Database
            }
            $output | Select-Object Patient_ID, Study_ID, Study_Type, Study_Date, Physician, Location, Referrer, Diagnose, Authorized, Remarks, Database
        }
    }
}

function Move-RVCStudy
{
	<#
	.SYNOPSIS
		Moves a study to another patient
	.DESCRIPTION
		Sometimes personnel create a study and link it to the wrong patient
		This function can correct this
	.EXAMPLE
		# find out the ID of the study
		Get-RVCStudy -PatientID 0000015 -StudyType RUSTECG | Select Study_ID,Study_Date

		Study_ID                                                         Study_Date
		--------                                                         ----------
		2.16.840.1.113883.2.4.3.25.20141215.11.34.54.554.188266016674382 15-12-2014 11:31:46
		2.16.840.1.113883.2.4.3.25.20141215.11.30.54.294.176241446013114 15-12-2014 11:25:48
		....

		# pick the correct Study_ID from the list(copy) and use it in the next command

		Get-RVCStudy -PatientID 0000015 -StudyID 2.16.840.1.113883.2.4.3.25.20141215.11.30.54.294.176241446013114 | Move-RVCStudy -PatientID 0001000

		Patient_ID : 0001000
		Study_ID   : 2.16.840.1.113883.2.4.3.25.20141215.11.30.54.294.176241446013114 
		Study_Type : RUSTECG
		Study_Date : 15-12-2014 11:25:48
		Physician  :
		Location   : CCU 6
		Referrer   :
		Diagnose   :
		Authorized : adm-bouweh01 31-10-2016 09:12:04
		Remarks    :
		Database   : RVC

		On databaselevel the Patient_ID is set to the new value 0001000
	.INPUTS
		[string]Study_ID
		[string]PatientID
		[string]SQLInstance
		[string]Database
		[switch]ShowQuery
	.OUTPUTS
		System.Management.Automation.PSCustomObject
	.NOTES
		General notes
	#>
	[CmdletBinding()]
	param (

		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[System.String]
		$Study_ID,

		[Parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, HelpMessage="Vul hier de 7 cijferige PatientID in!")]
		[ValidateScript({"{0:D7}" -f $_})]
		[System.String]
		$PatientID,

		[Parameter(Mandatory=$false)]
		[System.String]
		$SQLInstance='CLT-SQL02CLSQL2\CLTSQL02SH02',

		[Parameter(Mandatory=$false,  ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Mogelijke waarden voor deze parameter zijn RVC, RVC_ACC")]
		[ValidateSet("RVC", "RVC_ACC", IgnoreCase = $true)]
		[System.String]
		$Database='RVC',

		[Parameter(Mandatory=$false)]
		[Switch]
		$ShowQuery		
	)
	
	begin
	{

	}
	
	process 
	{
		# construct query
		$updateqry="
		UPDATE [$Database].[dbo].[All_Studies]
		SET [Patient_ID] = '$PatientID'
		WHERE [Study_ID] = '$Study_ID'
		"

		if($ShowQuery)
		{
			Write-Output $updateqry
			break
		}

		# update record with new patientid
		Invoke-Sqlcmd -ServerInstance $SQLInstance -Database $Database -Query $updateqry

		# check if correct
		Get-RVCStudy -PatientID $PatientID -StudyID $Study_ID -Database $Database -SQLInstance $SQLInstance

	}
	
	end
	{

	}
}

function Get-RVCAuthorisedStudies
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory=$false)]
		[System.DateTime]
		$Date = [datetime]::Today,

		[Parameter(Mandatory=$false,  ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage="Mogelijke waarden voor deze parameter zijn RVC, RVC_ACC")]
		[ValidateSet("RVC", "RVC_ACC", IgnoreCase = $true)]
		[System.String]
		$Database='RVC',

		[Parameter(Mandatory=$false)]
		[System.String]
		$SQLInstance='CLT-SQL02CLSQL2\CLTSQL02SH02',


		[Parameter(Mandatory=$false)]
		[Switch]
		$IncludeFailed,	

		[Parameter(Mandatory=$false)]
		[Switch]
		$MailReport,

		[Parameter(Mandatory=$false)]
		[System.String[]]
		$MailTo = "h.bouwens@antoniuszorggroep.nl",

		[Parameter(Mandatory=$false)]
		[Switch]
		$ShowQuery		
	)

	begin
	{
		$mydate       = Get-Date $Date -Format "dd-MM-yyyy"
		$mailserver  = "srv-mail02.antoniuszorggroep.local"

		# check if we can execute invoke-sqlcmd
		if ( -not (Get-Command -Name Invoke-Sqlcmd -ErrorAction SilentlyContinue) )
		{
			try {
				Write-Verbose "importing module SqlServer"
				Import-Module SqlServer
			}
			catch {
				Write-Warning "Could not import module SqlServer..exiting"
				break
			}
			
		}


		$qry_error="
		SELECT [allstudies].[Patient_ID]
			,[allstudies].[Department]
			,[allstudies].[Study_ID]
			,[allstudies].[Study_Type]
			,[allstudies].[Study_Date]
			,[allstudies].[Referrer]
			,[allstudies].[Location]
			,[allstudies].[Physician]
			,[allstudies].[Remarks]
			,[allstudies].[Authorized]
			,[allstudies].[Report]
			,phys.[Email]
		FROM [RVC].[dbo].[All_Studies] as allstudies
		Inner Join [RVC].[dbo].[All_Physicians] as phys
		On[allstudies].[Physician] = phys.[Physician]
		Where Authorized like '%$mydate%'
		And allstudies.Remarks like '%Error%'
		"

		$qry_nonerror="
		SELECT [allstudies].[Patient_ID]
			,[allstudies].[Department]
			,[allstudies].[Study_ID]
			,[allstudies].[Study_Type]
			,[allstudies].[Study_Date]
			,[allstudies].[Referrer]
			,[allstudies].[Location]
			,[allstudies].[Physician]
			,[allstudies].[Remarks]
			,[allstudies].[Authorized]
			,[allstudies].[Report]
			,phys.[Email]
		FROM [RVC].[dbo].[All_Studies] as allstudies
		Inner Join [RVC].[dbo].[All_Physicians] as phys
		On[allstudies].[Physician] = phys.[Physician]
		Where Authorized like '%$mydate%'
		And allstudies.Remarks not like '%Error%'
		"

		if($ShowQuery)
		{
			Write-Output $qry_error
			Write-Output $qry_nonerror
			break
		}
	}
	process
	{

		$results_nonerr     = Invoke-Sqlcmd -ServerInstance $sqlinstance -Query $qry_nonerror -MaxCharLength 50000

		if($IncludeFailed)
		{
			$results_err        = Invoke-Sqlcmd -ServerInstance $sqlinstance -Query $qry_error -MaxCharLength 50000
			foreach ($result_err in $results_err) {
			
				$subject = "Fout bij het autoriseren van verslag van patient $($result_err.Patient_ID)"
				$body    = "Opmerkingen : " + $result_err.Remarks.ToString()  + "</BR>"
				$body   += "Study_ID : "    + $result_err.Study_ID.ToString() + "</BR>"
				$body   += $result_err.Report.ToString()
	
				if ($MailReport)
				{
					Write-Verbose "sending mail to $MailTo"
					Send-MailMessage -SmtpServer $mailserver -Body $body -From auth_error_rvc@antoniuszorggroep.nl -Subject $subject -BodyAsHtml -To $MailTo
				}
				else
				{
					Write-Output $results_err | Select-Object Patient_ID,Department,Study_ID,Study_Type,Study_Date,Referrer,Location,Physician,Remarks,Authorized,Email
				}
			}
		}


		foreach ($result_nonerr in $results_nonerr)
		{
			$subject = "Verslag van patient $($result_nonerr.Patient_ID) is goed verzonden"
			$body    = "Opmerkingen : " + $result_nonerr.Remarks.ToString()  + "</BR>"
			$body   += "Study_ID : "    + $result_nonerr.Study_ID.ToString() + "</BR>"
			$body   += $result_nonerr.Report.ToString()

			if ($MailReport)
			{	
				Write-Verbose "sending mail to $MailTo"
				Send-MailMessage -SmtpServer $mailserver -Body $body -From rvc@antoniuszorggroep.nl -Subject $subject -BodyAsHtml -To $MailTo
			}
			else
			{
				Write-Output $results_nonerr | Select-Object Patient_ID,Department,Study_ID,Study_Type,Study_Date,Referrer,Location,Physician,Remarks,Authorized,Email
			}
		}
	}
}