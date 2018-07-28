<#
	version 1.0.0.1
	formatted script

    version 1.0.0.0
    ported from function get-registryvalue

	version 1.0.1.1
	changed test-connection to portscan on 139

	version 1.0.1
	Converted $Hive parameter to dynamic parameter

	version 1.0.0
	Initial upload
	Added aliases to ClientName parameter to support pipeline in from WMI,SCCM & Active Directory
#>
Function Get-RegistryTree {
	<#
		.SYNOPSIS
			Gets a registry values remote or local.

		.DESCRIPTION
			Gets a registry values remote or local.

		.PARAMETER  ComputerName
			The ComputerName(s) on which to operate.

		.PARAMETER  Hive
			The Hive to pick ...LocalMachine.

		.PARAMETER  Key
			The key to query.

		.EXAMPLE
			Get-RegistryKey -ComputerName C120VMXP -Hive LocalMachine -Key "SOFTWARE\MyKey\MySubKey" -Depth 2

		.INPUTS
			System.String,System.String[]

		.OUTPUTS
			System.String

		.NOTES
			Additional information about the function go here.

		.LINK
			about_functions_advanced

		.LINK
			about_comment_based_help
#Requires –Version 3
	#>
	[CmdletBinding()]
	[OutputType([System.Object])]
	param(
		[Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
		[Alias("CN","Name","PSComputerName","MachineName","Workstation","ServerName","HostName","ComputerName")]
		[ValidateNotNullOrEmpty()]
		$ClientName=@($env:COMPUTERNAME),

		[Parameter(Position=2, Mandatory=$false)]
		[System.String]
        $Key
	)
    DynamicParam
    {
        $HiveParameterAttributes                   = New-Object System.Management.Automation.ParameterAttribute
        $HiveParameterAttributes.Mandatory         = $true
        $HiveParameterAttributes.HelpMessage       = "Press `'TAB`' to cycle through the different values"
		$HiveParameterAttributes.ParameterSetName  = '__AllParameterSets'
		$HiveAttributeCollection                   = New-Object  System.Collections.ObjectModel.Collection[System.Attribute]
		$HiveAttributeCollection.Add($HiveParameterAttributes)
		$Hive                                      = [enum]::GetNames([Microsoft.Win32.RegistryHive])
		$HiveAttributeCollection.Add((New-Object  System.Management.Automation.ValidateSetAttribute($Hive)))
		$HiveRuntimeParameters                     = New-Object System.Management.Automation.RuntimeDefinedParameter('Hive', [System.String], $HiveAttributeCollection)

		$RuntimeParametersDictionary               = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
		$RuntimeParametersDictionary.Add('Hive', $HiveRuntimeParameters)

		return  $RuntimeParametersDictionary
	}

begin{

	}

process{

		# Loop through collection
		ForEach($Computer in $ClientName){

			# Test connectivity
			if([System.Net.Sockets.TcpClient]::new().ConnectAsync($Computer,139).AsyncWaitHandle.WaitOne(1000,$false))
			{
				try
				{
					$RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]$PSBoundParameters.Hive,$Computer,[Microsoft.Win32.RegistryView]::Registry64)
                    $objKey = $RemoteRegistry.OpenSubKey($Key)

                    foreach($iValueName in $objKey.GetValueNames())
                    {
                        [PSCustomObject]$out = "" | Select-Object Computer,Hive,Key,ValueName,Value,ValueKind
                        $out.Computer  = $Computer
                        $out.Hive      = $PSBoundParameters.Hive
                        $out.Key       = $Key
                        $out.ValueName = $iValueName
                        $out.Value     = $objKey.GetValue($iValueName)
                        $out.ValueKind = $objKey.GetValueKind($iValueName)

                        Write-Output $out
                    }

                    if($RemoteRegistry.SubKeyCount -gt 0)
                    {
                        foreach($iKey in $objKey.GetSubKeyNames())
                        {
                            Get-RegistryTree -Computer $Computer -Hive $PSBoundParameters.Hive -Key $Key\$iKey
                        }
                    }
				}

				catch [System.Management.Automation.MethodInvocationException]
				{
					Write-Warning "Cannot open RegistryKey $Key in $PSBoundParameters.Hive on $Computer"
				}
            }

		else{
				Write-Warning "$Computer is not online!"
			}
		}
	}

	end
	{
		try
		{

		}
		catch
		{

		}
	}
}