<# 
    version 1.0.0.2
    little bug fixed when no applications found

    version 1.0.0.1
    added TransformApplicationsToVisioList()

    version 1.0.0.0
    init upload
#>
Class CM_ProviderLocation
{
    #region properties
    $SCCMServer = 'srv-sccm02'
    $Machine
    $NamespacePath
    $NameSpace
    $ProviderForLocalSite
    $SiteCode
    $PSComputerName
    #endregion properties

    #region constructors
    CM_ProviderLocation() : Base()
    {
        $qry       = "select * from SMS_ProviderLocation where ProviderForLocalSite = true"
        $rootsms = "root\sms"

        $result    = Get-WmiObject -Query $qry -ComputerName $this.SCCMServer-Namespace $rootsms
        $this.Machine = $result.Machine
        $this.NamespacePath = $result.NamespacePath
        $this.ProviderForLocalSite = $result.ProviderForLocalSite
        $this.PSComputerName = $result.PSComputerName
        $this.SiteCode = $result.SiteCode
        $this.NameSpace = "$rootsms\site_$($this.SiteCode)"
    }
    #endregion constructors

    #region static
    static [String]GetSiteCode()
        {
            return [CM_ProviderLocation]::new().SiteCode
        }

    static [String]GetNameSpace()
        {
            return [CM_ProviderLocation]::new().NameSpace
        }
    #endregion static
}

Class CM_ARP
{
    #region properties
    [System.Collections.ArrayList]$Applications
    $Query
    #endregion properties

    #region constructors

    CM_ARP([String]$ClientName) : Base()
    {
        $this.GetApps($([SMS_R_SYSTEM]::new($ClientName).ResourceID))
    }

    CM_ARP([Int]$ResourceID) : Base()
    {
        $this.GetApps($($ResourceID))
    }
    #endregion constructors

    #region methods
    GetApps([Int]$ResourceID)
    {
        $this.Query = "Select * From SMS_G_System_ADD_REMOVE_PROGRAMS Where ResourceID='$($ResourceID)'"
        $this.Applications = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $this.Query -Namespace ([CM_ProviderLocation]::GetNameSpace()) | Select-Object DisplayName,Version,Publisher,ProdID,InstallDate
    }
    #endregion methods
}

Class CM_ARP64
{
    #region properties
    [System.Collections.ArrayList]$Applications
    $Query
    #endregion properties

    #region constructors

    CM_ARP64([String]$ClientName) : Base()
    {
        $this.GetApps($([SMS_R_SYSTEM]::new($ClientName).ResourceID))
    }

    CM_ARP64([Int]$ResourceID) : Base()
    {
        $this.GetApps($($ResourceID))
    }
    #endregion constructors

    #region methods
    GetApps([Int]$ResourceID)
    {
        $this.Query = "Select * From SMS_G_System_ADD_REMOVE_PROGRAMS_64 Where ResourceID='$($ResourceID)'"
        $this.Applications = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $this.Query -Namespace ([CM_ProviderLocation]::GetNameSpace()) | Select-Object DisplayName,Version,Publisher,ProdID,InstallDate
    }
    #endregion methods
}

Class CM_MemberOfCollection
{
    #region properties
    [System.Collections.ArrayList]$Collections=@()
    $Query
    #endregion properties

    #region constructors

    CM_MemberOfCollection([String]$ClientName) : Base()
    {
        $this.GetCollections($([SMS_R_SYSTEM]::new($ClientName).ResourceID))
    }

    CM_MemberOfCollection([Int]$ResourceID) : Base()
    {
        $this.GetCollections($($ResourceID))
    }
    #endregion constructors

    #region methods
    GetCollections([Int]$ResourceID)
    {
        $this.Query = "Select CollectionID From SMS_FullCollectionMembership Where ResourceID='$([int]$ResourceID)'"
        #$qry="Select CollectionID From SMS_FullCollectionMembership Where ResourceID='$([int]$resID)'"
        #$FullCollectionMembership = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
        $FullCollectionMembership = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $this.Query -Namespace ([CM_ProviderLocation]::GetNameSpace())
        foreach($item in $FullCollectionMembership)
        {
            $res=Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Class SMS_Collection -Namespace ([CM_ProviderLocation]::GetNameSpace()) -Filter "CollectionID='$($item.CollectionID)'"
            $this.Collections.Add($res)
        }
    }
    #endregion methods
}

Class SMS_R_System
{
    #region properties
    $ResourceID
    $Name
    $WMIObject
    #endregion properties

    #region constructors
    SMS_R_System([String]$Name) : Base()
        {
            $qry             = "Select * FROM SMS_R_SYSTEM WHERE Name='$($Name)'"
            $this.Name       = $Name
            $this.WMIObject  = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
            $this.ResourceId = $this.WMIObject.ResourceID
        }
    SMS_R_System([Int]$ResourceID) : Base()
        {
            $qry             = "Select Name FROM SMS_R_SYSTEM WHERE ResourceID='$($ResourceID)'"
            $this.ResourceId = $ResourceID
            $this.WMIObject  = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
            $this.Name       = $this.WMIObject.Name
        }
    #endregion constructors
}

Class CM_SYSTEM
{
    #region properties
    $Active
    $ADSiteName
    $Advertisements
    $AgentName
    $AgentSite
    $AgentTime
    $AlwaysInternet
    $AMTFullVersion
    $AMTStatus
    $Applications
    $Applications32
    $Applications64
    $Client
    $ClientEdition
    $ClientType
    $ClientVersion
    $Collections
    $CPUType
    $CreationDate
    $Decommissioned
    $department
    $DeviceOwner
    $DistinguishedName
    $EASDeviceID
    $FullDomainName
    $HardwareID
    $InternetEnabled
    $IPAddresses
    $IPSubnets
    $IPv6Addresses
    $IPv6Prefixes
    $IsAOACCapable
    $IsAssignedToUser
    $IsClientAMT30Compatible
    $IsMachineChangesPersisted
    $IsPortableOperatingSystem
    $IsVirtualMachine
    $IsWriteFilterCapable
    $LastLogonTimestamp
    $LastLogonUserDomain
    $LastLogonUserName
    $MACAddresses
    $MDMComplianceStatus
    $Name
    $NetbiosName
    $ObjectGUID
    $Obsolete
    $OperatingSystemNameandVersion
    $PreviousSMSUUID
    $PrimaryGroupID
    $PublisherDeviceID
    $ResourceDomainORWorkgroup
    $ResourceId
    $ResourceNames
    $ResourceType
    $SecurityGroupName
    $SID
    $SMBIOSGUID
    $SMSAssignedSites
    $SMSInstalledSites
    $SMSResidentSites
    $SMSUniqueIdentifier
    $SMSUUIDChangeDate
    $SNMPCommunityName
    $SuppressAutoProvision
    $SystemContainerName
    $SystemGroupName
    $SystemOUName
    $SystemRoles
    $Unknown
    $UserAccountControl
    $VirtualMachineHostName
    $VirtualMachineType
    $WipeStatus
    $WTGUniqueKey
    $VisioAppList
    $PSComputerName
    hidden $WMIObject
    #endregion properties

    #region constructors
    CM_SYSTEM([int]$ResourceID) : Base()
    {
        $qry            = "Select * FROM SMS_R_SYSTEM WHERE ResourceID='$($ResourceID)'"
        $this.ResourceId = $ResourceID
        $this.WMIObject = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
        $this.ProcessProperties()
    }

    CM_SYSTEM([String]$Name) : Base()
    {
        $qry    = "Select * FROM SMS_R_SYSTEM WHERE Name='$($Name)'"
        $this.Name = $Name
        $this.WMIObject = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
        $this.ProcessProperties()
    }
    #endregion constructors

    #region methods
    ProcessProperties()
    {
        # map values to the class properties
        foreach($item in $this.WMIObject.Properties){
            $this."$($item.Name)"=$item.Value
        }


        $this.LastLogonTimestamp = [DateTimeConverter]::ToDateTime($this.LastLogonTimestamp)
        $this.SMSUUIDChangeDate = [DateTimeConverter]::ToDateTime($this.SMSUUIDChangeDate)
        $this.AgentTime = [DateTimeConverter]::ToDateTime($this.AgentTime)
        $this.CreationDate = [DateTimeConverter]::ToDateTime($this.CreationDate)
        #$this.GetApplications()
        #$this.GetCollections()
        #$this.GetAdvertisements()
    }

    GetAdvertisements()
    {
        #[System.Collections.ArrayList]$col_advertisements=@()
        #$clientadvstatus     = [CM_ClientAdvertisementStatus]::new($this.ResourceId).AdvertisementName
        $this.Advertisements = [CM_ClientAdvertisementStatus]::new($this.ResourceId).AdvertisementNames
    }

    GetApplications()
    {
        $this.Applications32 = [CM_ARP]::new([int]$this.ResourceId).Applications
        $this.Applications64 = [CM_ARP64]::new([int]$this.ResourceId).Applications
        $this.Applications = $this.Applications32 + $this.Applications64
    }

    GetCollections()
    {
        $this.Collections = [CM_MemberOfCollection]::new([int]$this.ResourceId).Collections
    }

    RemoveFromCM()
    {
        try
        {
            Remove-WmiObject -InputObject $this.WMIObject -Confirm -Verbose
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Warning "could not remove system from SCCM...binding to object failed(does `"$($this.Name)$($this.ResourceID)`" actually exist?)"
        }
    }

    [String]TransformApplicationsToVisioList()
    {
        $this.GetApplications()
        [System.Collections.ArrayList]$applicationlist = $this.Applications.DisplayName
        if($applicationlist)
        {
            $applicationlist.Sort()
            $this.VisioAppList = $applicationlist -join ";"
            return $this.VisioAppList
        }
        else
        {
            return "no applications"    
        }

    }
    #endregion methods

    #region static
    static RemoveFromCM([int]$ResourceID)
    {
        [CM_System]::new($ResourceID).RemoveFromCM()
    }
    static RemoveFromCM([string]$Name)
    {
        [CM_System]::new($Name).RemoveFromCM()
    }
    #endregion static
}

Class CM_Collection
{
    #region properties
    $Name
    $WMIObject
    $CollectionID
    $CollectionMembers
    $CollectionRules
    $CollectionType
    $CollectionVariablesCount
    $Comment
    $CreationDate
    $CurrentStatus
    $HasProvisionedMember
    $IncludeExcludeCollectionsCount
    $IsBuiltIn
    $IsReferenceCollection
    $ISVData
    $ISVDataSize
    $ISVString
    $LastChangeTime
    $LastMemberChangeTime
    $LastRefreshTime
    $LimitToCollectionID
    $LimitToCollectionName
    $LocalMemberCount
    $MemberClassName
    $MemberCount
    $MonitoringFlags
    $OwnedByThisSite
    $PowerConfigsCount
    $RefreshSchedule
    $RefreshType
    $ReplicateToSubSites
    $ServiceWindowsCount
    #endregion properties

    #region constructors
    CM_Collection([String]$Name) : Base()
    {
        $qry    = "Select * FROM SMS_Collection WHERE Name='$($Name)'"
        $this.Name = $Name
        $this.WMIObject = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
        $this.ProcessProperties()
    }
    #endregion constructors

    #region methods
    GenerateClientConfigurationRequestByName($Name,$PushSiteCode,$Forced)
    {
    
    }

    ProcessProperties()
    {
        # map values to the class properties
        foreach($item in $this.WMIObject.Properties){

            $this."$($item.Name)"=$item.Value
        }

        $this.LastChangeTime = [DateTimeConverter]::ToDateTime($this.LastChangeTime)
        $this.LastMemberChangeTime = [DateTimeConverter]::ToDateTime($this.LastMemberChangeTime)
        $this.LastRefreshTime = [DateTimeConverter]::ToDateTime($this.LastRefreshTime)
        $this.CreationDate = [DateTimeConverter]::ToDateTime($this.CreationDate)

        $this.GetCollectionMembers()
    }

    GetCollectionMembers()
    {
        $this.CollectionMembers = Get-WmiObject -Class SMS_CollectionMember_a -Namespace ([CM_ProviderLocation]::GetNameSpace()) -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Filter "CollectionID='$($this.CollectionID)'"
    }

    RemoveFromCM()
    {
        try
        {
            Remove-WmiObject -InputObject $this.WMIObject -Confirm -Verbose
        }
        catch [System.Management.Automation.ParameterBindingException]
        {
            Write-Warning "could not remove collection from SCCM...binding to object failed(does `"$($this.Name)`" actually exist?)"
        }
    }

    RequestRefresh([Boolean]$IncludeSubCollections)
    {
        $this.WMIObject.RequestRefresh($IncludeSubCollections)
    }
    #endregion methods

    #region static
    [System.Collections.ArrayList]
    static GetCollectionMembers([String]$CollectionName)
    {
        $Collection = [CM_Collection]::New($CollectionName)
        return $Collection.CollectionMembers
    }

    static RemoveFromCM([int]$CollectionID)
    {
        [CM_Collection]::new($CollectionID).RemoveFromCM()
    }

    static RemoveFromCM([string]$Name)
    {
        [CM_Collection]::new($Name).RemoveFromCM()
    }

    static RequestRefresh([String]$Name,[Boolean]$IncludeSubCollections)
    {
        [CM_Collection]::new($Name).RequestRefresh($IncludeSubCollections)
    }
    #endregion static
}

Class CM_Client
{
    #region properties
    $Namespace = "root\CCM"
    $Class = "SMS_Client"
    $WMIObject
    $AllowLocalAdminOverride
    $ClientType
    $ClientVersion
    $EnableAutoAssignment
    $AssignedSite
    $PSComputerName
    #endregion properties

    #region constructors
    CM_Client() : Base()
    {

    }

    CM_Client([String]$ComputerName)
    {
        $this.WMIObject = Get-WmiObject -Class $this.Class -Namespace $this.Namespace -ComputerName $ComputerName
        $this.PSComputerName = $ComputerName
        $this.ProcessProperties()
    }
    #endregion constructors

    #region methods
    GetAssignedSite()
    {
        $this.AssignedSite = (Invoke-WMIMethod -Name "GetAssignedSite" -ComputerName $this.PSComputerName -Class $this.Class -Namespace $this.Namespace).sSiteCode
    }

    EvaluateMachinePolicy()
    {
        Invoke-WMIMethod -Name "EvaluateMachinePolicy" -ComputerName $this.PSComputerName -Class $this.Class -Namespace $this.Namespace
    }

    RequestMachinePolicy()
    {
        Invoke-WMIMethod -Name "RequestMachinePolicy" -ComputerName $this.PSComputerName -Class $this.Class -Namespace $this.Namespace
    }

    ProcessProperties()
    {
        # map values to the class properties
        foreach($item in $this.WMIObject.Properties){
            $this."$($item.Name)"=$item.Value
        }

        $this.GetAssignedSite()

    }

    ExecuteApplicationDeploymentCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000121}')
    }

    ExecuteDiscoveryDataCollectionsCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000103}')
    }
 
     ExecuteFileCollectionsCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000104}')
    }

     ExecuteHardwareInventoryCollectionsCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000101}')
    }

     ExecuteSoftwareInventoryCollectionsCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000102}')
    }

     ExecuteSoftwareUpdatesDeploymentCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000108}')
    }

     ExecuteSoftwareUpdatesScanCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000113}')
    }

     ExecuteMachinePolicyRetrievalEvaluationCycle()
    {
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000021}')
        $this.RunTriggerSchedule('{00000000-0000-0000-0000-000000000022}')
    }
    
    RunTriggerSchedule($sScheduleID)
    {
        Invoke-WMIMethod -Name "TriggerSchedule" -ArgumentList $sScheduleID -ComputerName $this.PSComputerName -Class $this.Class -Namespace $this.Namespace
    }
    #endregion methods

    #region static
    static [String]GetAssignedSite([String]$ComputerName)
    {
        return (Invoke-WMIMethod -Name "GetAssignedSite" -ComputerName $ComputerName -Class SMS_Client -Namespace root\CCM).sSiteCode
    }

    static EvaluateMachinePolicy([String]$ComputerName)
    {
        Invoke-WMIMethod -Name "EvaluateMachinePolicy" -ComputerName $ComputerName -Class SMS_Client -Namespace root\CCM
    }

    static RequestMachinePolicy([String]$ComputerName)
    {
        Invoke-WMIMethod -Name "RequestMachinePolicy" -ComputerName $ComputerName -Class SMS_Client -Namespace root\CCM
    }


    static RunTriggerSchedules($ComputerName)
    {
        $CMClient = [CM_Client]::New($ComputerName)
        $CMClient.RequestMachinePolicy()
        $CMClient.EvaluateMachinePolicy()
        $CMClient.ExecuteApplicationDeploymentCycle()
        $CMClient.ExecuteSoftwareUpdatesScanCycle()
        $CMClient.ExecuteSoftwareUpdatesDeploymentCycle()
        $CMClient.ExecuteMachinePolicyRetrievalEvaluationCycle()
    }
    #endregion static
}

Class CM_Advertisement
{
    #region properties
    $ActionInProgress
    $AdvertFlags
    $AdvertisementID
    $AdvertisementName
    $AssignedSchedule
    $AssignedScheduleEnabled
    $AssignedScheduleIsGMT
    $AssignmentID
    $CollectionID
    $Comment
    $DeviceFlags
    $ExpirationTime
    $ExpirationTimeEnabled
    $ExpirationTimeIsGMT
    $HierarchyPath
    $IncludeSubCollection
    $ISVData
    $ISVDataSize
    $IsVersionCompatible
    $ISVString
    $MandatoryCountdown
    $OfferType
    $PackageID
    $PresentTime
    $PresentTimeEnabled
    $PresentTimeIsGMT
    $Priority
    $ProgramName
    $RemoteClientFlags
    $SourceSite
    $TimeFlags
    $WMIObject
    #endregion properties

    #region constructors
    CM_Advertisement([String]$AdvertisementID) : Base()
    {
        $qry                  = "Select * FROM SMS_Advertisement WHERE AdvertisementID='$($AdvertisementID)'"
        $this.AdvertisementID = $AdvertisementID
        $this.WMIObject       = Get-WmiObject -ComputerName ([CM_ProviderLocation]::new().SCCMServer) -Query $qry -Namespace ([CM_ProviderLocation]::GetNameSpace())
        $this.ProcessProperties()
    }
    #endregion constructors

    #region methods
    ProcessProperties()
    {
        # map values to the class properties
        foreach($item in $this.WMIObject.Properties)
        {
            $this."$($item.Name)"=$item.Value
        }

        #convert datetime properties to readable format
        $this.ExpirationTime = [DateTimeConverter]::ToDateTime($this.ExpirationTime)
        $this.PresentTime    = [DateTimeConverter]::ToDateTime($this.PresentTime)

    }

    [System.Collections.ArrayList]
    ConvertDateTimes(
        [String[]]
        $DateCollection
    )
    {
        #convert times to readable format
        $newdatecollection = New-Object System.Collections.ArrayList
        foreach($wmidate in $DateCollection)
        {
             $newdatecollection.Add([System.Management.ManagementDateTimeConverter]::ToDateTime($wmidate))
        }
        return $newdatecollection
    }
    #endregion methods

    #region static

    #endregion static
}

Class CM_ClientAdvertisementStatus
{
    #region properties
    $AdvertisementID
    $AdvertisementNames
    $LastAcceptanceMessageID
    $LastAcceptanceMessageIDName
    $LastAcceptanceMessageIDSeverity
    $LastAcceptanceState
    $LastAcceptanceStateName
    $LastAcceptanceStatusTime
    $LastExecutionContext
    $LastExecutionResult
    $LastState
    $LastStateName
    $LastStatusMessageID
    $LastStatusMessageIDName
    $LastStatusMessageIDSeverity
    $LastStatusTime
    $ResourceID
    $WMIObject
    #endregion properties

    #region constructors
    CM_ClientAdvertisementStatus([Int]$ResourceID) : Base()
    {
        #$this.AdvertisementID = $AdvertisementID
        $this.ResourceID      = $ResourceID
        $qry                  = "Select * from SMS_ClientAdvertisementStatus WHERE ResourceID='$($ResourceID)'"
        $this.WMIObject       = Get-WMIObject -Computer ([CM_ProviderLocation]::new().SCCMServer) -Namespace ([CM_ProviderLocation]::new().NameSpace) -Query $qry
        $this.ProcessProperties()
    }   
    #endregion constructors

    #region methods
    ProcessProperties()
    {
        # map values to the class properties
        foreach($item in $this.WMIObject.Properties)
        {
                $this."$($item.Name)"=$item.Value
        }
    
        $this.LastAcceptanceStatusTime = [DateTimeConverter]::ToDateTime($this.LastAcceptanceStatusTime)
        $this.LastStatusTime = [DateTimeConverter]::ToDateTime($this.LastStatusTime)
        $this.GetAdvertisementNames()
    }

    GetAdvertisementNames()
    {
        [System.Collections.ArrayList]$col_adv=@()
        foreach($item in $this.WMIObject)
        {
            $col_adv.Add([CM_Advertisement]::new($($item.AdvertisementID)).AdvertisementName)
        }
        $this.AdvertisementNames = $col_adv
    }
    #endregion methods

    #region static

    #endregion static
}

Class DateTimeConverter
{
    #region properties
    $WMIDate
    $DateTime
    #endregion properties

    #region constructors
    DateTimeConverter([String[]]$WMIDate) : Base()
    {
        $this.WMIDate = $WMIDate
        $this.ToDateTime()
    }
    #endregion constructors

    #region methods
    ToDateTime()
    {
        [System.Collections.ArrayList]$datecollection = @()
        foreach($dte in $this.WMIDate)
        {
            $datecollection.Add([System.Management.ManagementDateTimeConverter]::ToDateTime($dte))
        }

        $this.DateTime=$datecollection

    }
    #endregion methods

    #region static
    static [String[]]ToDateTime($WMIDate)
    {
        return [DateTimeConverter]::new($WMIDate).DateTime
    }
    #endregion static
}
