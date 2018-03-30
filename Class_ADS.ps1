Add-Type -AssemblyName System.DirectoryServices.AccountManagement

<#
    version 1.0.1
    test validatcredentials
#>

<# wishlist
    properties for required and post install groups
    ordering of collections
    adapt constructor to load commands if schemaclassname is group
    validatecredentials test...done
#>


<# problem...assembly needs to be loaded but does not load because the script processes the class first.
# it also does not want to load when included in the class
# it needs to be called from separate script
[reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")
#>

Class ADS : System.DirectoryServices.DirectorySearcher
{

#region properties
[System.DirectoryServices.DirectoryEntry]$ADObject
[System.Collections.ArrayList]$AllCommands=@()
[System.Collections.ArrayList]$MainCommands=@()
[System.Collections.ArrayList]$RequiredCommands=@()
[System.Collections.ArrayList]$PostCommands=@()
$mail
#endregion properties

#region constructors
ADS([String]$ObjectName) : base() {
#Add-Type -AssemblyName System.DirectoryServices.AccountManagement
#[reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement")
        $this.SizeLimit = 4000
        $this.Filter    = "(&(objectCategory=*)(CN=$ObjectName))"
        $this.mail = New-Object System.Net.Mail.MailAddressCollection

        try {
                [System.DirectoryServices.DirectoryEntry]$this.ADObject=([System.DirectoryServices.SearchResult]$this.FindOne()).GetDirectoryEntry()
                if ($this.ADObject.SchemaClassName -eq 'group')
                {
                    #$this.ReturnCommands()
                    #Write-Host "SchemaClassName of $ObjectName is `'group`'"
                }
            }
        catch
        {
            #Write-Host $Error[0].Exception.Message
            Write-Host "Could not find object $ObjectName"
        }
    }

ADS([String]$ObjectCategory,[String]$ObjectName) : base() {
#Add-Type -AssemblyName System.DirectoryServices.AccountManagement

        $this.SizeLimit = 4000
        $this.Filter    = "(&(objectCategory=$ObjectCategory)(CN=$ObjectName))"
        $this.mail = New-Object System.Net.Mail.MailAddressCollection

        try {

                [System.DirectoryServices.DirectoryEntry]$this.ADObject=([System.DirectoryServices.SearchResult]$this.FindOne()).GetDirectoryEntry()
                if ($this.ADObject.SchemaClassName -eq 'group')
                {
                    #$this.ReturnCommands()
                    #Write-Host "SchemaClassName of $ObjectName is `'group`'"
                }
            }
        catch
        {
            #Write-Host $Error[0].Exception.Message
            Write-Host "Could not find object $ObjectName of type `"$ObjectCategory`""
        }
    }
#endregion constructors

#region methods

#AddToGroup
    AddToGroup([String]$GroupName)
    {
        try
        {
            $objGroup = [ADS]::GetObject('group',$GroupName)
            $objGroup.Properties["member"].Add($($this.ADObject.distinguishedName))
            $objGroup.CommitChanges()
        }
        catch
        {
            Write-Host $Error[0].Exception.Message
        }

    }

#DisableUser
    DisableUser()
    {

        if($this.ADObject.SchemaClassName -ne 'user')
        {
            Write-Host "disabling objects of type `"$($this.ADObject.SchemaClassName)`" not allowed!"
        }
        else
        {
            $this.ADObject.InvokeSet("AccountDisabled",$true)
            $this.ADObject.CommitChanges()
        }
    }

#DisableComputer
    DisableComputer()
    {

        if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "disabling objects of type `"$($this.ADObject.SchemaClassName)`" not allowed!"
        }
        else
        {
            $this.ADObject.InvokeSet("AccountDisabled",$true)
            $this.ADObject.CommitChanges()
        }
    }

#EnableComputer
    EnableComputer()
    {

        if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "disabling objects of type `"$($this.ADObject.SchemaClassName)`" not allowed!"
        }
        else
        {
            $this.ADObject.InvokeSet("AccountDisabled",$false)
            $this.ADObject.CommitChanges()
        }
    }

#EnableUser
    EnableUser()
    {
        if($this.ADObject.SchemaClassName -ne 'user')
        {
            Write-Host "disabling objects of type `"$($this.ADObject.SchemaClassName)`" not allowed!"
        }
        else
        {
            $this.ADObject.InvokeSet("AccountDisabled",$false)
            $this.ADObject.CommitChanges()
        }
    }

#RemoveFromGroup
    RemoveFromGroup([String]$GroupName)
    {
        $objGroup = [ADS]::GetObject('group',$GroupName)
        $objGroup.Properties["member"].Remove($($this.ADObject.distinguishedName))
        $objGroup.CommitChanges()
    }

#SetLabeledUri
    SetlabeledUri([String]$PathToPatch)
    {

    if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Setting labeledUri on `"$($this.ADObject.SchemaClassName)`" object not allowed!"
        }
        else
        {
            $this.ADObject.labeledUri.Add($PathToPatch)
            $this.ADObject.CommitChanges()
        }
    }

#SetLabeledUri
    SetlabeledUri([System.Object[]]$PathToPatches)
    {

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Setting labeledUri on `"$($this.ADObject.SchemaClassName)`" object not allowed!"
        }
        else
        {
            foreach($PathToPatch in $PathToPatches)
            {
                Write-Host "Adding $PathToPatch to labeledUri attribute"
                $this.ADObject.labeledUri.Add($PathToPatch)
            }
            $this.ADObject.CommitChanges()
        }
    }

#RemovelabeledUriItem
    RemovelabeledUriItem([System.Object[]]$PathToPatches)
    {

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Setting labeledUri on `"$($this.ADObject.SchemaClassName)`" object not allowed!"
        }
        else
        {
            foreach($PathToPatch in $PathToPatches)
            {
                Write-Host "Removing $PathToPatch to labeledUri attribute(System.Object)"
                $this.ADObject.labeledUri.Remove($PathToPatch)
            }
            $this.ADObject.CommitChanges()
        }
    }

#RemovelabeledUriItem
    RemovelabeledUriItem([System.String]$PathToPatch)
    {

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Setting labeledUri on `"$($this.ADObject.SchemaClassName)`" object not allowed!"
        }
        else
        {
            Write-Host "Removing $PathToPatch to labeledUri attribute(System.String)"
            $this.ADObject.labeledUri.Remove($PathToPatch)
            $this.ADObject.CommitChanges()
        }
    }

#ClearLabeledUri
    ClearlabeledUri()
    {

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Setting labeledUri on `"$($this.ADObject.SchemaClassName)`" object not allowed!"
        }
        else
        {
            $this.ADObject.labeledUri=@()
            $this.ADObject.CommitChanges()
        }
    }

#MoveToNC
    MoveToNC()
    {

        $OU_NC='LDAP://OU=NC,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local'

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Move of `"$($this.ADObject.SchemaClassName)`" object not allowed to OU $OU_NC"
        }
        else
        {
            $this.ADObject.MoveTo($OU_NC)
        }
    }

#MoveToPC
    MoveToPC()
    {

        $OU_PC='LDAP://OU=PC,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local'

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Move of `"$($this.ADObject.SchemaClassName)`" object not allowed to OU $OU_PC"
        }
        else
        {
            $this.ADObject.MoveTo($OU_PC)
        }
    }

#MoveToBeheer
    MoveToBeheer()
    {

        $OU_Beheer='LDAP://OU=Beheer,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local'

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Move of `"$($this.ADObject.SchemaClassName)`" object not allowed to OU $OU_Beheer"
        }
        else
        {
            $this.ADObject.MoveTo($OU_Beheer)
        }
    }

#MoveToLaptop
    MoveToLaptop()
    {

        $OU_Laptop='LDAP://OU=Laptop,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local'

    `	if($this.ADObject.SchemaClassName -ne 'computer')
        {
            Write-Host "Move of `"$($this.ADObject.SchemaClassName)`" object not allowed to OU $OU_Laptop"
        }
        else
        {
            $this.ADObject.MoveTo($OU_Laptop)
        }

    }

#ValidateCredentials
    [System.Boolean]ValidateCredentials([System.Management.Automation.PSCredential]$CredentialObject)
    {
        $contexttype = New-Object [System.DirectoryServices.AccountManagement.ContextType]::Domain
        $principalcontext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $contexttype, "antoniuszorggroep.local", "DC=antoniuszorggroep,DC=local"
        return $principalcontext.ValidateCredentials($CredentialObject.GetNetworkCredential().UserName,$CredentialObject.GetNetworkCredential().Password)
    }

#ReturnCommands
    ReturnCommands()
    {
    `	if($this.ADObject.SchemaClassName -ne 'group')
        {
            Write-Host "Object is not of type group...cannot return commands"
        }

        $obj_app = New-Object PSCustomObject | Select-Object Name,InstallCommands,RemoveCommands,RegistryKey,RegistryValueName,RegistryValue,Mail
        

        $objGroup     = [ADS]::GetGroupObject($($this.ADObject.CN))

        if($($objGroup.mail)){
        $this.mail.add($($objGroup.mail))}
        $required     = $($objGroup.extensionAttribute14)
        $post         = $objGroup.extensionAttribute15
        $installcmds  = $objGroup.labeledUri
        $removecmds   = $objGroup.wbemPath
        $regkey       = $objGroup.extensionAttribute1
        $regvaluename = $objGroup.extensionAttribute2
        $regvalue     = $objGroup.extensionAttribute3


        if($required)
        {

            foreach ($requiredgroup in $required.Split(";"))
            {
                #Write-Host "required : $requiredgroup"
                $reqcmd=[ADS]::ReturnCommands($requiredgroup)
                #$reqcmd.Sort()
                if($reqcmd.Mail){
                    $this.mail.add($reqcmd.Mail)
                }
                $this.RequiredCommands.Add($reqcmd)
                $this.AllCommands.Add($reqcmd)
            }

        }
        
        if($post)
        {
            foreach ($postgroup in $post.Split(";"))
            {
                #Write-Host "post : $postgroup"
                $postcmd=[ADS]::ReturnCommands($postgroup)
                if($postcmd.Mail){
                    $this.mail.add($postcmd.Mail)
                }
                $this.PostCommands.Add($postcmd)
                $this.AllCommands.Add($postcmd)
            }
        }

        #main
        $obj_app.Name=$this.ADObject.CN
        $obj_app.InstallCommands=$installcmds | Sort-Object
        $obj_app.RemoveCommands=$removecmds | Sort-Object
        $obj_app.RegistryKey=$regkey
        $obj_app.RegistryValueName=$regvaluename
        $obj_app.RegistryValue=$regvalue

        $this.MainCommands.Add($obj_app)
        $this.AllCommands.Add($obj_app)
        
    }

#endregion methods

#region static

#DisableComputer
    static [System.DirectoryServices.DirectoryEntry]DisableComputer($ComputerName)
    {
        $ComputerObject = [ADS]::GetComputerObject($ComputerName)
        $ComputerObject.InvokeSet("AccountDisabled",$true)
        return $ComputerObject.CommitChanges()
    }

#DisableUser
    static [System.DirectoryServices.DirectoryEntry]DisableUser($UserName)
    {
        $UserObject = [ADS]::GetUserObject($UserName)
        $UserObject.InvokeSet("AccountDisabled",$true)
        return $UserObject.CommitChanges()
    }

#EnableComputer
    static [System.DirectoryServices.DirectoryEntry]EnableComputer($ComputerName)
    {
        $ComputerObject = [ADS]::GetComputerObject($ComputerName)
        $ComputerObject.InvokeSet("AccountDisabled",$false)
        return $ComputerObject.CommitChanges()
    }

#EnableUser
    static [System.DirectoryServices.DirectoryEntry]EnableUser($UserName)
    {
        $UserObject = [ADS]::GetUserObject($UserName)
        $UserObject.InvokeSet("AccountDisabled",$false)
        return $UserObject.CommitChanges()
    }

#GetObject
    static [System.DirectoryServices.DirectoryEntry]GetObject($ObjectName)
    {
        return [ADS]::New($ObjectName).ADObject
    }

#GetObject
    static [System.DirectoryServices.DirectoryEntry]GetObject($ObjectCategory,$ObjectName)
    {
        return [ADS]::New($ObjectCategory,$ObjectName).ADObject
    }

#GetComputerObject
    static [System.DirectoryServices.DirectoryEntry]GetComputerObject($ComputerName)
    {
        return [ADS]::GetObject("computer",$ComputerName)
    }

#GetGroupObject
    static [System.DirectoryServices.DirectoryEntry]GetGroupObject($GroupName)
    {
        return [ADS]::GetObject("group",$GroupName)
    }

#GetUserObject
    static [System.DirectoryServices.DirectoryEntry]GetUserObject($UserName)
    {
        return [ADS]::GetObject("user",$UserName)
    }

#MoveToNC
    static [System.DirectoryServices.DirectoryEntry]MoveToNC($ComputerName)
    {
        return [ADS]::GetComputerObject($ComputerName).MoveTo("LDAP://OU=NC,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local")
    }

#MoveToPC
    static [System.DirectoryServices.DirectoryEntry]MoveToPC($ComputerName)
    {
        return [ADS]::GetComputerObject($ComputerName).MoveTo("LDAP://OU=PC,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local")
    }

#MoveToPC\Proxy8081
    static [System.DirectoryServices.DirectoryEntry]MoveToPC8081($ComputerName)
    {
        return [ADS]::GetComputerObject($ComputerName).MoveTo("LDAP://OU=Proxy8081,OU=PC,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local")
    }

#MoveToNC\Proxy8081
    static [System.DirectoryServices.DirectoryEntry]MoveToNC8081($ComputerName)
    {
        return [ADS]::GetComputerObject($ComputerName).MoveTo("LDAP://OU=Proxy8081,OU=NC,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local")
    }

#MoveTo Beheer OU
    static [System.DirectoryServices.DirectoryEntry]MoveToBeheer($ComputerName)
    {
        return [ADS]::GetComputerObject($ComputerName).MoveTo("LDAP://OU=Beheer,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local")
    }

#MoveTo Laptop OU
    static [System.DirectoryServices.DirectoryEntry]MoveToLaptop($ComputerName)
    {
        return [ADS]::GetComputerObject($ComputerName).MoveTo("LDAP://OU=Laptop,OU=Werkstations,OU=AZG,DC=antoniuszorggroep,DC=local")
    }

#AddToGroup
    static AddToGroup([String]$ObjCategory,[String]$ObjectName,[String]$GroupName)
    {
        $objAD    = [ADS]::GetObject($ObjCategory,$ObjectName)
        $objGroup = [ADS]::GetObject('group',$GroupName)
        $objGroup.Properties["member"].Add($($objAD.distinguishedName))
        $objGroup.CommitChanges()
    }

#AddComputerToGroup
    static AddComputerToGroup([String]$ComputerName,[String]$GroupName)
    {
        [ADS]::AddToGroup('computer',$ComputerName,$GroupName)
    }

#IsMemberOfGroup
    static [System.Boolean]IsMemberOfGroup([String]$GroupName)
    {
        $objGroup    = [ADS]::GetGroupObject($GroupName)
        $objComputer = [ADS]::GetComputerObject($env:COMPUTERNAME)
        if($objGroup.member -contains $objComputer.distinguishedName)
        {
            return $true
        }
        else
        {
            return $false
        }
    }

    #IsMemberOfGroup
    static [System.Boolean]IsMemberOfGroup([String]$GroupName,[String]$ObjectName)
    {
        $objGroup    = [ADS]::GetGroupObject($GroupName)
        $objAD = [ADS]::GetObject($ObjectName)
        if($objGroup.member -contains $objAD.distinguishedName)
        {
            return $true
        }
        else
        {
            return $false
        }
    }

#AddUserToGroup
    static AddUserToGroup([String]$UserName,[String]$GroupName)
    {
        [ADS]::AddToGroup('user',$UserName,$GroupName)
    }


    [Boolean] static ValidateUserCredentials([System.Management.Automation.PSCredential]$CredentialObject)
    {
        $contexttype = New-Object [System.DirectoryServices.AccountManagement.ContextType]::Domain
        $principalcontext = New-Object -TypeName System.DirectoryServices.AccountManagement.PrincipalContext -ArgumentList $contexttype, "antoniuszorggroep.local", "DC=antoniuszorggroep,DC=local"
        return $principalcontext.ValidateCredentials($CredentialObject.GetNetworkCredential().UserName,$CredentialObject.GetNetworkCredential().Password)
    }

#ReturnCommands
    static [PSCustomObject]ReturnCommands([String]$GroupName)
    {

        $allcmds = New-Object PSCustomObject | Select-Object Name,InstallCommands,RemoveCommands,RegistryKey,RegistryValueName,RegistryValue,Mail
        $objGroup     = [ADS]::GetGroupObject($GroupName)
        #$required     = $objGroup.extensionAttribute14
        #$post         = $objGroup.extensionAttribute15
        $installcmds  = $objGroup.labeledUri
        $removecmds   = $objGroup.wbemPath
        $regkey       = $objGroup.extensionAttribute1
        $regvaluename = $objGroup.extensionAttribute2
        $regvalue     = $objGroup.extensionAttribute3
        #$m =$objGroup.mail

        
        $allcmds.Name=$GroupName
        $allcmds.InstallCommands=$installcmds | Sort-Object
        $allcmds.RemoveCommands=$removecmds | Sort-Object
        $allcmds.RegistryKey=$regkey
        $allcmds.RegistryValueName=$regvaluename
        $allcmds.RegistryValue=$regvalue
        if($objGroup.mail){
            $allcmds.Mail=$objGroup.mail
        }

        return $allcmds

    }

#endregion static

}