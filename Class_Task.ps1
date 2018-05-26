<#
    version 1.0.2
    renamed method DeleteExpiredTaskAfter to DeleteExpiredTaskAfterXHours

    version 1.0.1
    fixed few typo's
#>

#region Enums

Enum TaskType{
    TASK_ACTION_EXEC          = 0
    TASK_ACTION_COM_HANDLER   = 5
    TASK_ACTION_SEND_EMAIL    = 6
    TASK_ACTION_SHOW_MESSAGE  = 7
}

Enum TaskCompatibility  { 
    TASK_COMPATIBILITY_AT  = 0
    TASK_COMPATIBILITY_V1  = 1
    TASK_COMPATIBILITY_V2  = 2
    TASK_COMPATIBILITY_V3  = 3
}

Enum TaskResult {
    SCHED_S_TASK_OK                       = 0x00000000
    SCHED_S_TASK_INSTALLED_REBOOTREQUIRED = 0x0000065E
    SCHED_S_TASK_READY                    = 0x00041300
    SCHED_S_TASK_RUNNING                  = 0x00041301
    SCHED_S_TASK_DISABLED                 = 0x00041302
    SCHED_S_TASK_HAS_NOT_RUN              = 0x00041303
    SCHED_S_TASK_NO_MORE_RUNS             = 0x00041304
    SCHED_S_TASK_NOT_SCHEDULED            = 0x00041305
    SCHED_S_TASK_TERMINATED               = 0x00041306
    SCHED_S_TASK_NO_VALID_TRIGGERS        = 0x00041307
    SCHED_S_EVENT_TRIGGER                 = 0x00041308
    SCHED_E_TRIGGER_NOT_FOUND             = 0x80041309
    SCHED_E_TASK_NOT_READY                = 0x8004130A
    SCHED_E_TASK_NOT_RUNNING              = 0x8004130B
    SCHED_E_SERVICE_NOT_INSTALLED         = 0x8004130C
    SCHED_E_CANNOT_OPEN_TASK              = 0x8004130D
    SCHED_E_INVALID_TASK                  = 0x8004130E
    SCHED_E_ACCOUNT_INFORMATION_NOT_SET   = 0x8004130F
    SCHED_E_ACCOUNT_NAME_NOT_FOUND        = 0x80041310
    SCHED_E_ACCOUNT_DBASE_CORRUPT         = 0x80041311
    SCHED_E_NO_SECURITY_SERVICES          = 0x80041312
    SCHED_E_UNKNOWN_OBJECT_VERSION        = 0x80041313
    SCHED_E_UNSUPPORTED_ACCOUNT_OPTION    = 0x80041314
    SCHED_E_SERVICE_NOT_RUNNING           = 0x80041315
    SCHED_E_UNEXPECTEDNODE                = 0x80041316
    SCHED_E_NAMESPACE                     = 0x80041317
    SCHED_E_INVALIDVALUE                  = 0x80041318
    SCHED_E_MISSINGNODE                   = 0x80041319
    SCHED_E_MALFORMEDXML                  = 0x8004131A
    SCHED_S_SOME_TRIGGERS_FAILED          = 0x0004131B
    SCHED_S_BATCH_LOGON_PROBLEM           = 0x0004131C
    SCHED_E_TOO_MANY_NODES                = 0x8004131D
    SCHED_E_PAST_END_BOUNDARY             = 0x8004131E
    SCHED_E_ALREADY_RUNNING               = 0x8004131F
    SCHED_E_USER_NOT_LOGGED_ON            = 0x80041320
    SCHED_E_INVALID_TASK_HASH             = 0x80041321
    SCHED_E_SERVICE_NOT_AVAILABLE         = 0x80041322
    SCHED_E_SERVICE_TOO_BUSY              = 0x80041323
    SCHED_E_TASK_ATTEMPTED                = 0x80041324
    SCHED_S_TASK_QUEUED                   = 0x00041325
    SCHED_E_TASK_DISABLED                 = 0x80041326
    SCHED_E_TASK_NOT_V1_COMPAT            = 0x80041327
    SCHED_E_START_ON_DEMAND               = 0x80041328
}

Enum LogonType{ 
    TASK_LOGON_NONE                           = 0
    TASK_LOGON_PASSWORD                       = 1
    TASK_LOGON_S4U                            = 2
    TASK_LOGON_INTERACTIVE_TOKEN              = 3
    TASK_LOGON_GROUP                          = 4
    TASK_LOGON_SERVICE_ACCOUNT                = 5
    TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD  = 6
}

Enum RunLevel{ 
    TASK_RUNLEVEL_LUA      = 0
    TASK_RUNLEVEL_HIGHEST  = 1
}

Enum Priority{
    THREAD_PRIORITY_TIME_CRITICAL = 0
    THREAD_PRIORITY_HIGHEST       = 1
    THREAD_PRIORITY_ABOVE_NORMAL2 = 2
    THREAD_PRIORITY_ABOVE_NORMAL3 = 3
    THREAD_PRIORITY_NORMAL4       = 4
    THREAD_PRIORITY_NORMAL5       = 5
    THREAD_PRIORITY_NORMAL6       = 6
    THREAD_PRIORITY_BELOW_NORMAL7 = 7
    THREAD_PRIORITY_BELOW_NORMAL8 = 8
    THREAD_PRIORITY_LOWEST        = 9
    THREAD_PRIORITY_IDLE          = 10
}

Enum TriggerType{
    TASK_TRIGGER_EVENT                = 1
    TASK_TRIGGER_TIME                 = 1
    TASK_TRIGGER_DAILY                = 2
    TASK_TRIGGER_WEEKLY               = 3
    TASK_TRIGGER_MONTHLY              = 4
    TASK_TRIGGER_MONTHLYDOW           = 5
    TASK_TRIGGER_IDLE                 = 6
    TASK_TRIGGER_REGISTRATION         = 7
    TASK_TRIGGER_BOOT                 = 8
    TASK_TRIGGER_LOGON                = 9
    TASK_TRIGGER_SESSION_STATE_CHANGE = 11
}

Enum TaskCreation{
    TASK_VALIDATE_ONLY                = 1
    TASK_CREATE                       = 2
    TASK_UPDATE                       = 4
    TASK_CREATE_OR_UPDATE             = 6
    TASK_DISABLE                      = 8
    TASK_DONT_ADD_PRINCIPAL_ACE       = 16
    TASK_IGNORE_REGISTRATION_TRIGGERS = 32
}

Enum State{
    TASK_STATE_UNKNOWN  = 0
    TASK_STATE_DISABLED = 1
    TASK_STATE_QUEUED   = 2
    TASK_STATE_READY    = 3
    TASK_STATE_RUNNING  = 4
}

#endregion Enums

Class Task
{

#region properties
    $TaskObject
    $objScheduler
    $TaskFolder
    [System.Management.Automation.PSCredential]$CredentialObject
#endregion properties

#region constructors
    Task([String]$ComputerName
        #[String]$DisplayName,
        #[String]$TaskFolder,
        #[String]$TaskType
        ) : Base()
    {
        if(Test-Connection $ComputerName -Count 1 -Quiet)
        {
            $this.objScheduler=New-Object -ComObject "Schedule.Service"
            $this.objScheduler.Connect($ComputerName)
        }
        else
        {
            Write-Warning "$ComputerName is not online..."
            break
        }
    }

    Task([String]$ComputerName,
        [System.Management.Automation.PSCredential]$CredentialObject
        ) : Base()
    {
        if(Test-Connection $ComputerName -Count 1 -Quiet)
        {
            $this.objScheduler     = New-Object -ComObject "Schedule.Service"
            $this.CredentialObject = $CredentialObject
            $this.objScheduler.Connect($ComputerName)
        }
        else
        {
            Write-Warning "$ComputerName is not online..."
            break
        }
    }
#endregion constructors

#region methods
    AddNewTask($TaskType)
    {
        $this.TaskObject=$this.objScheduler.NewTask($TaskType)
    }

    AddExecAction($Path,$Arguments,$WorkingDirectory)
    {

        $newAction                  = $this.TaskObject.Actions.Create([TaskType]::TASK_ACTION_EXEC)
        $newAction.Path             = $Path
        $newAction.Arguments        = $Arguments
        $newAction.WorkingDirectory = $WorkingDirectory
    }

    AddMailAction($Server,$To,$Cc,$From,$Subject,$Body)
    {
        $newAction         = $this.TaskObject.Actions.Create([TaskType]::TASK_ACTION_SEND_EMAIL)
        $newAction.Server  = $Server
        $newAction.To      = $To
        $newAction.Cc      = $Cc
        $newAction.From    = $From
        $newAction.Subject = $Subject
        $newAction.Body    = $Body
    }

    AddTimeTrigger([System.DateTime]$StartTime,[System.DateTime]$EndTime)
    {
		if(([datetime]::Compare($StartTime,$EndTime)) -eq 1){
			Write-Warning "$StartTime is after $EndTime....please enter different start- and endtimes!"
			break
		}
        $newTrigger               = $this.TaskObject.Triggers.Create([TriggerType]::TASK_TRIGGER_TIME)
        $newTrigger.StartBoundary = (Get-Date $StartTime -Format "yyyy-MM-ddTHH:mm:ss")
        $newTrigger.EndBoundary   = (Get-Date $EndTime -Format "yyyy-MM-ddTHH:mm:ss")
    }

    CreateFolder($ParentFolder,$FolderName)
    {
        $Parent=$this.objScheduler.GetFolder($ParentFolder)
        $Parent.CreateFolder($FolderName)
    }

    DeleteExpiredTaskAfterXHours([System.Int16]$Hours)
    {
        $this.TaskObject.Settings.DeleteExpiredTaskAfter = "PT$($Hours)H"
    }

    DeleteFolder($ParentFolder,$FolderName)
    {
        $Parent=$this.objScheduler.GetFolder($ParentFolder)
        $Parent.DeleteFolder($FolderName,$null)
    }

    DeleteTask()
    {
        $this.TaskFolder.DeleteTask($this.TaskObject.Name,$null)
    }

    DeleteTask($TaskName)
    {
        $this.TaskFolder.DeleteTask($TaskName,$null)
    }


    DeleteTask($TaskFolder,$TaskName)
    {
        $Parent=$this.objScheduler.GetFolder($TaskFolder)
        $Parent.DeleteTask($TaskName,$null)
    }

    Disable()
    {
        $this.TaskObject.Enabled = $false
    }

    Enable()
    {
        $this.TaskObject.Enabled = $true
    }

    [System.Object]GetFolder($Path)
    {
        $this.TaskFolder=$this.objScheduler.GetFolder($Path)
        return $this.TaskFolder
    }

    [System.Object]GetFolders($ParentFolder)
    {
        $Parent=$this.objScheduler.GetFolder($ParentFolder)
        return $Parent.GetFolders(0) | Select-Object Name
    }

    [String]GetState()
    {
        return [Enum]::GetName([State], $this.TaskObject.State)
    }

    GetTask($TaskName)
    {
        #$Parent=$this.objScheduler.GetFolder($ParentFolder)
        $this.TaskObject=$this.TaskFolder.GetTask($TaskName)
    }

    GetTask($ParentFolder,$TaskName)
    {
        $Parent=$this.objScheduler.GetFolder($ParentFolder)
        $this.TaskObject=$Parent.GetTask($TaskName)
    }

    [System.Object]GetTasks($ParentFolder)
    {
        $Parent=$this.objScheduler.GetFolder($ParentFolder)
        return $Parent.GetTasks(0) | Select-Object Name
    }

    Hide()
    {
        $this.TaskObject.Settings.Hidden=$true
    }

    RegisterTaskAsSystem($DisplayName)
    {
        #$this.RegisteredTask=$this.TaskFolder.RegisterTaskDefinition($this.DisplayName,$this.TaskObject,$TaskCreation,"SYSTEM",$null,$LogonType)
        #$Path=Split-Path $this.TaskObject.Path
        $objFolder=$this.GetFolder("\")
        $this.TaskObject=$objFolder.RegisterTaskDefinition($DisplayName,$this.TaskObject,[taskcreation]::TASK_CREATE_OR_UPDATE,"SYSTEM",$null,[logontype]::TASK_LOGON_SERVICE_ACCOUNT)
    }

    RegisterTaskAsSystem($DisplayName,$FolderPath)
    {
        #$this.RegisteredTask=$this.TaskFolder.RegisterTaskDefinition($this.DisplayName,$this.TaskObject,$TaskCreation,"SYSTEM",$null,$LogonType)
        #$Path=Split-Path $this.TaskObject.Path
        $objFolder=$this.GetFolder($FolderPath)
        $this.TaskObject=$objFolder.RegisterTaskDefinition($DisplayName,$this.TaskObject,[taskcreation]::TASK_CREATE_OR_UPDATE,"SYSTEM",$null,[logontype]::TASK_LOGON_SERVICE_ACCOUNT)

    }

    RegisterTask([System.String]$DisplayName)
    {
        $objFolder=$this.GetFolder("\")
        $this.TaskObject=$objFolder.RegisterTaskDefinition($DisplayName,$this.TaskObject,[TaskCreation]::TASK_CREATE_OR_UPDATE,$this.CredentialObject.UserName,$this.CredentialObject.GetNetworkCredential().Password,[LogonType]::TASK_LOGON_PASSWORD)
    }

    RegisterTask([System.String]$DisplayName,$FolderPath)
    {
        $objFolder=$this.GetFolder($FolderPath)
        $this.TaskObject=$objFolder.RegisterTaskDefinition($DisplayName,$this.TaskObject,[TaskCreation]::TASK_CREATE_OR_UPDATE,$this.CredentialObject.UserName,$this.CredentialObject.GetNetworkCredential().Password,[LogonType]::TASK_LOGON_PASSWORD)
    }

    Run()
    {
        $this.TaskObject.Run(0)
    }
    
    EnableRunWithHighestPrivilege([System.Boolean]$RunWithHighestPrivilege)
    {
        if($RunWithHighestPrivilege){
            $this.TaskObject.Principal.RunLevel = 1
        }
        else{
            $this.TaskObject.Principal.RunLevel = 0
        }
    }

    SetCompatibilityLevel([System.Int16]$CompatibilityLevel)
    {
        $this.TaskObject.Settings.Compatibility = $CompatibilityLevel
    }

    SetExecutionTimeLimit([System.Int16]$Hours)
    {
        $this.TaskObject.Settings.ExecutionTimeLimit = "PT$($Hours)H"
    }

    SetLogonType([System.Int16]$LogonType)
    {
        $this.TaskObject.Principal.LogonType=$LogonType
    }

    SetPriority([System.Int16]$Priority)
    {
        $this.TaskObject.Settings.Priority = $Priority
    }

    Stop()
    {
        $this.TaskObject.Stop(0)
    }
#endregion methods
}
