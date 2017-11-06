/******   Recreate JOB DYNPERF_COLLECT_AOS_CONFIG with PowerShell script   
		Author: hraunfjord@gmail.com
		When script has been run, go to the job and configure DB servers and DB name
		(As with VBS script)                                                         
		Check Powershell script security : https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1
		Check that the the SQLAgent account has privileges to access AOS the Registry and EventLog 
		(or configure SQL credentials and Proxy user to run this job)
******/


USE [msdb]
GO

/****** Object:  Job [DYNPERF_COLLECT_AOS_CONFIG]    Script Date: 09/19/2015 07:23:26 ******/
IF  EXISTS (SELECT job_id FROM msdb.dbo.sysjobs_view WHERE name = N'DYNPERF_COLLECT_AOS_CONFIG')
EXEC msdb.dbo.sp_delete_job @job_name=N'DYNPERF_COLLECT_AOS_CONFIG', @delete_unused_schedule=1
GO

USE [msdb]
GO

IF (SELECT SERVERPROPERTY('PRODUCTVERSION')) >= '13.'
BEGIN

/****** Object:  Job [DYNPERF_COLLECT_AOS_CONFIG]    Script Date: 09/19/2015 07:23:26 ******/
BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0
/****** Object:  JobCategory [[Uncategorized (Local)]]]    Script Date: 09/19/2015 07:23:26 ******/
IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'DYNPERF_COLLECT_AOS_CONFIG', 
		@enabled=0, 
		@notify_level_eventlog=0, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'No description available.', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'sa', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
/****** Object:  Step [COLLECT AOS CONFIG]    Script Date: 09/19/2015 07:23:26 ******/
EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'COLLECT AOS CONFIG', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=0, 
		@retry_interval=0, 
		@os_run_priority=0, @subsystem=N'PowerShell', 
	@command=N'# hraunfjord@gmail.com - powershell script for VBS DynPerf job DYNPERF_COLLECT_AOS_CONFIG

$SQLServer    = "DAXDBServer" # Dynamics AX Database Server
$SQLDBName    = "DynamicsDB"
$DPASQLServer = "DynPerfDBServer" # Dynamics Perf Database Server
$DPASQLDBName = "DynamicsPerf"

$NowDate = Get-Date
#Write-Host "BEGIN script" $NowDate

# Connect to DAXDB Server 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True;"
$SqlConnection.Open()
# Quit if the SQL connection didnt open properly.
if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
    "Connection to AXDB is not open."
    Exit
}

# Connect to DynamicsPerf Server
$DPASqlConnection = New-Object System.Data.SqlClient.SqlConnection
$DPASqlConnection.ConnectionString = "Server = $DPASQLServer; Database = $DPASQLDBName; Integrated Security = True;"
$DPASqlConnection.Open()
# Quit if the SQL connection didnt open properly.
if ($DPAsqlConnection.State -ne [Data.ConnectionState]::Open) {
    "Connection to DynPerfDB is not open."
    Exit
}

# Get AOS server instances from DAXDB
$SqlQuery = "select substring(serverid, charindex(''@'',serverid, 0)+1, len(serverid)-charindex(''@'',serverid, 0)) AS Servers `
             from sysserverconfig"
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
$SqlAdapter.SelectCommand = $SqlCmd
$DataSet = New-Object System.Data.DataSet
$SqlAdapter.Fill($DataSet) | Out-Null

$currDate = Get-Date
$currDate14 = $currDate.AddDays(-14)

function Is-Numeric ($Value) {
    return $Value -match "^[\d\.]+$"
}

function ClearAOSRegLog([Data.SqlClient.SqlConnection]$_DPASqlConnection) {
    $DPAsqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $DPAsqlCmd.Connection = $DPASqlConnection
    $DPAsqlCmd.CommandText = "TRUNCATE TABLE dbo.AOS_REGISTRY"
    $InsertedID = $DPAsqlCmd.ExecuteScalar()
}


function AOSreg([String]$strAOS){
    #Write-Host "   BEGIN Collect Registry info from" $strAOS

    # start by truncating AOS_REGISTRY table in DynamicsPerf DB    
    $clrLogStatus = ClearAOSRegLog($DPASqlConnection)

    # construct TSQL Insert string and parameters
    $DPAsqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $DPAsqlCmd.Connection = $DPASqlConnection
    $DPASqlCmd.CommandText = "SET NOCOUNT ON; `
                              Insert into AOS_REGISTRY (Server_Name, AX_MAJOR_VERSION, AOS_INSTANCE_NAME, AX_BUILD_NUMBER, `
                              AOS_CONFIGURATION_NAME, IS_CONFIGURATION_ACTIVE, SETTING_NAME, SETTING_VALUE ) `
                              Values (@theServerName, @theAxMajorVersion, @theAOSInstanceName, @theAxBuildNumber, `
                              @theAOSConfiguationName, @theIsConfigurationActive, @theSettingName, @theSettingValue); " 
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheServerName",[Data.SQLDBType]::NVarChar,255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theAxMajorVersion",[Data.SQLDBType]::NVarChar, 255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theAOSInstanceName",[Data.SQLDBType]::NVarChar, 255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theAxBuildNumber",[Data.SQLDBType]::NVarChar, 25))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theAOSConfiguationName",[Data.SQLDBType]::NVarChar, 255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theIsConfigurationActive",[Data.SQLDBType]::NVarChar, 1))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theSettingName",[Data.SQLDBType]::NVarChar, 255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@theSettingValue",[Data.SQLDBType]::NVarChar))) | Out-Null             


    # Branch of the Registry  
    $Branch="LocalMachine"
 
    # Main Sub Branch you need to open  
    $SubBranch="SYSTEM\\CurrentControlSet\\Services\\Dynamics Server"  
 
    $registry=[microsoft.win32.registrykey]::OpenRemoteBaseKey("Localmachine",$strAOS)  
    $registrykey=$registry.OpenSubKey($Subbranch)  
    $DAXKeys=$registrykey.GetSubKeyNames()  


    # Drill through each key from the list and pull out the value of  
    # “DisplayName” – Write to the Host console the name of the computer  
    # with the application beside it 

    $SERVER_NAME = $strAOS
    $AX_MAJOR_Version = ""       
    $AOS_INSTANCE_NAME = ""      
    $AX_BUILD_NUMBER = ""        
    $AOS_CONFIGURATION_NAME = "" 
    $IS_CONFIGURATION_ACTIVE = ""
    $SETTING_NAME = ""
    $SETTING_VALUE = ""
    $SETTING_ACTIVE_CONFIG = ""  
 
    # Open Versions of DAX
    Foreach ($StrVersion  in $DAXkeys)  
    {  
        # only numeric versions - not "Performance"  
        if (Is-Numeric($StrVersion[0])){
            $AX_MAJOR_Version = $StrVersion 
            #$exactkey=$key  
            $StrVersionSubKey=$SubBranch+"\\"+$StrVersion   
            #Write-host "Working on" $StrVersionSubKey

            $StrVersionReg=$registry.OpenSubKey($StrVersionSubKey)  
            $StrVersionValue  = $StrVersionReg.GetSubKeyNames()  

            # Open Instance for this version of DAX  
            Foreach ($StrInstance in $StrVersionValue) {
                $StrInstanceSubKey=$StrVersionSubKey+"\\"+$StrInstance
                $StrInstanceReg=$registry.OpenSubKey($StrInstanceSubKey)  
                $StrInstanceValue = $StrInstanceReg.GetSubKeyNames()  
                $AOS_INSTANCE_NAME     = $StrInstanceReg.GetValue("InstanceName")
                $AX_BUILD_NUMBER       = $StrInstanceReg.GetValue("ProductVersion")
                $SETTING_ACTIVE_CONFIG = $StrInstanceReg.GetValue("Current")
            
                # open each configuration
                Foreach ($StrConfig in $StrInstanceValue){
                    $StrConfigSubKey=$StrInstanceSubKey+"\\"+$StrConfig
                    $StrConfigReg=$registry.OpenSubKey($StrConfigSubKey) 
                    $AOS_CONFIGURATION_NAME = $StrConfig
                    if($StrConfig -eq $SETTING_ACTIVE_CONFIG) {$IS_CONFIGURATION_ACTIVE = "Y"} ELSE {$IS_CONFIGURATION_ACTIVE = "N"}  
                    # Get all key names in this branch
                    $StrConfigKeys = $StrConfigReg.GetValueNames()

                    # open each key in the config and write to SQL
                    Foreach ($StrValueName in $StrConfigKeys){                                
                        $SETTING_NAME = $StrValueName
                        $SETTING_VALUE = $StrConfigReg.GetValue($StrValueName)
                        #Write-Host $SERVER_NAME $AX_MAJOR_Version $AOS_INSTANCE_NAME $AX_BUILD_NUMBER $AOS_CONFIGURATION_NAME $IS_CONFIGURATION_ACTIVE $SETTING_NAME $SETTING_VALUE
                        $DPASqlCmd.Parameters[0].Value = $SERVER_NAME
                        $DPASqlCmd.Parameters[1].Value = $AX_MAJOR_Version
                        $DPASqlCmd.Parameters[2].Value = $AOS_INSTANCE_NAME
                        $DPASqlCmd.Parameters[3].Value = $AX_BUILD_NUMBER
                        $DPASqlCmd.Parameters[4].Value = $AOS_CONFIGURATION_NAME
                        $DPASqlCmd.Parameters[5].Value = $IS_CONFIGURATION_ACTIVE
                        $DPASqlCmd.Parameters[6].Value = $SETTING_NAME
                        $DPASqlCmd.Parameters[7].Value = $SETTING_VALUE
                        $InsertedID = $DPASqlCmd.ExecuteScalar()


                    } #StrValueName

                } # StrConfig

            } # StrInstance
        } # Major version
    } # Foreach
    #Write-Host "   END   Collect Registry info from" $strAOS

}

function ClearAOSEventLog([Data.SqlClient.SqlConnection]$_DPASqlConnection) {
    $DPAsqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $DPAsqlCmd.Connection = $DPASqlConnection
    $DPAsqlCmd.CommandText = "TRUNCATE TABLE dbo.AOS_EVENTLOG"
    $InsertedID = $DPAsqlCmd.ExecuteScalar()
}


function AOSevt([String]$strAOS){
    #Write-host "   BEGIN - Collect Events from " $strAOS

    # start by truncating AOS_EventLog table in DynamicsPerf DB    
    $clrLogStatus = ClearAOSEventLog($DPASqlConnection)

    # construct TSQL Insert string and parameters
    $DPAsqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $DPAsqlCmd.Connection = $DPASqlConnection
    $DPASqlCmd.CommandText = "SET NOCOUNT ON; `
                              Insert into AOS_EVENTLOG (Time_Written, Server_Name, Event_Code, Event_Type, Message, Source_Name ) `
                              Values (@theDate, @theServerName, @theEventCode, @theEventType, @theMessage, @theSourceName); " 
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheDate",[Data.SQLDBType]::DateTime))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheServerName",[Data.SQLDBType]::NVarCHar, 255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheEventcode",[Data.SQLDBType]::Int))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheEventType",[Data.SQLDBType]::NVarCHar, 255))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheMessage",[Data.SQLDBType]::NVarCHar))) | Out-Null             
    $DPASqlCmd.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@TheSourceName",[Data.SQLDBType]::NVarCHar, 255))) | Out-Null             


    # Get events from Application Log and filter
    $DPAEvents = (Get-EventLog -LogName Application -Computer $strAOS -After $currDate14) | `
            select EventID, TimeGenerated, EntryType, Source, InstanceId, Message, MachineName | `
            where {($_.EntryType -eq"Error") -or ($_.EntryType -eq"Warning") -or `
                   (($_.EntryType -eq"Information") -and ($_.EventID -eq 149)) } 
    
    #Write-host "   END   - Collect Events from " $strAOS
    #Write-host "   BEGIN - Insert Events into SQL" $DPASQLServer $DPASQLDBName
    
    # insert events into SQL
    $EventCount = 0
    foreach ($item in $DPAEvents) {
        $DPASqlCmd.Parameters[0].Value = $item.TimeGenerated
        $DPASqlCmd.Parameters[1].Value = $item.MachineName
        $DPASqlCmd.Parameters[2].Value = $item.EventID
        $DPASqlCmd.Parameters[3].Value = $item.EntryType
        $DPASqlCmd.Parameters[4].Value = $item.Message
        $DPASqlCmd.Parameters[5].Value = $item.Source
        $InsertedID = $DPASqlCmd.ExecuteScalar()
        $EventCount += 1
    }
    #Write-host "   END   - Insert Events into SQL" $DPASQLServer $DPASQLDBName "Total" $EventCount "events"

}  #AOSEvt

# main program
# Loop throgh each AOS instance
foreach ($item in $DataSet) {

    # exctact Server name
    $Server = $item.Tables[0].servers
    #Write-Host "AOS server: " $Server 
        
    # Get Events and write to DB
    AOSevt($Server)

    # Get Registry info and write to DB
    AOSreg($Server)
    
}

$NowDate = Get-Date
#Write-Host "END   script" $NowDate
# end close DB connection
$SqlConnection.Close()
$DPASqlConnection.Close()



', 
		@database_name=N'PowerShell', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'DAILY', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20150919, 
		@active_end_date=99991231, 
		@active_start_time=60000, 
		@active_end_time=235959

IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:
END