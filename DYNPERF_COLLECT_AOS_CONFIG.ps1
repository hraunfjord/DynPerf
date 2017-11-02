# hraunfjord@gmail.com - powershell script for VBS DynPerf job DYNPERF_COLLECT_AOS_CONFIG

$SQLServer    = "DAXDBServer" # Dynamics AX Database Server
$SQLDBName    = "DynamicsAX"
$DPASQLServer = "DynPerfDBServer" # Dynamics Perf Database Server
$DPASQLDBName = "DynamicsPerf"

$NowDate = Get-Date
#Write-Host "BEGIN script" $NowDate

# Connect to DAXDB Server 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True;"
$SqlConnection.Open()
# Quit if the SQL connection didn't open properly.
if ($sqlConnection.State -ne [Data.ConnectionState]::Open) {
    "Connection to AXDB is not open."
    Exit
}

# Connect to DynamicsPerf Server
$DPASqlConnection = New-Object System.Data.SqlClient.SqlConnection
$DPASqlConnection.ConnectionString = "Server = $DPASQLServer; Database = $DPASQLDBName; Integrated Security = True;"
$DPASqlConnection.Open()
# Quit if the SQL connection didn't open properly.
if ($DPAsqlConnection.State -ne [Data.ConnectionState]::Open) {
    "Connection to DynPerfDB is not open."
    Exit
}

# Get AOS server instances from DAXDB
$SqlQuery = "select substring(serverid, charindex('@',serverid, 0)+1, len(serverid)-charindex('@',serverid, 0)) AS Servers `
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
    $DPAEvents = (Get-EventLog -LogName Application -Computer $strAOS -After $currDate14) | 
            select EventID, TimeGenerated, EntryType, Source, InstanceId, Message, MachineName | 
            where {($_.EntryType -eq"Error") -or ($_.EntryType -eq"Warning") -or 
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

