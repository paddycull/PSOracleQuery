<#
.SYNOPSIS
    PowerShell script to query an Oracle database, using Windows Authentication. The user can pass a query directly, or a SQL file to run against a database.

.DESCRIPTION
    This script allows the user to query an Oracle database. the primary features are;
        - You can pass a query directly or a sql file with queries in it.
        - It uses Windows Authentication, so no credentials are requried.
        - It uses the Oracle.ManagedDataAccess.dll that is already on the target, so there is no dependencies. 
        - Can pass multiple queries at once, either directly through the Query parameter, or in a SQL file. Result set for over 1 query is a PSObject.
        - If there's only one query, it returns the PSObject array as is. If it's multiple queries, it returns a special PSObject array, with a Query and ResultSet patameters.

.EXAMPLE
    Invoke-OracleQuery -HostName HostServer1 -ServiceName PATPDB1 -Query "select username, account_status from dba_users;"
        This example runs the given query against the PATPDB1 database on HostServer1

.EXAMPLE
    Invoke-OracleQuery -HostName HostServer1 -ServiceName PATPDB1 -SqlFile C:\example\oracle.sql
        This runs the oracle.sql file against the database.

.NOTES
    Author: Patrick Cull
    Date: 2020-02-03

    Update 2020-02-25
    - Now compatible with Oracle 11g
    - Now uses full connection string instead of EZConnect string to connect to databases.

    Update 2020-03-12
    - Now defaults to localhost to use the dll files, if the required files are available. This makes it faster and also allows the function to query db's on Linux servers.
    - If ODP.NET is not installed, function tries to use the remote computers ODP.NET to query the database instead.

#>
function Invoke-OracleQuery {
    [Cmdletbinding()]
    param(   
        #Server the database is on 
        [Parameter(Mandatory)]
        [string] $HostName,

        #The database service name to query
        [Parameter(Mandatory)]
        [string] $ServiceName,
        
        #Credential to connect to the target computer. Defaults to current credential if not passed.
        [System.Management.Automation.PSCredential] $TargetCredential=[System.Management.Automation.PSCredential]::Empty, 

        #User credential to connect to the database with. Defaults to current windows user as sysdba if not passed.
        [System.Management.Automation.PSCredential] $DatabaseCredential,

        # Query to run
        [string] $Query,
        
        #Sql file to run against the database.
        [string] $SqlFile,

        #Connect as sysdba
        [switch]$AsSysdba
    )

    if(!$Query -and !$SqlFile){
        Throw "Please pass either a Query or a SqlFile to run against the target database."
    }

    if($Query -and $SqlFile){
        Throw "Cannot use the Query and SqlFile parameters together, use one or the other."
    }

    if($SqlFile) {
        $Query = (Get-Content $SqlFile | Out-String).Trim()
    }

    #If the final character in the query is a slash or semicolon, we remove it, as the oracleManagedDataAccess module does not allow these characters at the end of a query
    if($Query[-1] -eq "/" -or $Query[-1] -eq ';') {
        $Query = $Query -replace ".$"
    }

    #Split the queries into individual queries, either on semicolon outside single quotes, or a / on it's own line - i.e. the query terminators in Oracle.
    $OracleQueries = $Query -Split ";+(?=(?:[^\']*\'[^\']*\')*[^\']*$)"
    $OracleQueries = $OracleQueries.Split([string[]]"`r`n/", [StringSplitOptions]::None)


    #Check if the localhost has ODP.NET installed. We use the local files if it does.
    $LocalOracleHome = (Get-ItemProperty "HKLM:SOFTWARE\ORACLE\KEY_ora*" -Name ORACLE_HOME).ORACLE_HOME | Select-Object -First 1

    #Default location of the older and newer versions of the ODP.NET dll files.
    $NewerDllPath = "$LocalOracleHome\ODP.NET\managed\common\Oracle.ManagedDataAccess.dll"
    $OlderDllPath = "$LocalOracleHome\ODP.NET\bin\2.x\Oracle.DataAccess.dll"
    
    if((Test-Path $NewerDllPath) -or (Test-Path $OlderDllPath)) {
        $ComputerWithODP = "localhost"
        Write-Verbose "ODP.NET installed locally - Running query from localhost."
        $OracleHome = $LocalOracleHome
    }

    else {    
        Write-Verbose "ODP.NET is not installed locally. Running query on $HostName."

        $OracleHome = Invoke-Command -ComputerName $HostName -ScriptBlock {
            (Get-ItemProperty "HKLM:SOFTWARE\ORACLE\KEY_OraDB*" -Name ORACLE_HOME).ORACLE_HOME | Select-Object -First 1
        }
        if(!$OracleHome) {
            Throw "No Oracle Home on $HostName"
        }

        Write-Verbose "Using the $OracleHome Oracle home on $HostName to get the dll files."

        #Default location of the older and newer versions of the ODP.NET dll files.
        $NewerDllPath = "$OracleHome\ODP.NET\managed\common\Oracle.ManagedDataAccess.dll"
        $OlderDllPath = "$OracleHome\ODP.NET\bin\2.x\Oracle.DataAccess.dll"

        $NetworkNewerDllPath = "\\$HostName\$NewerDllPath" -replace ':', '$'
        $NetworkOlderDllPath = "\\$HostName\$OlderDllPath" -replace ':', '$'

        $NewDllExists = Test-Path $NetworkNewerDllPath
        $OldDllExists = Test-Path $NetworkOlderDllPath

        #Need to unblock the dll files so it can be imported on the target. DLL's may be blocked by Windows if the home was copied from a different server, i.e. gold images.   
        if($NewDllExists) {
            Unblock-File $NetworkNewerDllPath
        }
        if($OldDllExists) {
            Unblock-File $NetworkOlderDllPath
        } 

        #If neither of the dll files are on the target we can't query the database.
        if(!$NewDllExists -and !$OldDllExists) {
            Throw "No Oracle.ManagedDataAccess.dll or Oracle.DataAccess.dll found in the Oracle home $OracleHome of $HostName."
        }     
        
        $ComputerWithODP = $HostName

    }#End else

    #The query block to run on localhost or the target host, depending on if ODP is installed locally.
    $QueryScriptBlock = { 
        param($HostName, $ServiceName, $OracleQueries, $DatabaseCredential, $AsSysdba, $OlderDllPath, $NewerDllPath, $VerboseSetting)

        $VerbosePreference = $VerboseSetting

        try {
            #Try to import the newer dll first.
            Add-Type -Path $NewerDllPath
            $OracleDllLocation = $NewerDllPath
        }
        catch {
            #If it fails we try to import the older version instead.
            Write-Verbose $_.Exception.Message
            Write-Verbose "Attempting to add the older DLL type."         
            try {
                Add-Type -Path $OlderDllPath
                $OracleDllLocation = $OlderDllPath
            }
            catch {
                Write-Error $_.Exception.Message
                Throw "Issue adding Oracle.ManagedDataAccess.dll from the Oracle Home."      
            }
        }
        Write-Verbose "Using $OracleDllLocation to access database."
        
        #Set the DLL type depending on the file name that was loaded.
        if($OracleDllLocation -like '*ManagedDataAccess*') {
            $DllType = "ManagedDataAccess"
        }
        else {
            $DllType = "DataAccess"
        }
        
        $AllResults = @()
        $QueryCount = $OracleQueries.Count

        #Get IP address for a better Connection String. Using "localhost" in connect descriptor caused queries on some servers to fail, using the IP address is better.
        $TargetIpAddress = Test-Connection -ComputerName $HostName -Count 1  | Select-Object -ExpandProperty IPV4Address | Select-Object -ExpandProperty IPAddressToString

        #If a credential is passed, we use that instead of using windows auth
        if($DatabaseCredential) {
            $DbAccountUsername = $DatabaseCredential.UserName
            $DbAccountPassword = $DatabaseCredential.GetNetworkCredential().password

            if($AsSysdba) {
                $ConnectionString = "User Id=$DbAccountUsername;Password=$DbAccountPassword;DBA Privilege=SYSDBA;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$TargetIpAddress)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=$ServiceName)))"
            }
            else {
                $ConnectionString = "User Id=$DbAccountUsername;Password=$DbAccountPassword;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$TargetIpAddress)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=$ServiceName)))"
            }
        }

        else {
            $ConnectionString = "User Id=/;DBA Privilege=SYSDBA;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$TargetIpAddress)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=$ServiceName)))"
        }

        Write-Verbose "Connection String: $ConnectionString"


        :Queries foreach($Query in $OracleQueries) {
            #Clear any empty lines and uneeded whitespace.
            $Query = $Query.Trim()

            $QueryResult = @()
            try { 
                $Connection = New-Object Oracle.$DllType.Client.OracleConnection($ConnectionString)
                $cmd=$Connection.CreateCommand()

                $cmd.CommandText= $Query

                $da=New-Object Oracle.$DllType.Client.OracleDataAdapter($cmd);
                $QueryResult=New-Object System.Data.DataTable
                [void]$da.fill($QueryResult)
                #This expands the results set to human readable form - without this it just shows "System.Data.DataRow" for each resultset.
                $QueryResult = $QueryResult | Select-Object -Property * -ExcludeProperty RowError,RowState,table,ItemArray,HasErrors
            } 

            catch {
                $ErrorMessage = $_.Exception.Message.ToString()

                #If the error message is a connection error, we set connectionError to true, so we can break out of the query loop.
                if(($ErrorMessage -like '*ORA-12154: TNS:could not resolve the connect identifier specified*') -or ($ErrorMessage -like "*ORA-01017: invalid username/password; logon denied*")){
                    $ConnectionError = $True
                }

                #If the error message is a normal ORA- Error, we return that;
                if($ErrorMessage -like '*ORA-*'){
                    $QueryResult = "`"ORA-" + ($ErrorMessage -split 'ORA-')[1]
                }
                #Otherwise return it, unchanged.
                else {
                    $QueryResult = $ErrorMessage
                }
            } 
            finally {
                if ($Connection.State -eq 'Open') { $Connection.close() }
            }   

            #If nothing is returned, it means there were no errors or tables returned. So we give it a value based on the query.
            if(!$QueryResult) {

                #Split the query up so we can access the words separately.
                $QueryWords = $Query -split " "
                
                #The first word of the command is the verb - e.g. create, drop, select.
                $QueryVerb = $QueryWords[0] 
                
                #The subject of the query is the second word -e.g. "DROP TABLE.." the subject is the table. The Get-Culture function lets us capitalize the first letter.
                $QuerySubjectType = (Get-Culture).TextInfo.ToTitleCase($QueryWords[1])

                #If the verb is SELECT and there is no query result, that means the result set is empty.
                if($QueryVerb -like 'SELECT') {
                    $QueryResult = "no rows selected"
                }
                
                #if it wasn't a select, and there was nothing returned, it means it was successful. So we create a result message.
                else {
                    if($QueryVerb -like 'DROP'){
                        $QueryVerbPastTense = "dropped"
                    }
                    elseif($QueryVerb -like 'CREATE'){
                        $QueryVerbPastTense = "created"
                    }
                    elseif($QueryVerb -like 'TRUNCATE'){
                        $QueryVerbPastTense = "truncated"
                    }
                    elseif($QueryVerb -like 'DELETE'){
                        $QueryVerbPastTense = "deleted"
                    }
                    elseif($QueryVerb -like 'ALTER'){
                        $QueryVerbPastTense = "altered"
                    }
                    elseif($QueryVerb -like 'EXECUTE'){
                        $QueryVerbPastTense = "executed"
                    }
                    elseif($QueryVerb -like 'GRANT'){
                        $QueryVerbPastTense = "granted"
                    }
                    #If it's not one of the above verbs, we default to succeeded.
                    else{
                        $QueryVerbPastTense = "succeeded"
                    }

                    $QueryResult = "$QuerySubjectType $QueryVerbPastTense."
                }
            }

            #If more than one query, we return an object array instead og just the resultset.
            if($QueryCount -gt 1) { 
                $QueryObject = New-Object PSObject     
                Add-Member -memberType NoteProperty -InputObject $QueryObject -Name Query -Value $Query  
                Add-Member -memberType NoteProperty -InputObject $QueryObject -Name ResultSet -Value $QueryResult     
                
                $AllResults += $QueryObject 

                #If there was a connection error, we do not continue with the remaining queries.
                if($ConnectionError){
                    Write-Host "[ERROR] : Issue connecting to the database. Not continuing with remaining queries to prevent lockouts." -ForegroundColor Red
                    break Queries
                }
            }
            #Otherwise we just return the resultset
            else {
                $AllResults = $QueryResult
            }
        }#End foreachloop
        
        return $AllResults
    }#EndScriptBlock

    #Need to pass verbose setting into the scriptblock so they are shown on both localhost and target host
    $VerboseSetting = $VerbosePreference

    #If ODP is installed locally, run it from localhost. Otherwise run it on the target host
    if($ComputerWithODP -eq "localhost") {
        $Output = Invoke-Command -ArgumentList $HostName, $ServiceName, $OracleQueries, $DatabaseCredential, $AsSysdba, $OlderDllPath, $NewerDllPath, $VerboseSetting -ScriptBlock $QueryScriptBlock
    }
    else {
        $Output = Invoke-Command -ComputerName $HostName -Credential $TargetCredential -HideComputerName -ArgumentList $HostName, $ServiceName, $OracleQueries, $DatabaseCredential, $AsSysdba, $OlderDllPath, $NewerDllPath, $VerboseSetting -ScriptBlock $QueryScriptBlock
    }

    #If the Output type is an object, it means a proper query resultset has been returned, so we remove the PSComputerName, RunspaceId and PSShowComputerName that gets added to objects by PowerShell after the Invoke-Command
    if($Output.GetType().Name -like '*Object*') {
        $Output | Select-Object -Property * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName
    }
    #Otherwise we leave the output unchanged. Using the select-object from above on a non object returns the length of the variable, which is not what we want.
    else {
        $Output
    }
}

#This ArgumentCompleter runs 'lsnrctl status' on the target and extracts the database services, for suggestions for the ServiceName parameter.
Register-ArgumentCompleter -CommandName Invoke-OracleQuery -ParameterName ServiceName -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $HostName = $fakeBoundParameter.HostName

    $ServiceNames = Invoke-Command -ComputerName $HostName -ScriptBlock {
        $ListenerOutput = lsnrctl status
        
        #Extract any string that starts with 'Service' up to the end '"'
        $ServiceNames = ([regex]::Matches($ListenerOutput, 'Service([^/)]+")') |ForEach-Object { $_.Groups[1].Value }).Trim()

        #remove extra uneeded strings
        $ServiceNames = ($ServiceNames -replace "s Summary... Service ") -replace '"'
          
        return $ServiceNames
    }
    
    #Remove XDB service, CLRextProc and any PDB GUID services (which is 32 chars long)
    $DatabaseServices = $ServiceNames.ToUpper() | Where-Object{$_ -NotLike '*XDB' -and $_ -ne 'CLRExtProc' -and $_.Length -ne 32}

    return $DatabaseServices
}