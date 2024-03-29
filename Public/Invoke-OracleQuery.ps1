<#
.SYNOPSIS
    PowerShell script to query an Oracle database, using Windows Authentication. The user can pass a query directly, or a SQL file to run against a database.

.DESCRIPTION
    This script allows the user to query an Oracle database. the primary features are;
        - You can pass a query directly or a sql file with queries in it.
        - It uses Windows Authentication, so no credentials are requried if you are part of the ORA_DBA group.
        - It uses the Oracle.ManagedDataAccess.dll to access the database. 
        - Can pass multiple queries at once, either directly through the Query parameter, or in a SQL file.

.EXAMPLE
    Invoke-v1OracleQuery -HostName HostServer1 -ServiceName PATPDB1 -Query "select username, account_status from dba_users;"
        This example runs the given query against the PATPDB1 database on HostServer1

.EXAMPLE
    Invoke-v1OracleQuery -HostName HostServer1 -ServiceName PATPDB1 -SqlFile C:\example\oracle.sql
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

    Update 2020-09-23
    - Query can now contain PL/SQL blocks.
    - Automatically remove comment lines - they do not work with ODP.NET commands.
    - Added ExitOnError switch so the user can decide to stop the execution when a query encounters an error. Default behaviour is to continue with remaining queries.

    Update 2021-02-05
    - Query returns a "Query" and "ResultSet" even if there is only one query. Previously single queries would only return "ResultSet"
    - Function now uses an included Oracle.ManagedDataAccess.dll to access the remote database.
        - This speeds up the query massively and removes all dependencies, including remoting to the Host to run queries if ODP was not installed locally.

    Update 2021-11-29
     - Added -ResultSetOnly switch to optionally return only the result set and not the query
     - Added PortNumber parameter

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

        #User credential to connect to the database with. Defaults to current windows user as sysdba if not passed.
        [System.Management.Automation.PSCredential] $DatabaseCredential,

        # Query to run
        [string[]] $Query,
        
        #Sql file to run against the database.
        [string] $SqlFile,

        #Port number to use for the connection string.
        [int] $PortNumber = 1521,

        #Connect as sysdba
        [switch]$AsSysdba,

        #Optionally stop as soon as a query encounters an error. Default behaviour is to continue with the remaining queries after an error occcurs.
        [switch]$ExitOnError,

        #By default, the function will return the Query and the ResultSet - which is useful for multiple queries. This switch will return only the ResultSet - which can be useful for single queries.
        [switch] $ResultSetOnly
    )

    if(!$Query -and !$SqlFile){
        Throw "Please pass either a Query or a SqlFile to run against the target database."
    }

    if($Query -and $SqlFile){
        Throw "Cannot use the Query and SqlFile parameters together, use one or the other."
    }

    if($SqlFile) {
        $Query = Get-Content $SqlFile -ErrorAction Stop
    }
   

    ##################################
    # Start  Query Prep
    ##################################
    #Split the input command up. We need to account for PL/SQL Declare/Begin/End blocks
        #Remove comments, they don't work when using ODP.NET to run the query.
        $NoCommentQuery = ($Query | Where-Object {$_ -notlike 'set *' -and $_ -notlike 'PROMPT*' -and $_ -notlike 'column*' -and $_ -notlike 'compute*' -and $_ -notlike 'Rem *'}) | Out-String

        #Get the location of Declare/Begin/End blocks as well as the locations of the end command characters - i.e. a ; outside single quotes or a / on it's own line
        $BeginEndBlocks = $NoCommentQuery | Select-String 'DECLARE[^/]+/|DECLARE[^/]+END;|BEGIN[^/]+/|BEGIN[^/]+END;' -AllMatches | ForEach-Object { $_.Matches } | Sort-Object Index -Descending
        $EndCommandCharacters = $NoCommentQuery | Select-String ";+(?=(?:[^\']*\'[^\']*\')*[^\']*$)|`r`n/" -AllMatches | ForEach-Object { $_.Matches } 

        #Find any end command characters that are not within a PL/SQL block
        $NonBlockCommands = $EndCommandCharacters | ForEach-Object {
            $BlockChecks = 0
            foreach($block in $BeginEndBlocks) {
                $StartBlockLocation = $block.Index
                $EndBlockLocation = ($block.Index + $Block.Length)

                #If the index of the end character is outside the block range, we count the block as checked
                if($_.Index -le $StartBlockLocation -or $_.Index -ge $EndBlockLocation) {
                    $BlockChecks++
                }
            }

            #Only if all blocks have been checked and do not contain the character do we include this as a non block character.
            if($BlockChecks -eq $BeginEndBlocks.Count) {
                $_
            }
        }

        #Remove the PL/SQL Blocks from the query (so we can find non block commands)
        $RemovedEndBlocks = $NoCommentQuery
        foreach($block in $BeginEndBlocks) {
            $startIndex = $block.Index
            $Length = $block.Length
            $RemovedEndBlocks = $RemovedEndBlocks.Remove($startIndex, $Length)
        }


        #Split the remaining, non block queries
        $NonBlockQueries = $RemovedEndBlocks -Split ";+(?=(?:[^\']*\'[^\']*\')*[^\']*$)"
        $NonBlockQueries = $NonBlockQueries.Split([string[]]"`r`n/", [StringSplitOptions]::None)

        #Cleanup and remove any empty elements
        $NonBlockQueries = ($NonBlockQueries.Trim()) | Where-Object {$_}

        #Now that we know the locations of any blocks and non-block commands, we can build the array of commands in order that they appear.
        $AllCommandLocations = @()
        $AllCommandLocations += $BeginEndBlocks 
        $AllCommandLocations += $NonBlockCommands

        [string[]]$OracleQueries = @()

        $CommandIndex = 0
        foreach($command in ($AllCommandLocations | Sort-Object Index)) {
            #Any object with a length of over 3 is a command block, as opposed to ';' or '\r\n/'
            
            if($command.Length -gt 3) {
                #remove trailing backslash from a PL/SQL block if it exists. The trailing backslash will cause an error.
                if($command.Value[-1] -eq '/') {
                    $sqlText = $command.Value -replace ".$"
                }
                else {
                    $sqlText = $command.Value
                }

                $OracleQueries += $sqlText
            }
            else {
                #We can only access using the array index if there is more than one member
                if($NonBlockQueries.Count -gt 1) {
                    $OracleQueries += $NonBlockQueries[$CommandIndex]
                }
                else {
                    $OracleQueries += $NonBlockQueries
                }
                $commandIndex++
            }
        }

        #This ensures that if the last query in the command is missing a ';' is also included. CommandLocations works off of ';' locations, so if it's missing one it means the query is at the end, but we still need to include it.
        if($NonBlockQueries.Count -eq ($AllCommandLocations.Count + 1)) {
            if($NonBlockQueries.Count -gt 1) {
                $OracleQueries += $NonBlockQueries[-1] 
            }
            else {
                $OracleQueries += $NonBlockQueries
            }
        }

    ##################################
    # End Query Prep
    ##################################

    $ModuleRoot = $MyInvocation.MyCommand.Module.ModuleBase
    
    # Load the Oracle.ManagedDataAccess.dll
    $LocalOracleFiles = "$ModuleRoot\bin"
    $DllFilePath = "$LocalOracleFiles\Oracle.ManagedDataAccess.dll"
    
    if(Test-Path $DllFilePath) {
        Write-Verbose "$DllFilePath found successfully. Attempting to run Unblock-File"
        Unblock-File $DllFilePath
    }
    else {
        Throw "Oracle.ManagedDataAccess.dll not found at $LocalOracleFiles"
    }

    try {
        Add-Type -Path $DllFilePath
    }
    catch {
        Write-Error $_.Exception.Message
        Throw "Issue adding Oracle.ManagedDataAccess.dll from the Oracle Home."           
    }

    # Update the TNS_ADMIN to use the local sqlnet.ora file (this sqlnet.ora SQLNET.AUTHENTICATION_SERVICES = (NTS))
    $CurrentTnsAdmin = $env:TNS_ADMIN 
    $env:TNS_ADMIN = $LocalOracleFiles

    Write-Verbose "Current TNS_ADMIN: $CurrentTnsAdmin"
    Write-Verbose "Setting TNS_ADMIN to $LocalOracleFiles temporarily."

    #Get IP address for a better Connection String. Using "localhost" in connect descriptor caused queries on some servers to fail, using the IP address is better.
    $TargetIpAddress = Test-Connection -ComputerName $HostName -Count 1  | Select-Object -ExpandProperty IPV4Address | Select-Object -ExpandProperty IPAddressToString

    #If a credential is passed, we use that instead of using windows auth
    if($DatabaseCredential) {
        $DbAccountUsername = $DatabaseCredential.UserName
        $DbAccountPassword = $DatabaseCredential.GetNetworkCredential().password

        if($AsSysdba) {
            $ConnectionString = "User Id=$DbAccountUsername;Password=$DbAccountPassword;DBA Privilege=SYSDBA;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$TargetIpAddress)(PORT=$PortNumber))(CONNECT_DATA=(SERVICE_NAME=$ServiceName)))"
        }
        else {
            $ConnectionString = "User Id=$DbAccountUsername;Password=$DbAccountPassword;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$TargetIpAddress)(PORT=$PortNumber))(CONNECT_DATA=(SERVICE_NAME=$ServiceName)))"
        }
        Write-Verbose "Connection String: $($ConnectionString -replace $DbAccountPassword, 'xxxx')"

    }

    else {
        $ConnectionString = "User Id=/;DBA Privilege=SYSDBA;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$TargetIpAddress)(PORT=$PortNumber))(CONNECT_DATA=(SERVICE_NAME=$ServiceName)))"
        Write-Verbose "Connection String: $ConnectionString"
    }

    :Queries foreach($Query in $OracleQueries) {
        #Clear any empty lines and uneeded whitespace.
        $Query = $Query.Trim()

        $QueryResult = @()
        try { 
            $Connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($ConnectionString)
            $cmd=$Connection.CreateCommand()

            $cmd.CommandText= $Query

            $da=New-Object Oracle.ManagedDataAccess.Client.OracleDataAdapter($cmd);
            $QueryResult=New-Object System.Data.DataTable
            [void]$da.fill($QueryResult)
            #This expands the results set to human readable form - without this it just shows "System.Data.DataRow" for each resultset.
            $QueryResult = $QueryResult | Select-Object -Property * -ExcludeProperty RowError,RowState,table,ItemArray,HasErrors
        } 

        catch {
            $ErrorMessage = $_.Exception.Message.ToString()

            $EncounteredError = $true

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
            if($QueryVerb -eq 'SELECT') {
                $QueryResult = "no rows selected"
            }
            elseif($QueryVerb -like 'BEGIN*' -or $QueryVerb -like 'DECLARE*') {
                $QueryResult = "PL/SQL procedure successfully completed."                    
            }
            
            #if it wasn't a select, and there was nothing returned, it means it was successful. So we create a result message.
            else {
                if($QueryVerb -eq 'DROP'){
                    $QueryVerbPastTense = "dropped"
                }
                elseif($QueryVerb -eq 'CREATE'){
                    $QueryVerbPastTense = "created"
                }
                elseif($QueryVerb -eq 'TRUNCATE'){
                    $QueryVerbPastTense = "truncated"
                }
                elseif($QueryVerb -eq 'DELETE'){
                    $QueryVerbPastTense = "deleted"
                }
                elseif($QueryVerb -eq 'ALTER'){
                    $QueryVerbPastTense = "altered"
                }
                elseif($QueryVerb -eq 'EXECUTE'){
                    $QueryVerbPastTense = "executed"
                }
                elseif($QueryVerb -eq 'GRANT'){
                    $QueryVerbPastTense = "granted"
                }
                #If it's not one of the above verbs, we default to succeeded.
                else{
                    $QueryVerbPastTense = "succeeded"
                }

                $QueryResult = "$QuerySubjectType $QueryVerbPastTense."
            }
        }   


        if(!$ResultSetOnly) {
            $QueryObject = New-Object PSObject  
            Add-Member -memberType NoteProperty -InputObject $QueryObject -Name Query -Value $Query  
            Add-Member -memberType NoteProperty -InputObject $QueryObject -Name ResultSet -Value $QueryResult     
        }
        else {
            $QueryObject = @() 
            $QueryObject += , $QueryResult  

        }

        $QueryObject 


        #If there was a connection error, we do not continue with the remaining queries.
        if($ConnectionError){
            Write-Warning "Issue connecting to the database. Not continuing with remaining queries to prevent lockouts."
            break Queries
        }
        elseif($EncounteredError -and $ExitOnError) {
            Write-Warning "Error encountered and EncounteredError switch is set to true. Stopping script."
            break Queries
        }

    }#End foreach Query loop
    
    #Reset TNS_ADMIN to default value in the session.
    Write-Verbose "Resetting TNS_ADMIN to $CurrentTnsAdmin."
    $env:TNS_ADMIN = $CurrentTnsAdmin
}
