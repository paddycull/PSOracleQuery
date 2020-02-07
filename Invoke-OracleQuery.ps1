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
Invoke-OracleQuery -TargetComputer VIRTORA195s -TargetDatabase PATPDB1 -Query "select username, account_status from dba_users;"
    This example runs the given query against the PATPDB1 database on VIRTORA195s

.EXAMPLE
Invoke-OracleQuery -TargetComputer VIRTORA195s -TargetDatabase PATPDB1 -SqlFile C:\example\oracle.sql
    This runs the oracle.sql file against the database.

.NOTES
Author: Patrick Cull
Date: 03/02/2020
#>
function Invoke-OracleQuery {
    param(   
        #Computer the database is on 
        [Parameter(Mandatory)]
        [string] $TargetComputer,

        #Databsase to query
        [Parameter(Mandatory)]
        [string] $TargetDatabase,
        
        #Credential to connect to the target computer. Defaults to current credential if not passed.
        [PSCredential] $TargetCredential=[System.Management.Automation.PSCredential]::Empty, 

        #User credential to connect to the database with. Defaults to current windows user as sysdba if not passed.
        [PSCredential] $DatabaseCredential,

        #Query to run
        [string] $Query,
        #Sql file to run against the database.
        [string] $SqlFile
    )

    if(!$Query -and !$SqlFile){
        Throw "Please pass either a Query or a SqlFile to run against the target database."
    }

    if($SqlFile) {
        $Query = (Get-Content $SqlFile | Out-String).Trim()
    }

    #If the final character in the query is a slash or semicolon, we remove it, as the oracleManagedDataAccess module does not allow thse characters at the end of a query
    if($Query[-1] -eq "/" -or $Query[-1] -eq ';') {
        $Query = $Query -replace ".$"
    }

    #Split the queries into individual queries, either on semicolon outside single quotes, or a / on it's own line - i.e. the query terminators in Oracle.
    $OracleQueries = $Query -Split ";+(?=(?:[^\']*\'[^\']*\')*[^\']*$)"
    $OracleQueries = $OracleQueries.Split([string[]]"`r`n/", [StringSplitOptions]::None)

    $Output = Invoke-Command -ComputerName $TargetComputer -Credential $TargetCredential -HideComputerName -ArgumentList $TargetDatabase, $OracleQueries, $DatabaseCredential {
        param($TargetDatabase, $OracleQueries, $DatabaseCredential)

        #Get the oracle home on the server so we can get then import the Oracle dll on the target 
        $OracleServices = Get-WmiObject win32_service | Where-Object {$_.Name -like 'OracleService*' -and $_.State -eq 'Running'} | Select-Object Name, PathName
        #Only need one home, just use the first one returned.
        $FirstOracleService = $OracleServices[0].PathName
        #Split service string to extract home version and location
        $OracleHome = $FirstOracleService.Substring(0, $FirstOracleService.lastIndexOf('\bin\ORACLE.EXE'))

        Add-Type -Path "$OracleHome\ODP.NET\managed\common\Oracle.ManagedDataAccess.dll"

        
        $AllResults = @()
        $QueryCount = $OracleQueries.Count


        #If a credential is passed, we use that instead of using windows auth
        if($DatabaseCredential) {
            $DbAccountUsername = $DatabaseCredential.UserName
            $DbAccountPassword = $DatabaseCredential.GetNetworkCredential().password
            $ConnectionString = "User Id=$DbAccountUsername;Password=$DbAccountPassword;Data Source=$TargetDatabase"
        }

        else {
            $ConnectionString = "User Id=/;DBA Privilege=SYSDBA;Data Source=$TargetDatabase"
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
                $StringQuery = $Query | Out-String     
                Add-Member -memberType NoteProperty -InputObject $QueryObject -Name Query -Value $StringQuery  
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
    }

    #If the Output type is an object, it means a proper query resultset has been returned, so we remove the PSComputerName, RunspaceId and PSShowComputerName that gets added to objects by PowerShell after the Invoke-Command
    if($Output.GetType().Name -like '*Object*') {
        $Output | Select-Object -Property * -ExcludeProperty PSComputerName, RunspaceId, PSShowComputerName
    }
    #Otherwise we leave the output unchanged.
    else {
        $Output
    }
}

#Register Argument that runs 'lsnrctl status' on the target and extracts the database services for the TargetDatabase parameter.
Register-ArgumentCompleter -CommandName Invoke-OracleQuery -ParameterName TargetDatabase -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
    $TargetComputer = $fakeBoundParameter.TargetComputer

    $ServiceNames = Invoke-Command -ComputerName $TargetComputer -ScriptBlock {
        $ListenerOutput = lsnrctl status
        
        #Extract any string that starts with 'Service' up to the end '"'
        $ServiceNames = ([regex]::Matches($ListenerOutput, 'Service([^/)]+")') |ForEach-Object { $_.Groups[1].Value }).Trim()

        #remove extra uneeded strings
        $ServiceNames = ($ServiceNames -replace "s Summary... Service ") -replace '"'
          
        return $ServiceNames
    }
    
    #Remove XDB service, CLRextProc and any PDB GUID services (which 32 chars long)
    $DatabaseServices = $ServiceNames.ToUpper() | Where-Object{$_ -NotLike '*XDB' -and $_ -ne 'CLRExtProc' -and $_.Length -ne 32}

    return $DatabaseServices
}