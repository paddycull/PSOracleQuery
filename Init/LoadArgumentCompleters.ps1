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
