#Get public and private function definition files.

$Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )

$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

foreach($import in @($Public + $Private)) {
    try {
        . $import.fullname
    }
    catch{
        Write-Error -Message "Failed to import function $($import.fullname): $_"
    }
}

Export-ModuleMember -Function $Public.Basename

#region Initialization
Get-ChildItem -Path "$PSScriptRoot/Init" | ForEach-Object {
    Write-Verbose "Initializing module: [$($_.Name)]"
    . $_.FullName
}
