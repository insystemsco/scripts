If($credentials.GetNetworkCredential().password -eq $null )
{
    Write-Warning "Credential validation failed"
    pause
    Break
}
Else
{
    $CredCheck = $Credentials  | Test-Cred
    If($CredCheck -ne "Authenticated")
    {
        Write-Warning "Credential validation failed"
        pause
        Break
    }     
}