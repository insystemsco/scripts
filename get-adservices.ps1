function Get-ADServices {
          
    [CmdletBinding()]
    [OutputType([Array])] 
    param
    (
        [Parameter(Position=0, Mandatory = $true, HelpMessage="Provide server names", ValueFromPipeline = $true)]
        $Computername
    )
 
    $ServiceNames = "HealthService","NTDS","NetLogon","DFSR"
    $ErrorActionPreference = "SilentlyContinue"
    $report = @()
 
        $Services = Get-Service -ComputerName $Computername -Name  $ServiceNames
 
        If(!$Services)
        {
            Write-Warning "Something went wrong"
        }
        Else
        {
            # Adding properties to object
            $Object = New-Object PSCustomObject
            $Object | Add-Member -Type NoteProperty -Name "ServerName" -Value $Computername
 
            foreach($item in $Services)
            {
                $Name = $item.Name
                $Object | Add-Member -Type NoteProperty -Name "$Name" -Value $item.Status 
            }
             
            $report += $object
        }
     
    $report
}