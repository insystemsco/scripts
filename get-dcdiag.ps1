 function Get-DCDiag {
          
    [CmdletBinding()]
    [OutputType([Array])] 
    param
    (
        [Parameter(Position=0, Mandatory = $true, HelpMessage="Provide server names", ValueFromPipeline = $true)]
        $Computername
    )
    $DCDiagArray = @()
 
            # DCDIAG ===========================================================================================
            $Dcdiag = (Dcdiag.exe /s:$Computername) -split ('[\r\n]')
            $Results = New-Object Object
            $Results | Add-Member -Type NoteProperty -Name "ServerName" -Value $Computername
            $Dcdiag | %{ 
            Switch -RegEx ($_) 
            { 
                "Starting test"      { $TestName   = ($_ -Replace ".*Starting test: ").Trim() } 
                "passed test|failed test" { If ($_ -Match "passed test") {  
                $TestStatus = "Passed" 
                # $TestName 
                # $_ 
                }  
                Else 
                {  
                $TestStatus = "Failed" 
                # $TestName 
                # $_ 
                } 
                } 
            } 
            If ($TestName -ne $Null -And $TestStatus -ne $Null) 
            { 
                $Results | Add-Member -Name $("$TestName".Trim()) -Value $TestStatus -Type NoteProperty -force
                $TestName = $Null; $TestStatus = $Null
            } 
            } 
            $DCDiagArray += $Results
 
    $DCDiagArray
             
}