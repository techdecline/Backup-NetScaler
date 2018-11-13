<#
.Synopsis
   Test Connection for a resource on HTTP/S and SSH
.DESCRIPTION
   Test Connection for a resource on HTTP/S and SSH
.EXAMPLE
   PS> Test-NetScalerConnection 172.16.0.1
   Tests "172.16.0.1" on HTTP (TCP:80) and SSH (TCP:22)
.EXAMPLE
   PS> Test-NetScalerConnection -ComputerName 172.16.0.1 -Secure
   Tests "172.16.0.1" on HTTPS (TCP:443) and SSH (TCP:22)
#>
function Test-NetScalerConnection
{
    [CmdletBinding()]
    [OutputType([bool])]
    Param
    (
        # NetScaler Host Name or IP
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [String]$ComputerName,

        # Switch to toggle HTTPS
        [switch]
        $Secure
    )

    Process {
        Write-Verbose "Checking for NetScaler Connectivity on Host: $ComputerName"
        Write-Verbose "HTTPS selection: $Secure"

        Write-Verbose "Testing Web Connectivity"
        switch ($Secure)
        {
            $true {
                $webTest = Test-NetConnection -ComputerName $ComputerName -Port 443 -InformationLevel Quiet    
            }
            $false {
                $webTest = Test-NetConnection -ComputerName $ComputerName -Port 80 -InformationLevel Quiet                    
            }
        }
        if ($webTest) {
            Write-Verbose "Web Connectivity verified successfully"
            $scpTest = Test-NetConnection -ComputerName $ComputerName -Port 22 -InformationLevel Quiet
            if ($scpTest) {
                Write-Verbose "SCP Connectivity verified successfully"
                return $true
            }
            else {
                Write-Verbose "SCP Connectivity verification failed"            
                return $false
            }
        }
        else {
            Write-Verbose "Web Connectivity verification failed"
            return $false
        }
    }
}