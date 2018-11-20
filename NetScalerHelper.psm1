# Module for Backup-NetScaler Project

function Split-Action
{
    param (
        [String]$ActionCommand
    )

    if ($ActionCommand -match "\.exe.+")
    {
        # There are probably parameters!
        $ActionExecutable = $ActionCommand -replace "\.exe.*$",""
        $ActionArgument = $ActionCommand -replace "^.*\.exe ",""
        $actionObj = New-Object -TypeName PSCustomObject
        Add-Member -InputObject $actionObj -MemberType NoteProperty -Name "Executable" -Value ($ActionExecutable + ".exe")
        Add-Member -InputObject $actionObj -MemberType NoteProperty -Name "Arguments" -Value $ActionArgument
        return $actionObj
    }
    else
    {
        $actionObj = New-Object -TypeName PSCustomObject
        Add-Member -InputObject $actionObj -MemberType NoteProperty -Name "Executable" -Value $ActionCommand
        Add-Member -InputObject $actionObj -MemberType NoteProperty -Name "Arguments" -Value $null
        return $actionObj
    }
}

function New-ScheduledJob
{
    param (
        [Parameter(Mandatory=$true)]
        [System.DayOfWeek]$DayOfWeek,

        [Parameter(Mandatory=$true)]
        [Int]$HourOfDay,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Daily","Weekly")]
        [String]$RepetitionInterval,

        [Parameter(Mandatory=$true)]
        [String]$TaskName,

        [Parameter(Mandatory=$true)]
        [String]$ActionCommand,

        [Parameter(Mandatory=$false)]
        [switch]$OverwriteExisting,

        [Parameter(Mandatory=$false)]
        [PSCredential]$UserAccount
    )

    $osVersion = Get-WmiObject -Class Win32_OperatingSystem
    switch -regex ($osVersion.Caption) {
        '^Microsoft Windows Server 2008.*' { 
            return $null
        }
        default {
            $clockVar = $HourOfDay.ToString() + ":00"
    
            $existingJob = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    
            #return $UserAccount.UserName
    
            if ($existingJob)
            {
                if ($OverwriteExisting)
                {
                    Unregister-ScheduledTask -InputObject $existingJob -Confirm:$false
                }
                else
                {
                    return $false
                }
            }
    
    
            Switch ($RepetitionInterval)
            {
                "Daily"
                    {
                        $trigger = New-ScheduledTaskTrigger -At $clockVar -Daily
                    }
                "Weekly"
                    {
                        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $clockVar
                    }
            }
    
            # Split Action Command
    
            $actionObj = Split-Action -ActionCommand $ActionCommand
            if ($actionObj.Arguments)
            {
                $action = New-ScheduledTaskAction -Execute $actionObj.Executable -Argument $actionObj.Arguments
            }
            else
            {
                $action = New-ScheduledTaskAction -Execute $actionObj.Executable
            }
    
            if ($UserAccount)
            {
                # $cimSession = New-CimSession -Credential $UserAccount
                # $cred = New-ScheduledTaskPrincipal -UserId $UserAccount.UserName
                # $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -CimSession $cimSession -Principal $cred
                $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -User $UserAccount.UserName -Password $UserAccount.GetNetworkCredential().Password
            }
            else
            {
                $userId = $env:USERDOMAIN + "\" + $env:USERNAME
                # $cred = New-ScheduledTaskPrincipal -LogonType Password -UserId $userId -
                $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger # -Principal $cred
            }
    
            return $scheduledTask
        }       
    }
}

function Backup-Netscaler
{
    [CmdletBinding(DefaultParameterSetName="Default")]
    param (
        [Parameter(Mandatory)]
        [ValidateScript({Test-Connection $_ -Quiet -Count 3})]
        [String]
        $NetScalerIp,

        [Parameter(Mandatory)]
        [ValidateSet("HTTP","HTTPS")]
        [String]
        $NetScalerProtocol,

        [Parameter(Mandatory)]
        [String]
        $NetScalerUser,

        [Parameter(Mandatory)]
        [String]
        $NetScalerPassword,

        [Parameter(Mandatory)]
        [String]
        $BackupFileNamePrefix,

        [Parameter(Mandatory=$false)]
        [ValidateSet("full","basic")]
        [String]
        $BackupLevel = "full",

        [Parameter(Mandatory)]
        [ValidateScript({Get-Command $_})]
        [String]
        $PathToScp,

        [Parameter(Mandatory)]
        [String]
        $BackupLocation,

        [Parameter(Mandatory,ParameterSetName="ByMail")]
        [String]
        $EmailSmtp,

        [Parameter(Mandatory,ParameterSetName="ByMail")]
        [String]
        $EmailFrom,

        [Parameter(Mandatory,ParameterSetName="ByMail")]
        [String]
        $EmailTo
    )

    # Setup Persistence Flag
    $persistenceFlag = "persistent"

    # Sanitize Variables
    $NetScalerProtocol = $NetScalerProtocol -replace "\W.*$",""
    $BackupFileName = "$BackupFileNamePrefix-$(get-date -uformat '%d-%m-%Y-%H-%M')"

    if (-not (Test-Path $BackupLocation)) {
        try {
            New-Item $BackupLocation -ItemType Directory -ErrorAction Stop
        }
        catch [System.Management.Automation.ActionPreferenceStopException] {
            Write-Error -Message "Could not create backup directory" -ErrorAction Stop
        }
    }

    try {
        $nsCred = New-Object System.Management.Automation.PSCredential "$NetScalerUser",(ConvertTo-SecureString -AsPlainText $NetScalerPassword -Force)
        switch ($NetScalerProtocol)
        {
            "HTTP" {
                $nssession = Connect-NetScaler -IPAddress $NetScalerIp -Credential $nsCred -Timeout 60 -ErrorVariable connectionError -PassThru
            }
            "HTTPS" {
                $nssession = Connect-NetScaler -IPAddress $NetScalerIp -Credential $nsCred -Timeout 60 -Https -ErrorVariable connectionError -PassThru
            }
            default {}
        }
    }
    catch
    {
        Write-Error -Message $PSItem.exception.Message -ErrorAction Stop
    }

    try {
        # save config and create system backup
        Save-NSConfig -Session $nssession -ErrorVariable nitroError
        New-NSBackup -Session $nssession -Name $BackupFileName -Level $BackupLevel -ErrorVariable nitroError
    }
    catch
    {
        Write-Error -Message $PSItem.Exception.Message -ErrorAction Stop
    }
    $NetScalerPassword = '"' + $NetScalerPassword + '"'
    $Arguments = '/c echo y  | ' + " $PathToScp -pw $NetScalerPassword $NetScalerUser@`"$NetScalerIp`":/var/ns_sys_backup/$BackupFileName.tgz $BackupLocation"
    Start-Process cmd.exe -ArgumentList $Arguments -Wait | Out-Null

    # Clear oldest Backup File
    $backupArr = (Get-NSBackup -Session $nssession) | Where-Object {$_.filename -ne "$BackupFileName.tgz" -and $_.filename -notmatch "$persistenceFlag"}

    if ($backupArr)
    {
        ($backupArr | Sort-Object creationtime)[0] | Remove-NSBackup -Confirm:$false -Session $nssession
    }

    $backupFile = Join-Path -Path $BackupLocation -ChildPath "$backupFileName.tgz"
    if (Test-Path $backupFile -ErrorAction SilentlyContinue)
    {
        return (get-item $backupFile)
    }
    else
    {
        Write-Error -Message "Could not save backup file at $backupFile" -ErrorAction Stop
    }
}

function Reset-FormObject
{
    param (
        $FormObject
    )

    $childArr = $FormObject.Content.Children

    if ($childArr)
    {
        $childArr | ForEach-Object { Reset-FormObject -FormObject $_ }
    }
    else
    {
        switch ( $FormObject.GetType() )
        {
            "System.Windows.Controls.Combobox" {
                $FormObject.SelectedIndex = -1
            }
            "System.Windows.Controls.StackPanel" {
                $FormObject.Children | ForEach-Object {Reset-FormObject -FormObject $_}
            }
            "System.Windows.Controls.TextBox" {
                $FormObject.Text = $null
            }
            "System.Windows.Controls.RadioButton" { $FormObject.IsChecked = $false }
            default {return $true}
        }
    }
}
<#
function Create-ScheduledJob
{
    param (
        [Parameter(Mandatory=$true)]
        [System.DayOfWeek]$DayOfWeek,

        [Parameter(Mandatory=$true)]
        [Int]$HourOfDay,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Daily","Weekly")]
        [String]$RepetitionInterval,

        [Parameter(Mandatory=$true)]
        [String]$TaskName,

        [Parameter(Mandatory=$true)]
        [String]$ActionCommand,

        [Parameter(Mandatory=$false)]
        [switch]$OverwriteExisting,

        [Parameter(Mandatory=$false)]
        [PSCredential]$UserAccount
    )

    $clockVar = $HourOfDay.ToString() + ":00"

    $existingJob = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    #return $UserAccount.UserName

    if ($existingJob)
    {
        if ($OverwriteExisting)
        {
            Unregister-ScheduledTask -InputObject $existingJob -Confirm:$false
        }
        else
        {
            return $false
        }
    }


    Switch ($RepetitionInterval)
    {
        "Daily"
            {
                $trigger = New-ScheduledTaskTrigger -At $clockVar -Daily
            }
        "Weekly"
            {
                $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $clockVar
            }
    }

    # Split Action Command

    $actionObj = Split-Action -ActionCommand $ActionCommand
    if ($actionObj.Arguments)
    {
        $action = New-ScheduledTaskAction -Execute $actionObj.Executable -Argument $actionObj.Arguments
    }
    else
    {
        $action = New-ScheduledTaskAction -Execute $actionObj.Executable
    }

    if ($UserAccount)
    {
        # $cimSession = New-CimSession -Credential $UserAccount
        # $cred = New-ScheduledTaskPrincipal -UserId $UserAccount.UserName
        # $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -CimSession $cimSession -Principal $cred
        $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -User $UserAccount.UserName -Password $UserAccount.GetNetworkCredential().Password
    }
    else
    {
        $userId = $env:USERDOMAIN + "\" + $env:USERNAME
        # $cred = New-ScheduledTaskPrincipal -LogonType Password -UserId $userId -
        $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger # -Principal $cred
    }

    return $scheduledTask
}
#>

# external function declarations
. .\function-Test-NetScalerConnection.ps1

Export-ModuleMember -Function Split-Action,New-ScheduledJob,Backup-NetScaler,Reset-FormObject,Test-NetScalerConnection