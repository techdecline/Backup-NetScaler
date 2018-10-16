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
        $cimSession = New-CimSession -Credential $UserAccount
        $cred = New-ScheduledTaskPrincipal -LogonType s4U -UserId $UserAccount.UserName
        $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -CimSession $cimSession #-Principal $cred
    }
    else
    {
        $userId = $env:USERDOMAIN + "\" + $env:USERNAME
        # $cred = New-ScheduledTaskPrincipal -LogonType Password -UserId $userId -
        $scheduledTask = Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger # -Principal $cred
    }

    return $scheduledTask
}

#######################################################################################################
# $nsip = NetScaler NS IP Address
# $NetScalerProtocol = HTTP oder HTTPS
# $NetScalerUser = NetScaler-Login Nutzername
# $NetScalerPassword = NetScaler-Login Passwort
# $BackupFileName = Benennung Backup-File
# $BackupLevel = full oder basic | http://support.citrix.com/proddocs/topic/ns-system-10-5-map/ns-system-backup1-tsk.html
# $PathToScp = Speicherort der pscp.exe
# $BackupLocation = Speicherort des Backup-File (Ordner wird im Skript, wenn nicht vorhanden, angelegt)
# $SmtpServer = SMTP Address des Mailservers (leere Zeichenkette: kein Mailversand)
# $MailTo = Mail-Empfänger-Adresse / Verteiler
#######################################################################################################

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
        $BackupFileName,

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
    $BackupFileName = "$BackupFileName-$(get-date -uformat '%d-%m-%Y-%H-%M')"


    New-Item -Path $BackupLocation -ItemType directory -ErrorAction SilentlyContinue

    if ( Get-Module NetScaler)
    {
        try {
            #Set-NSMgmtProtocol $NetScalerProtocol
            #$nssession = Connect-NSAppliance -NSAddress $NetscalerIp -NSUserName $NetScalerUser -NSPassword $NetScalerPassword -Timeout 60
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
            if ($EmailSmtp) {
                Send-MailMessage -From $EmailFrom -SmtpServer $EmailSmtp -To $EmailTo -Body "$connectionError" -Subject "Netscaler Backup failed"
            }
            return $connectionError
        }
        #return $nssession
        #$payload = @{"level"="full";"filename"="$BackupFileName"}
        $payload = @{level="full"}

        try {
            # save config and create system backup
            Save-NSConfig -Session $nssession -ErrorVariable nitroError
            New-NSBackup -Session $nssession -Name $BackupFileName -Level $BackupLevel -ErrorVariable nitroError
        }
        catch
        {
            if ($EmailSmtp) {
                Send-MailMessage -From $EmailFrom -SmtpServer $EmailSmtp -To $EmailTo -Body "$nitroError" -Subject "Netscaler Backup failed"
            }
            return $nitroError
        }
        # return "Backup Created."
        $NetScalerPassword = '"' + $NetScalerPassword + '"'
        $Arguments = '/c echo y  | ' + " $PathToScp -pw $NetScalerPassword $NetScalerUser@`"$NetScalerIp`":/var/ns_sys_backup/$BackupFileName.tgz $BackupLocation"
        Start-Process cmd.exe -ArgumentList $Arguments -Wait | Out-Null

        #[System.Windows.Forms.MessageBox]::Show("")
        <# Old backup File handling
        $backupFile = Join-Path $BackupLocation -ChildPath "$BackupFileName.tgz"
        #[System.Windows.Forms.MessageBox]::Show($backupFile)

        Remove-NSBackup -Session $nssession -Name "$BackupFileName.tgz" -Confirm:$false
        #>

        # Clear oldest Backup File
        $backupArr = (Get-NSBackup -Session $nssession) | Where-Object {$_.filename -ne "$BackupFileName.tgz" -and $_.filename -notmatch "$persistenceFlag"}

        if ($backupArr)
        {
            ($backupArr | Sort-Object creationtime)[0] | Remove-NSBackup -Confirm:$false -Session $nssession
        }

        if (Test-Path $backupFile -ErrorAction SilentlyContinue)
        {
            if ($EmailSmtp) {
                Send-MailMessage -From $EmailFrom -SmtpServer $EmailSmtp -To $EmailTo -Body "NetScaler has been backed up succesfully and stored at $backupFile" -Subject "Netscaler Backup succesfully created"
            }
            return $backupFile
        }
        else
        {
            if ($EmailSmtp) {
                Send-MailMessage -From $EmailFrom -SmtpServer $EmailSmtp -To $EmailTo -Body "No error was specified." -Subject "Netscaler Backup failed"
            }
            return $false
        }
    }
    else
    {
        if ($EmailSmtp) {
            Send-MailMessage -From $EmailFrom -SmtpServer $EmailSmtp -To $EmailTo -Body "No error was specified." -Subject "Netscaler Backup failed"
        }
        $false
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
        #[System.Windows.Forms.MessageBox]::Show($FormObject.Name)
        #[System.Windows.Forms.MessageBox]::Show($FormObject.GetType())

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

Export-ModuleMember -Function Split-Action,Create-ScheduledJob,Backup-NetScaler,Reset-FormObject