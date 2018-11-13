#Requires -Version 4.0 -RunAsAdministrator

param (
    [switch]$DebugMode
)

function Get-ScriptDirectory {
    $Invocation = (Get-Variable MyInvocation -Scope 1).Value
    Split-Path $Invocation.MyCommand.Path
}

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Class="Netscaler_Backup.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Netscaler_Backup"
        mc:Ignorable="d"
        Title="Netscaler Backup Automator" Height="675" Width="525">
    <Grid>
        <Image x:Name="imageLogo" HorizontalAlignment="Left" Height="100" Margin="10,15,0,0" VerticalAlignment="Top" Width="100" Source="C:\Code\Backup-NetScaler\img\softed.jpg" Visibility="Hidden"/>
        <TextBlock x:Name="textBlock" HorizontalAlignment="Left" Margin="115,43,0,0" TextWrapping="Wrap" Text="Netscaler Backup Automator" VerticalAlignment="Top" Height="30" Width="381" FontFamily="Calibri" FontSize="24" FontWeight="Bold" TextAlignment="Center"/>
        <GroupBox x:Name="groupBoxScheduler" Header="Scheduler" HorizontalAlignment="Left" Margin="10,115,0,0" VerticalAlignment="Top" Height="Auto" Width="490">
            <Grid x:Name="gridScheduler">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="75" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="60" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Grid.Row="0" x:Name="textBlock1" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Day of week:" VerticalAlignment="Center" />
                <ComboBox Grid.Column="1" Grid.Row="0" x:Name="comboBoxDayOfWeek" Width="Auto" />
                <TextBlock Grid.Column="2" Grid.Row="0" x:Name="textBlock2" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Hour:" VerticalAlignment="Center"/>
                <ComboBox Grid.Column="3" Grid.Row="0" x:Name="comboBoxHour" Width="Auto" />
                <TextBlock Grid.Column="0" Grid.Row="1" x:Name="textBlock3" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Repetition:" VerticalAlignment="Center" Margin="0,5"/>
                <!-- <ComboBox Grid.Column="1" Grid.Row="1" x:Name="comboBox2" Width="Auto" Margin="0,5"/> -->
                <StackPanel Grid.Column="1" Grid.Row="1" Margin="0,5" >
                    <RadioButton GroupName="Repetition" Margin="5" x:Name="rbuttonDaily">Daily</RadioButton>
                    <RadioButton GroupName="Repetition" Margin="5" x:Name="rbuttonWeekly">Weekly</RadioButton>
                    <!--    <RadioButton GroupName="Repetition" Margin="5">Monthly</RadioButton> -->
                </StackPanel>
                <TextBlock Grid.Column="2" Grid.Row="1" x:Name="textBlock4" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Task Name:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="3" Grid.Row="1" x:Name="textBoxTaskName" Width="Auto" Margin="5" Height="20"/>
                <TextBlock Grid.Column="0" Grid.Row="2" x:Name="textBlockUser" HorizontalAlignment="Center" TextWrapping="Wrap" Text="User:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="1" Grid.ColumnSpan="2" Grid.Row="2" x:Name="textBoxTaskUser" IsReadOnly="True" Width="Auto" Margin="5" Height="20"/>
                <Button Grid.Row="2" Grid.Column="3" Margin="0,5" x:Name="buttonSelectUser">Select User...</Button>
                <Button Grid.Row="3" Grid.ColumnSpan="4" Margin="0,5" x:Name="buttonSchedule" FontWeight="Bold">Schedule Job</Button>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="groupBox1" Header="Netscaler Backup Settings" HorizontalAlignment="Left" Margin="10,280,0,0" VerticalAlignment="Top" Height="Auto" Width="490">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="75" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="60" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Grid.Row="0" x:Name="textBlock5" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Netscaler IP:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="1" Grid.Row="0" Margin="0,5" Height="20" x:Name="textboxNSIP"></TextBox>
                <TextBlock Grid.Column="2" Grid.Row="0" x:Name="textBlock6" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Protocol:" VerticalAlignment="Center" Margin="0,5"/>
                <StackPanel Grid.Column="3" Grid.Row="0" Margin="0,5" >
                    <RadioButton GroupName="Protocol" Margin="5" x:Name="rButtonHttp">HTTP</RadioButton>
                    <RadioButton GroupName="Protocol" Margin="5" x:Name="rButtonHttps">HTTPS</RadioButton>
                    <!--    <RadioButton GroupName="Repetition" Margin="5">Monthly</RadioButton> -->
                </StackPanel>
                <TextBlock Grid.Column="0" Grid.Row="1" x:Name="textBlock7" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Username:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="1" Grid.Row="1" Margin="0,5" Height="20" x:Name="textboxUserName" ToolTip="Enter the user name for NetScaler Web Interface"></TextBox>
                <TextBlock Grid.Column="2" Grid.Row="1" x:Name="textBlock8" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Password:" VerticalAlignment="Center" Margin="0,5"/>
                <!--<TextBox Grid.Column="3" Grid.Row="1" Margin="0,5" Height="20"></TextBox>-->
                <PasswordBox  x:Name="passwordBox" Grid.Column="3" Grid.Row="1" Margin="0,5" Height="20"/>
                <TextBlock Grid.Column="0" Grid.Row="2" x:Name="textBlock9" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Path to SCP:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="1" Grid.Row="2" Margin="0,5" Height="20" x:Name="textboxScpPath" ToolTip="Please enter absolute Path to PSCP.exe"></TextBox>
                <TextBlock Grid.Column="2" Grid.Row="2" x:Name="textBlock10" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Backup Level:" VerticalAlignment="Center" Margin="0,5"/>
                <ComboBox Grid.Column="3" Grid.Row="2" x:Name="comboBoxBackupLevel" Width="Auto" Height="20"/>
                <TextBlock Grid.Column="0" Grid.Row="3" x:Name="textBlock11" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Location:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="1" Grid.Row="3" Margin="0,5" Height="20" x:Name="textboxBackupLocation" ToolTip="Select a folder where to store the backup files"></TextBox>
                <TextBlock Grid.Column="2" Grid.Row="3" x:Name="textBlock12" HorizontalAlignment="Center" TextWrapping="Wrap" Text="File Prefix:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="3" Grid.Row="3" Margin="0,5" Height="20" x:Name="textboxFilePrefix" ToolTip="Enter a name prefix for the backup files (date will be automatically added)"></TextBox>
                <Button Grid.Column="0" Grid.Row="4" Grid.ColumnSpan="2" Margin="5" x:Name="buttonImport">Import Configuration</Button>
                <Button Grid.Column="2" Grid.Row="4" Grid.ColumnSpan="2" Margin="5" x:Name="buttonExport">Export Configuration</Button>
                <Button Grid.Column="0" Grid.Row="5" Grid.ColumnSpan="4" Margin="5" x:Name="buttonVerify">Verify Connection</Button>
            </Grid>
        </GroupBox>
        <GroupBox x:Name="groupboxEmail" Header="Email Notification" Margin="10,525,5,5" HorizontalAlignment="Left" VerticalAlignment="Top" Height="auto" Width="490">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="80" />
                    <ColumnDefinition Width="80" />
                    <ColumnDefinition Width="*" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Grid.Row="0" x:Name="textblockEmailEnabled" VerticalAlignment="Center">Email Enabled:</TextBlock>
                <StackPanel Grid.Column="1" Grid.Row="0" Margin="0,5" >
                    <RadioButton GroupName="MailEnabled" Margin="5" x:Name="rButtonEmailYes">Enabled</RadioButton>
                    <RadioButton GroupName="MailEnabled" Margin="5" x:Name="rButtonEmailNo" IsChecked="True">Disabled</RadioButton>
                    <!--    <RadioButton GroupName="Repetition" Margin="5">Monthly</RadioButton> -->
                </StackPanel>
                <TextBlock Grid.Column="2" Grid.Row="0" x:Name="textBlockSmtp" VerticalAlignment="Center" HorizontalAlignment="Center">SMTP Server:</TextBlock>
                <TextBox Grid.Column="3" Grid.Row="0" x:Name="textboxSmtp" IsEnabled="False"></TextBox>
                <TextBlock Grid.Column="0" Grid.Row="1" x:Name="textblockFrom" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Email From:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="1" Grid.Row="1" Margin="0,5" Height="20" x:Name="textboxEmailFrom" IsEnabled="False"></TextBox>
                <TextBlock Grid.Column="2" Grid.Row="1" x:Name="textblockTo" HorizontalAlignment="Center" TextWrapping="Wrap" Text="Email To:" VerticalAlignment="Center" Margin="0,5"/>
                <TextBox Grid.Column="3" Grid.Row="1" Margin="0,5" Height="20" x:Name="textboxEmailTo" IsEnabled="False"></TextBox>
            </Grid>
        </GroupBox>
    </Grid>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'


[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[xml]$XAML = $inputXML
#Read XAML

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}

Get-FormVariables

#===========================================================================
# Actually make the objects work
#===========================================================================

# Import Helper Functions
Import-Module "$PSScriptRoot\NetScalerHelper.psm1"

# Create global variable to store user object
$Global:CredUserObj = $null

# START Initialze Controls
[Enum]::GetNames( [System.DayOfWeek] ) | ForEach-Object {
    $WPFcomboBoxDayOfWeek.Items.Add($_)
}

$hourArr = 1..24
$hourArr | ForEach-Object {
    $WPFcomboBoxHour.Items.Add($_)
}
$WPFcomboBoxBackupLevel.Items.Add("Full")

#[System.Windows.Forms.MessageBox]::Show($env:USERDOMAIN + "\" + $env:USERNAME)
$WPFtextboxTaskUser.Text = $env:USERDOMAIN + "\" + $env:USERNAME

# Check for PSCP.exe in script directory
$pscpPath = Join-Path -Path (Get-ScriptDirectory) -ChildPath "ExterneRessourcen\PSCP.EXE"

if (Test-Path $pscpPath) {
    $WPFtextboxScpPath.Text = $pscpPath
}

# Check for Logo
$imagePath = Join-Path -Path (Get-ScriptDirectory) -ChildPath "img\softed.jpg"

if (Test-Path $imagePath) {
    $wpfimageLogo.Source = $imagePath
    $wpfimageLogo.Visibility = "Visible"
}

# Check for Debug Mode
if ($DebugMode) {
    $WPFbuttonSchedule.Content = "Generate Script"
}


# END Initialize Controls

# START Event Handler Definition

$WPFbuttonSelectUser.Add_Click({
    # DEBUG [System.Windows.Forms.MessageBox]::Show("button works")
    # $credObj = New-Object [PSCredential]

    try {
        $credObj = Get-Credential -Message "Please select a user with local administrative privileges to schedule your task:"
        if ($credObj)
        {
            $WPFtextboxTaskUser.Text = $credObj.UserName
            $Global:CredUserObj = $credObj
        }
        else
        {
            throw " Non-implemented Exception ;-)"
        }
    }
    catch
    {
        $Global:CredUserObj = $null
    }
})

$WPFbuttonSchedule.Add_Click({
    <# DEBUG
    if ($Global:CredUserObj)
    {
        [System.Windows.Forms.MessageBox]::Show("A different user has been selected")
    }
    else
    {
        [System.Windows.Forms.MessageBox]::Show("Task will be using current user")
    }

    return $true
    #>


    # Get selected Radio Button Schedule
    $rButtonArr = @($WPFrbuttonDaily,$WPFrbuttonWeekly)
    $rButtonArr | ForEach-Object {
        if ($_.isChecked)
        {
            $schedule = $_.Content
        }
    }

    $dayOfWeek = $WPFcomboBoxDayOfWeek.SelectedItem
    # [System.Windows.Forms.MessageBox]::Show($dayOfWeek)

    $hour = $WPFcomboBoxHour.SelectedValue
    # [System.Windows.Forms.MessageBox]::Show($hour)

    $taskName = $wpfTextboxTaskName.Text
    # [System.Windows.Forms.MessageBox]::Show($taskName)

    # Collect Input from lower controls

    $rButtonArr = @($WPFrButtonHttp,$WPFrButtonHttps)
    $rButtonArr | ForEach-Object {
        if ($_.isChecked)
        {
            $protocol = $_.Content
            #$protocol = $protocol + '://'
        }
    }


    $settingsObj = New-Object -TypeName PSCustomObject
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name NetscalerIP -Value $WPFtextboxNSIP.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name NetscalerProtocol -Value $protocol
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name UserName -Value $WPFtextboxUserName.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name Password -Value $WPFpasswordBox.Password.ToString()
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name BackupLevel -Value $WPFcomboBoxBackupLevel.SelectedValue
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name ScpPath -Value $WPFtextboxScpPath.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name BackupLocation -Value $WPFtextboxBackupLocation.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name FileName -Value $WPFtextboxFilePrefix.Text

    if ($wpfrButtonEmailYes.isChecked) {
        Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name EmailSmtp -Value $wpfTextboxSmtp.Text
        Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name EmailFrom -Value $wpftextboxEmailFrom.Text
        Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name EmailTo -Value $WPFtextboxEmailTo.Text
    }


    $netscalerArgumentsMissing = $settingsObj.NetScalerIp -eq $null -or $settingsObj.NetscalerProtocol -eq $null -or $settingsObj.UserName -eq $null `
        -or $settingsObj.Password -eq $null -or $settingsObj.BackupLevel -eq $null -or $settingsObj.scpPath -eq $null `
        -or $settingsObj.BackupLocation -eq $null -or $settingsObj.FileName -eq $null

    $schedulingArgumentsMissing = $dayOfWeek -eq $null -or $hour -eq $null -or $schedule -eq $null -or $taskName -eq $null

    $emailArgumentsMissing = $settingsObj.EmailSmtp -eq $null -or $settingsObj.EmailFrom -eq $null -or $settingsObj.EmailTo -eq $null -and $wpfrButtonEmailYes.IsChecked

    if ($schedulingArgumentsMissing -or $netscalerArgumentsMissing -or $emailArgumentsMissing)
    {
        [System.Windows.Forms.MessageBox]::Show("Required input missing. Please check parameter set!")
        return
    }

    # START Generate Backup Script

    $scriptArr = @()
    $modulePath = (Get-Module NetScalerHelper).Path
    $backupScriptFile = "$env:Temp\tmp.ps1"
    #$exeFilePath = "$PSScriptRoot\Backup.exe"
    $exeFilePath = Join-Path $PSScriptRoot -ChildPath ($taskName + ".exe")

    $moduleLoadCommand = 'Import-Module ' + $modulePath
    $backupCommand = 'Backup-NetScaler ' + "-NetScalerIp " + '"' + $($settingsObj.NetscalerIp) + '"' + " -NetscalerProtocol " + '"' + $($settingsObj.NetscalerProtocol) + '"' + " -NetScalerUser " + '"' +$($settingsObj.UserName) + '"' + " -NetscalerPassword " + '"' + $($settingsObj.Password) + '"' + " -BackupFileNamePrefix " + '"' + $($settingsObj.FileName) + '"' + " -BackupLevel " + '"' + $($settingsObj.BackupLevel) + '"' + " -PathToScp " + '"' + $($settingsObj.ScpPath) + '"' + " -BackupLocation " + '"' + $($settingsObj.BackupLocation) + '"' 
    <#
    else {
        $backupCommand = 'Backup-NetScaler ' + "-NetScalerIp " + '"' + $($settingsObj.NetscalerIp) + '"' + " -NetscalerProtocol " + '"' + $($settingsObj.NetscalerProtocol) + '"' + " -NetScalerUser " + '"' +$($settingsObj.UserName) + '"' + " -NetscalerPassword " + '"' + $($settingsObj.Password) + '"' + " -BackupFileNamePrefix " + '"' + $($settingsObj.FileName) + '"' + " -BackupLevel " + '"' + $($settingsObj.BackupLevel) + '"' + " -PathToScp " + '"' + $($settingsObj.ScpPath) + '"' + " -BackupLocation " + '"' + $($settingsObj.BackupLocation) + '"'
    }

    if ($wpfrButtonEmailYes.isChecked){

    }#>

    $scriptArr += $moduleLoadCommand
    $scriptArr += 'try {'
    $scriptArr += $backupCommand
    if ($wpfrButtonEmailYes.isChecked) {
        $scriptArr += 'Send-MailMessage ' + "-Body " + '"' + $($settingsObj.NetscalerIp) + '"' + " -From " + '"' + $($settingsObj.EmailFrom) + '"' + " -SmtpServer " + '"' + $($settingsObj.EmailSmtp) + '"' + " -Subject 'Success: NetScalerBackup' -To " + '"' + $($settingsObj.EmailTo) + '"'
    }
    $scriptArr += '}'
    $scriptArr += 'catch {'
    $scriptArr += '$PSItem.Exception.Message'
    if ($wpfrButtonEmailYes.isChecked) {
        $scriptArr += 'Send-MailMessage -Body $PSItem.Exception.Message -From ' + '"' + $($settingsObj.EmailFrom) + '"' + " -SmtpServer " + '"' + $($settingsObj.EmailSmtp) + '"' + " -Subject 'Error: NetScalerBackup' -To " + '"' + $($settingsObj.EmailTo) + '"'      
    }
    $scriptArr += '}'
    $scriptArr | Out-File $backupScriptFile -Force

    # END Generate Backup Script

    if (-not ($DebugMode)) {
        # START Run PS2Exe for Credential Hiding

        $currentDir = $PWD
        cd $PSScriptRoot
        $scriptFile = ".\PS2EXE-v0.5.0.0\ps2exe.ps1"
        #$targetFile = ".\Backup.exe"
        &$scriptFile -inputFile $backupScriptFile -outputfile $exeFilePath -noconsole

        Remove-Item $backupScriptFile -Confirm:$false -Force
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Script has been generated at: $backupScriptFile")
        ise $backupScriptFile
    }
    <#
    $ps2exePath = "'" + "$PSScriptRoot\PS2EXE-v0.5.0.0\ps2exe.ps1" + "'"
    $targetFile = "'" + "$PSScriptRoot\Backup.exe" + "'"
    #$ps2exeCommandPath = '"' + "$ps2exePath  + -inputFile $backupScriptFile -outputFile $PSScriptRoot\Backup.exe -noconsole" + '"'
    $ps2exeCommandPath = "$ps2exePath -inputFile $backupScriptFile -outputFile $targetFile -noconsole"
    [System.Windows.Forms.MessageBox]::Show($ps2execommandPath)
    # $ps2exeCommandPath | Out-File C:\Temp\command.txt
    Start-Process -FilePath "powershell.exe" -ArgumentList "-file $ps2exePath -inputFile $backupScriptFile -outputFile $targetFile -noconsole"
    # Remove-Item $backupScriptFile -Confirm:$false -Force
    #>

    # END Run PS2Exe for Credential Hiding

    # START Task Definition


    if ($Global:CredUserObj)
    {
        $taskObj = New-ScheduledJob -DayOfWeek $dayOfWeek -HourOfDay $hour `
            -RepetitionInterval $schedule -TaskName $taskName -OverwriteExisting -ActionCommand "$exeFilePath" -UserAccount $Global:CredUserObj
    }
    else
    {
        $taskObj = New-ScheduledJob -DayOfWeek $dayOfWeek -HourOfDay $hour `
            -RepetitionInterval $schedule -TaskName $taskName -OverwriteExisting -ActionCommand "$exeFilePath"
    }
    if ($taskObj)
    {
        [System.Windows.Forms.MessageBox]::Show("Task has been created!")
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Task has not been created. Please setup manually!")
    }

    # END Task Definition

    # START Reset GUI
    Reset-FormObject -FormObject $wpfgroupBoxScheduler
    Reset-FormObject -FormObject $wpfgroupBoxEmail
    # END Reset GuI
})

$WPFbuttonExport.Add_Click({

    $rButtonArr = @($WPFrButtonHttp,$WPFrButtonHttps)
    $rButtonArr | ForEach-Object {
        if ($_.isChecked)
        {
            $protocol = $_.Content
        }
    }


    $settingsObj = New-Object -TypeName PSCustomObject
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name NetscalerIP -Value $WPFtextboxNSIP.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name NetscalerProtocol -Value $protocol
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name UserName -Value $WPFtextboxUserName.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name Password -Value $WPFpasswordBox.Password.ToString()
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name BackupLevel -Value $WPFcomboBoxBackupLevel.SelectedValue
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name ScpPath -Value $WPFtextboxScpPath.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name BackupLocation -Value $WPFtextboxBackupLocation.Text
    Add-Member -InputObject $settingsObj -MemberType NoteProperty -Name FileName -Value $WPFtextboxFilePrefix.Text

    # Create Backup File
    $obj = New-Object System.Windows.Forms.SaveFileDialog
    $obj.FileName = "NetscalerSettings"
    $obj.DefaultExt = ".csv"
    if ($obj.ShowDialog())
    {
        try {
            $settingsObj | Export-Csv $obj.FileName -ErrorVariable csvError -NoTypeInformation
            [System.Windows.Forms.MessageBox]::Show("Settings have been stored at $($obj.FileName)!")
        }
        catch
        {
            [System.Windows.Forms.MessageBox]::Show($csvError)
        }
    }
})

$WPFbuttonImport.Add_Click({

    # Load Backup File
    # Create Backup File
    $obj = New-Object System.Windows.Forms.OpenFileDialog
    $obj.DefaultExt = ".csv"
    if ($obj.ShowDialog())
    {
        $importObj = Import-Csv $obj.FileName
    }

    if ( $importObj )
    {
        $WPFtextboxNSIP.Text = $importObj.NetscalerIP

        switch ($importObj.NetscalerProtocol)
        {
            "HTTP" {
                $WPFrButtonHttp.IsChecked = $true
            }
            "HTTPS" {
                $WPFrButtonHttps.IsChecked = $true
            }
        }

        $WPFtextboxUserName.Text = $importObj.UserName
        $WPFpasswordBox.Password = $importObj.Password
        $WPFtextboxScpPath.Text = $importObj.ScpPath
        $WPFcomboBoxBackupLevel.SelectedValue = $importObj.BackupLevel
        $WPFtextboxBackupLocation.Text = $importObj.BackupLocation
        $WPFtextboxFilePrefix.Text = $importObj.FileName
    }
})

$wpfrButtonEmailYes.Add_Click({
    #[System.Windows.MessageBox]::Show("Clicked")

    # Enable Mail Controls
    $wpftextboxEmailFrom.IsEnabled = $true
    $wpftextboxEmailTo.IsEnabled = $true
    $wpfTextboxSmtp.IsEnabled = $true
})

$wpfrButtonEmailNo.Add_Click({
    # Disable Mail Controls
    $wpftextboxEmailFrom.IsEnabled = $false
    $wpftextboxEmailTo.IsEnabled = $false
    $wpfTextboxSmtp.IsEnabled = $false

    # Clear Input
    # Enable Mail Controls
    $wpftextboxEmailFrom.Text = $null
    $wpftextboxEmailTo.Text = $null
    $wpfTextboxSmtp.Text = $null

})

$wpfButtonVerify.Add_Click({
    $verifyParam = @{
        ComputerName = $WPFtextboxNSIP.Text
    }
    if ($protocol -eq "HTTPS") {
        $verifyParam.Add("Secure",$True)
    }
    $connection = Test-NetScalerConnection @verifyParam

    if ($connection) {
        [System.Windows.Forms.MessageBox]::Show("HTTPS/S and SSH Connectivity have been verified")
        
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Connection refused")
    }
})
# END Event Handler Definition

#$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})

#===========================================================================
# Shows the form
#===========================================================================
write-host "To show the form, run the following" -ForegroundColor Cyan
'$Form.ShowDialog() | out-null'

$Form.ShowDialog() | out-null