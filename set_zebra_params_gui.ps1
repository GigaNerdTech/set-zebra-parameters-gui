# Set Zebra printing language script
# Joshua Woleben
# Written 9/17/2019

Import-Module ActiveDirectory

# GUI Code

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Zebra Printer Fixer" Height="500" Width="450" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="PrinterLabel" Content="Printers to send command to (separated by commas):"/>
        <TextBox x:Name="PrinterListTextBox" Height="75"/>
        <Button x:Name="SendCommandsButton" Content="Send Commands" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls
$PrinterListTextBox = $global:Form.FindName('PrinterListTextBox')
$CommandButton = $global:Form.FindName('SendCommandsButton')

$username = $env:USERNAME
$email_user = (Get-ADUser -Properties EmailAddress -Identity $username | Select -ExpandProperty EmailAddress)
$CommandButton.Add_Click({
$functions = @'
function is_printer_done {
    Param($printers_done,$printer_name_check)
    foreach ($printer_name in $printers_done) {
        if ($printer_name_check -match $printer_name) {
            return $true
        }
    }
    return $false
}
function is_printer_online {
    Param($printer_name)
    if (Test-Connection -ComputerName $printer_name -Quiet -Count 1) {
        return $true
    }
    else {
        return $false
    }
}
function write_log {
    Param([string]$log_entry,
            [string]$TranscriptFile)

            $mutex_name = 'Mutex for handling log file'
            $mutex = New-Object System.Threading.Mutex($false, $mutex_name)
            $mutex.WaitOne(-1) | out-null

            try {
                Add-Content $TranscriptFile -Value $log_entry
                Write-Host $log_entry
                
            }
            finally {
                $mutex.ReleaseMutex() | out-null

            }
            $mutex.Dispose()
}
'@
function write_log {
    Param([string]$log_entry,
            [string]$TranscriptFile)

            $mutex_name = 'Mutex for handling log file'
            $mutex = New-Object System.Threading.Mutex($false, $mutex_name)
            $mutex.WaitOne(-1) | out-null

            try {
                Add-Content $TranscriptFile -Value $log_entry
                Write-Host $log_entry
 

            }
            finally {
                $mutex.ReleaseMutex() | out-null

            }
            $mutex.Dispose()
}


# List of EPS servers
$print_servers = @('print_server_host')
$printer_list = @()
$log_body = "\\hostname\C`$\Temp\Zebra_Language_Set_$(get-date -f MMddyyyyHHmmss).txt"
$printer_status_file = ($env:TEMP + "\printer_status.txt")

write_log "Initializing...`n" $log_body
$global:printer_count = 0

ForEach ($server in $print_servers) {
    
 
    # Error checking
    if ([string]::IsNullOrEmpty($PrinterListTextBox.Text)) {
        [System.Windows.MessageBox]::Show('No printers specified!')
        return
    }

    $printer_names = $PrinterListTextBox.Text.split(",")
    ForEach ($printer_name in $printer_names) {
         $formatted_printer_name = $printer_name.Trim()
         try {
            $printer_list += Get-Printer -ComputerName $server -Name "$formatted_printer_name"
        }
        catch {
            write_log ("Printer $formatted_printer_name encountered an error being retrieved!`n") $log_body
            write_log "$formatted_printer_name, ERROR`n"  $printer_status_file 
        }
    }
    # Set commands
    "! U1 setvar `"device.languages`" `"zpl`"`n`n" | Out-File "\\$server\C`$\Temp\zpl1.zpl" -Encoding ascii
    "~jc^xa^jus^xz`n`n" | Out-File "\\$server\C`$\Temp\zpl2.zpl" -Encoding ascii

  
    $printer_list | ForEach-Object {

        
        $printer=$_.Name
        $PrintJob = { 
        Invoke-Expression $using:functions
            
        write_log "Working on $using:printer..." $using:log_body 
        if ((is_printer_online $using:printer) -eq $false) {
            write_log ($using:printer + " is offline, skipping...`n") $using:log_body
            write_log "$using:printer, OFFLINE`n" $using:printer_status_file 
            
        }
        else {
       
            write_log "Changing $using:printer driver to generic / text...`n" $using:log_body
            try { 
                Set-Printer -ComputerName $using:server -Name $using:printer -DriverName "Generic / Text Only"
            }
            catch {
                write_log "$using:printer encountered error switching driver!`n" $using:log_body
                write_log "$using:printer, DRIVER SWITCH ERROR`n" $using:printer_status_file
                break
            }

            sleep 1
            
            $batch_text = "C:\Windows\ssdal.exe /p `"$using:printer`" send `"C:\Temp\zpl1.zpl`"`nC:\Windows\ssdal.exe /p `"$using:printer`" send `"C:\Temp\zpl2.zpl`""
            try {
                $batch_text | Out-File "\\$using:server\C`$\Temp\printer-$using:printer.cmd" -Encoding ascii
            }
            catch {
                write_log "$using:printer encountered error writing command file!`n" $using:log_body
                write_log "$using:printer, COMMAND FILE ERROR`n" $using:printer_status_file
                break
            }
            write_log "Issuing zpl commands to printer $using:printer..." $using:log_body
            $printer_com = $($using:printer)
           try {
            $command_output =  Invoke-Command -ArgumentList $printer_com -ComputerName $using:server { Param($printer_com)

              &  "C:\Temp\printer-$printer_com.cmd"
            }
            } catch {
                write_log "$using:printer encountered error executing command file!`n" $using:log_body
                write_log "$using:printer, COMMAND FILE EXECUTE ERROR`n" $using:printer_status_file
                break
            }
            write_log $command_output $using:log_body 
            sleep 1
            write_log "Switching $using:printer back to Zebra driver...`n" $using:log_body 
            try {
                Set-Printer -ComputerName $using:server -Name $using:printer -DriverName "ZDesigner QLn220"
            }
            catch {
                write_log "$using:printer encountered error switching driver!`n" $using:log_body
                write_log "$using:printer, DRIVER SWITCH ERROR`n" $using:printer_status_file
            }

#            $using:printer | Out-File -FilePath $using:printer_done_file -Append
            if (Test-Path -Path "\\$using:server\C`$\Temp\printer-$using:printer.cmd" -PathType Leaf) {
                   Remove-Item -Path "\\$using:server\C`$\Temp\printer-$using:printer.cmd" -Force
            }
            
            write_log "$using:printer, SUCCESS`n" $using:printer_status_file
        }
        } # End script block
        Start-Job -ScriptBlock $PrintJob
        while (@(Get-Job).Count -gt 30) {
            $log_output = Receive-Job -Keep
            write_log $log_output $log_body
            
            Remove-Job -State Completed
            sleep 5
        }


    }
    write_log "Waiting on jobs to finish..." $log_body

    $final_output = Get-Job | Wait-Job | Receive-Job
    write_log $final_output $log_body
    Get-Job | Wait-Job | Remove-Job
    write_log "Cleaning up..." $log_body
    Remove-Item -Path "\\$server\C`$\Temp\zpl1.zpl" -Force
    Remove-Item -Path "\\$server\C`$\Temp\zpl2.zpl" -Force
     
}

# Generate email report
$email_list=@("email1@example.com", $email_user)
$subject = "Zebra Printer Report called by $username"
$body = @()

$body += "Zebra language set report and transcript attached.`n`n"

if ($all) {
    $body += "`nScript ran against ALL online printers.`n`n"
}
if (-not [string]::IsNullOrEmpty($File)) {
    $body += "`nPrinter list loaded from file $File`n`n"
}
if (-not [string]::IsNullOrEmpty($PrinterList)) {
    $body += "`nPrinter list supplied at GUI.`n`n"
}

$status_log = Get-Content -Path $printer_status_file
$table_body = "<table border=`"3`"><thead><tr><th>Printer Name</th><th>Status</th></tr></thead><tbody>"
foreach ($line in $status_log) {
    if  ($line -match "ERROR") {
        $bgcolor = "red"
    }
    elseif ($line -match "OFFLINE") {
        $bgcolor = "yellow"
    }
    elseif ($line -match "SUCCESS") {
        $bgcolor = "green"
    }
    else {
        $bgcolor = "white"
    }
    $row = $line -split ','
    $table_body += ("<tr bgcolor=`"" + $bgcolor +"`"><td>" + $row[0] + "</td><td>" + $row[1] + "</td></tr>")
}
$table_body += "</tbody></table>"
$body += $table_body

$MailMessage = @{
    To = $email_list
    From = "ZebraLanguageReport<Donotreply@example.com>"
    Subject = $subject
    Body = ($body -join "<br/>")
    SmtpServer = "smtp.mhd.com"
    Attachment = $log_body 
    ErrorAction = "Stop"
}
Send-MailMessage @MailMessage -BodyAsHtml
Remove-Item -Path $printer_status_file -Force
write_log "Done!" $log_body
$PrinterListTextBox.Text = ""
$global:Form.InvalidateVisual()
}) # End button scriptblock
$global:Form.Add_Closing({
    # Clean up GUI
    Get-Variable | Where {$_.PSProvider -notmatch "Core" } | Remove-Variable -ErrorAction SilentlyContinue -Force 
})
# Show GUI
$global:Form.ShowDialog() | out-null