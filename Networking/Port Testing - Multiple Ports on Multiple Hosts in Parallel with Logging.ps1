# This script will test mulitple TCP ports on multiple hosts (while using ForEach-Object loops in parallel) and log the results to an Excel spreadsheet


# Loads the Windows Forms assembly needed for the file dialog boxes and the ImportExcel module for saving the results spreadsheet
Add-Type -AssemblyName System.Windows.Forms
Import-Module ImportExcel

# Introduction message
Write-Host "This script will test a list of multiple ports on a list of multiple hosts (IP addresses, hostnames, etc.)"
Write-Host "Please note that this will only test TCP ports and not UDP ports"
Start-Sleep -Seconds 3

# Function to open a file dialog box and select a file
Function Select-File {
    Param (
        [string]$Title
    )
    $FileOpenDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileOpenDialog.Title = $Title
    $FileOpenDialog.Filter = "All Files (*.*)|*.*"
    If ($FileOpenDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Return $FileOpenDialog.FileName
    }
    Return $Null
}

# Function to open a save file dialog box and select a save location
Function Save-File {
    $SaveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveDialog.Title = "Save the results Excel spreadsheet"
    $SaveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $SaveDialog.FileName = "Port Testing Results.xlsx"
    If ($SaveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Return $SaveDialog.FileName
    }
    Return $Null
}

# Selects the hosts list file
Write-Host `n"Please select the list of hosts to test (.txt, .csv, or .xlsx)"
Write-Host "For .csv and .xlsx files, make sure cell A1 has the heading Host"
Start-Sleep -Seconds 3
$HostsListFilePath = Select-File -title "Select the hosts list"
If (-not $HostsListFilePath) {
    Write-Host `n"No input file was selected. Exiting script."
    Start-Sleep -Seconds 3
    Exit
}
Write-Host `n"Hosts list file: $HostsListFilePath"
Start-Sleep -Seconds 3

# Selects the ports lists file
Write-Host `n"Please select the list of ports to test (.txt, .csv, or .xlsx)"
Write-Host "For .csv and .xlsx files, make sure cell A1 has the heading Port"
Start-Sleep -Seconds 3
$PortsListFilePath = Select-File -title "Select the ports list"
If (-not $PortsListFilePath) {
    Write-Host `n"No port file was selected. Exiting script."
    Start-Sleep -Seconds 3
    Exit
}
Write-Host `n"Ports list file: $PortsListFilePath"
Start-Sleep -Seconds 3

# Selects the save location for the results spreadsheet
Write-Host `n"Please select the save location for the results spreadsheet"
Start-Sleep -Seconds 3
$ResultsSpreadsheetFilePath = Save-File
If (-not $ResultsSpreadsheetFilePath) {
    Write-Host `n"No save location selected. Exiting script."
    Start-Sleep -Seconds 3
    Exit
}

# Determines the file type based on extension
$HostsListFileType = [System.IO.Path]::GetExtension($HostsListFilePath).TrimStart('.').ToLower()

# Reads the hosts list based on file type
Switch ($HostsListFileType) {
    "txt" { $Hosts = Get-Content -Path $HostsListFilePath }
    "csv" { $Hosts = Import-Csv -Path $HostsListFilePath | Select-Object -ExpandProperty Host }
    "xlsx" { $Hosts = Import-Excel -Path $HostsListFilePath | Select-Object -ExpandProperty Host }
    Default { Write-Host `n"Unsupported file type. Exiting script."; Exit }
}

# Determine file type based on extension for ports
$PortsListFileType = [System.IO.Path]::GetExtension($PortsListFilePath).TrimStart('.').ToLower()

# Reads the ports list based on file type
Switch ($PortsListFileType) {
    "txt" { $Ports = Get-Content -Path $PortsListFilePath }
    "csv" { $Ports = Import-Csv -Path $PortsListFilePath | Select-Object -ExpandProperty Port }
    "xlsx" { $Ports = Import-Excel -Path $PortsListFilePath | Select-Object -ExpandProperty Port }
    Default { Write-Host `n"Unsupported file type. Exiting script."; Exit }
}

# Combines the hosts and ports into a single array for parallel processing
$HostPortPairs = ForEach ($HostName in $Hosts) {
    ForEach ($Port in $Ports) {
        [PSCustomObject]@{HostName = $HostName; Port = $Port}
    }
}

# Uses ForEach-Object -Parallel to test connections faster and adds a timestamp
$Results = $HostPortPairs | ForEach-Object -Parallel {
    $Result = Test-NetConnection -ComputerName $_.HostName -Port $_.Port
    $Result | Add-Member -MemberType NoteProperty -Name "Timestamp" -Value (Get-Date).ToString("MM/dd/yyyy HH:mm:ss tt")
    $Result
} -ThrottleLimit 10

# Selects only the required columns and renames them
# Additional columns can be added back by adding to this list
$FilteredResults = $Results | Select-Object @{Name="Computer Name";Expression={$_.ComputerName}},
                                        @{Name="Remote Address";Expression={$_.RemoteAddress}},
                                        @{Name="Ping Succeeded";Expression={$_.PingSucceeded}},
                                        @{Name="TCP Port";Expression={$_.RemotePort}},
                                        @{Name="Port Test Succeeded";Expression={$_.TcpTestSucceeded}},
                                        @{Name="Timestamp";Expression={$_.Timestamp}}

# Reorders the columns for legibility
$FilteredResults = $FilteredResults | Select-Object "Computer Name", "Remote Address", "Ping Succeeded", "TCP Port", "Port Test Succeeded", "Timestamp"

# Saves the results to an Excel spreadsheet with formatting and table options
$FilteredResults | Export-Excel -Path $ResultsSpreadsheetFilePath -WorksheetName "Port Testing Results" -TableName "Port_Testing_Results" -TableStyle Medium9 -AutoSize

Write-Host `n"The results spreadsheet has been saved to $ResultsSpreadsheetFilePath"
Start-Sleep -Seconds 3
Write-Host `n"This script will exit shortly"
Start-Sleep -Seconds 5