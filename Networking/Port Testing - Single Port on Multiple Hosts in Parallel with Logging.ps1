# This script will test a single TCP port on multiple hosts (while using ForEach-Object loops in parallel) and log the results to an Excel spreadsheet


# Loads the Windows Forms assembly needed for the file dialog boxes and the ImportExcel module for saving the results spreadsheet
Add-Type -AssemblyName System.Windows.Forms
Import-Module ImportExcel

# Introduction message
Write-Host "This script will test a single TCP port on a list of multiple hosts (IP addresses, hostnames, etc.)"
Write-Host "Please note that this will only test TCP ports and not UDP ports"
Start-Sleep -Seconds 3

# Function to open a file dialog box and select a file
Function Select-File {
    Param (
        [string]$Title
    )
    $FileOpenDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileOpenDialog.Title = "Select the hosts list"
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

# Function to prompt user for port number and validate input
Function Get-ValidPortNumber {
    While ($True) {
        $PortNumber = Read-Host `n"Please enter the port number to test (0-65535)"
        If ($PortNumber -match '^\d+$' -and [int]$PortNumber -ge 0 -and [int]$PortNumber -le 65535) {
            Return [int]$PortNumber
            Write-Host `n"You will be testing port $PortNumber"
        } Else {
            Write-Host "Invalid port number. Please enter a valid number between 0 and 65535."
        }
    }
}

# Prompt user to enter the port number
$PortNumber = Get-ValidPortNumber

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

# Reads the hosts based on file type
Switch ($HostsListFileType) {
    "txt" { $Hosts = Get-Content -Path $HostsListFilePath }
    "csv" { $Hosts = Import-Csv -Path $HostsListFilePath | Select-Object -ExpandProperty Host }
    "xlsx" { $Hosts = Import-Excel -Path $HostsListFilePath | Select-Object -ExpandProperty Host }
    Default { Write-Host `n"Unsupported file type. Exiting script."; Exit }
}

# Uses a ForEach-Object -Parallel loop to test connections and add a timestamp
$Results = $Hosts | ForEach-Object -Parallel {
    $Result = Test-NetConnection -ComputerName $_ -Port $using:portNumber
    $Result | Add-Member -MemberType NoteProperty -Name "Timestamp" -Value (Get-Date).ToString("MM/dd/yyyy HH:mm:ss tt")
    $Result
}

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