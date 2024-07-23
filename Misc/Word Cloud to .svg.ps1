# This script uses a .txt file full of words to create a word cloud in the .svg format. It uses the PSWordCloud module.
# The PSWordCloud module says it is only compatible with PowerShell 7.0.0 and higher. Please check for errors if running a different version of PowerShell.

# Loads the System.Windows.Forms assembly needed for the file dialog boxes
Add-Type -AssemblyName System.Windows.Forms

# Imports the PSWordCloud module that draws the word cloud and saves it as a .svg file
Import-Module PSWordCloud

# Defines a list of PSWordCloud-compatible fonts that are native to Windows
$AllInstalledFonts = New-Object System.Drawing.Text.InstalledFontCollection

# A list of fancier PSWordCloud-compatible fonts that are native to Windows
$FancierFonts = @("Ink Free", "MV Boli", "Segoe Print", "Segoe Script")

# Introductory message to the user
Write-Host "`nThis script will take a .txt file full of words and convert it to a word cloud image in .svg format`n"
Start-Sleep -Seconds 3

# Prompts the user to enter the resolution of the image
$UserResolution = Read-Host "`nPlease enter the resolution of the word cloud image in WidthxHeight format (ex: 800x600)`n"
$Resolution = [System.String] $UserResolution

# Prompts the user to choose a random selection from all installed fonts, the list of fancier fonts, or manually specify a font
$FontListChoice = Read-Host "`nChoose 1 for a random font from the entire list of fonts (not all fonts are compatible with PSWordCloud)`nChoose 2 for a random fancier font`nChoose 3 to enter a font of your choosing`n"

# Switch statement to handle the $FontListChoice
Switch ($FontListChoice) {
    1 {
        $SelectedFont = Get-Random -InputObject $AllInstalledFonts.Families.Name
    }
    2 {
        $SelectedFont = Get-Random -InputObject $FancierFonts
    }
    3 {
        $SelectedFont = Read-Host "Please enter the font you wish to use"
    }
    Default {
        Write-Host "`nInvalid input. `nPlease enter 1 for a random font from the entire list of fonts or 2 for a random fancier font."
        # Re-prompt the user
        $FontListChoice = Read-Host "`nChoose 1 for a random font from the entire list of fonts`nChoose 2 for a random fancier font`n"
    }
}

# Converts the $SelectedFont value to a string since the PSWordCloud module is picky with formatting
$SelectedFontString = [System.String]$SelectedFont

# Asks the user if they want to use a "focus word" that will be in the center of the word cloud
$FocusWordPrompt = Read-Host "`nWould you like a focus word (or phrase) to appear in the center of the word cloud? (Yes/No) `n"
$FocusWord = $null
if ($FocusWordPrompt -match "^(Yes|yes|Y|y)$") {
    $FocusWord = Read-Host "`nPlease enter the focus word (or phrase)"`n
}

# Create an OpenFileDialog object to choose the source .txt file
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Title = "Select a Text File"
    Filter = 'Text files (*.txt)|*.txt'
}

# Lets the user know the .txt file selection will happen shortly
Write-Host `n"Please select a .txt file for the word source for the word cloud`n"
Start-Sleep -Seconds 3

# Show the OpenFileDialog box
$null = $OpenFileDialog.ShowDialog()

# Check if a .txt file was selected for the word cloud source
If ($OpenFileDialog.FileName) {
    # Create a SaveFileDialog object to choose the save location
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
        Title = "Save the Word Cloud as SVG"
        FileName = 'Word Cloud.svg'
        Filter = 'SVG files (*.svg)|*.svg'
    }

    # Lets the user know the .svg file save location selection will happen shortly
    Write-Host "Please select a save location for the word cloud .svg image file`n"
    Start-Sleep -Seconds 3
    Write-Host "The chosen font is: $SelectedFont`n"
    Write-Host "Your word cloud is being created. This may take a little while. Note the messages below for any skipped words.`n"

    # Show the SaveFileDialog box
    $null = $SaveFileDialog.ShowDialog()

    # Check if a save location was selected
    If ($SaveFileDialog.FileName) {
        # Define a hashtable for word sizes
        $WordFontSize = @{}

        # Get the content of the file
        $words = Get-Content $OpenFileDialog.FileName

        # Populate the hashtable with words and their random sizes
        foreach ($word in $words) {
            $WordFontSize[$word] = Get-Random -Maximum 20 -Minimum 1
        }

        # Create the word cloud with -ImageSize set to the user-specified $Resolution and -AllowRotaton to All
        if ($null -ne $FocusWord) {
            New-WordCloud -Path $SaveFileDialog.FileName -ImageSize $Resolution -Typeface $SelectedFontString -WordSizes $WordFontSize -AllowRotation All -FocusWord $FocusWord
        } else {
            New-WordCloud -Path $SaveFileDialog.FileName -ImageSize $Resolution -Typeface $SelectedFontString -WordSizes $WordFontSize -AllowRotation All
        }
    }
}

# Displays the file save location in the PowerShell console for 5 seconds before exiting the script
Write-Host "Your word cloud image was saved at the following location:" $SaveFileDialog.FileName `n
Start-Sleep -Seconds 5