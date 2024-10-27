Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Path to the JSON configuration file
$configFilePath = "$PSScriptRoot/config.json"

# Function to load the existing configuration from JSON
function Load-Config {
    if (Test-Path $configFilePath) {
        $configData = Get-Content -Path $configFilePath | ConvertFrom-Json
        return $configData
    }
    return $null
}

# Function to save the updated configuration back to JSON
function Save-Config {
    param (
        [hashtable]$configData
    )
    $configData | ConvertTo-Json -Depth 10 | Set-Content -Path $configFilePath -Force
}

# Create the form for updating configuration
$form = New-Object System.Windows.Forms.Form
$form.Text = "Update Configuration"
$form.Size = New-Object System.Drawing.Size(400, 350)  # Increased height for padding
$form.StartPosition = "CenterScreen"

# Load existing config
$config = Load-Config

# Check if config was loaded successfully
if ($config -eq $null) {
    # Show a message if no config file is found and close the form
    [System.Windows.Forms.MessageBox]::Show("Configuration file not found. Please create it first.", "Configuration Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    $form.Close()
    return
}

# Email Address
$labelEmail = New-Object System.Windows.Forms.Label
$labelEmail.Text = "Email Address:"
$labelEmail.AutoSize = $true
$labelEmail.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($labelEmail)

$emailTextbox = New-Object System.Windows.Forms.TextBox
$emailTextbox.Size = New-Object System.Drawing.Size(300, 20)
$emailTextbox.Location = New-Object System.Drawing.Point(20, 50)
$emailTextbox.Text = $config.EmailAddress
$form.Controls.Add($emailTextbox)

# Folder Path
$labelFolder = New-Object System.Windows.Forms.Label
$labelFolder.Text = "Folder Path:"
$labelFolder.AutoSize = $true
$labelFolder.Location = New-Object System.Drawing.Point(20, 80)
$form.Controls.Add($labelFolder)

$folderTextbox = New-Object System.Windows.Forms.TextBox
$folderTextbox.Size = New-Object System.Drawing.Size(300, 20)
$folderTextbox.Location = New-Object System.Drawing.Point(20, 110)
$folderTextbox.Text = $config.FolderPath
$form.Controls.Add($folderTextbox)

# Write CSV Checkbox
$writeCSVCheckbox = New-Object System.Windows.Forms.CheckBox
$writeCSVCheckbox.Text = "Write CSV"
$writeCSVCheckbox.Location = New-Object System.Drawing.Point(20, 140)
$writeCSVCheckbox.Checked = $config.WriteCSV
$form.Controls.Add($writeCSVCheckbox)

# Include Date in File Name Checkbox
$includeDateCheckbox = New-Object System.Windows.Forms.CheckBox
$includeDateCheckbox.Text = "Include Date in File Name"
$includeDateCheckbox.Location = New-Object System.Drawing.Point(20, 170)
$includeDateCheckbox.Checked = $config.IncludeDateInFileName
$form.Controls.Add($includeDateCheckbox)

# Include Summary Checkbox
$includeSummaryCheckbox = New-Object System.Windows.Forms.CheckBox
$includeSummaryCheckbox.Text = "Include Summary"
$includeSummaryCheckbox.Location = New-Object System.Drawing.Point(20, 200)
$includeSummaryCheckbox.Checked = $config.IncludeSummary
$form.Controls.Add($includeSummaryCheckbox)

# Silent Mode Checkbox
$silentModeCheckbox = New-Object System.Windows.Forms.CheckBox
$silentModeCheckbox.Text = "Silent Mode"
$silentModeCheckbox.Location = New-Object System.Drawing.Point(20, 230)
$silentModeCheckbox.Checked = $config.SilentMode
$form.Controls.Add($silentModeCheckbox)

# Send Mail Checkbox
$sendMailCheckbox = New-Object System.Windows.Forms.CheckBox
$sendMailCheckbox.Text = "Send Mail"
$sendMailCheckbox.Location = New-Object System.Drawing.Point(200, 230)
$sendMailCheckbox.Checked = $config.SendMail
$form.Controls.Add($sendMailCheckbox)

# Save Button
$saveButton = New-Object System.Windows.Forms.Button
$saveButton.Text = "Save"
$saveButton.Location = New-Object System.Drawing.Point(150, 270)  # Adjusted location for padding
$saveButton.Add_Click({
    # Validate Email Address
    if (-not [string]::IsNullOrWhiteSpace($emailTextbox.Text) -and
        -not $emailTextbox.Text -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid email address.", "Invalid Input", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Check if the folder path is not empty and exists
    if (-not [string]::IsNullOrWhiteSpace($folderTextbox.Text) -and -not (Test-Path $folderTextbox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("The specified folder path does not exist.", "Invalid Path", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Update the config hashtable with new values
    $updatedConfig = @{
        EmailAddress                = $emailTextbox.Text
        FolderPath                  = $folderTextbox.Text
        WriteCSV                    = $writeCSVCheckbox.Checked
        IncludeDateInFileName      = $includeDateCheckbox.Checked
        IncludeSummary              = $includeSummaryCheckbox.Checked
        SilentMode                  = $silentModeCheckbox.Checked
        SendMail                    = $sendMailCheckbox.Checked
    }

    try {
        # Save the updated config back to JSON
        Save-Config -configData $updatedConfig
        [System.Windows.Forms.MessageBox]::Show("Configuration saved successfully.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    } catch {
        # Improved error handling to show the actual error message
        $errorMessage = $_.Exception.Message  # Get the specific error message
        [System.Windows.Forms.MessageBox]::Show("An error occurred while saving the configuration: $errorMessage", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($saveButton)

# Show the form
$form.ShowDialog()
