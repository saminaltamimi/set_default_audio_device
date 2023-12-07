Add-Type -AssemblyName System.Windows.Forms

# Global variable to store the selected audio device
$soundVolumeViewPath = ".\soundvolumeview"
if (-not (Test-Path $soundVolumeViewPath)) {
    # Download the latest version of SoundVolumeView
    Invoke-WebRequest -Uri "https://www.nirsoft.net/utils/soundvolumeview.zip" -OutFile "soundvolumeview.zip"
    # Extract the zip file
    Expand-Archive -Path "soundvolumeview.zip" -DestinationPath $soundVolumeViewPath
    # Delete the zip file
    Remove-Item -Path "soundvolumeview.zip" -Force
}

# Export the command line friendly items to a JSON file
$jsonFilePath = ".\soundvolumeview_items.json"
& "$soundVolumeViewPath\SoundVolumeView.exe" /sjson $jsonFilePath
Start-Sleep -Seconds 3

# Read the JSON file
$items = Get-Content $jsonFilePath | ConvertFrom-Json

# Create an array to store the headphone devices
$headphoneDevices = @()
$items | Where-Object { $_.'Command-Line Friendly ID' -like '*Device*' -and $_.'Command-Line Friendly ID' -like '*Render*' } | ForEach-Object {
    $headphoneDevices += $_.'Device Name' 
}

# Create an array to store the microphone devices
$microphoneDevices = @()
$items | Where-Object { $_.'Command-Line Friendly ID' -like '*Device*' -and $_.'Command-Line Friendly ID' -like '*Capture*' } | ForEach-Object {
    $microphoneDevices += $_.'Device Name' 
}

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Audio Device Settings"
$form.Size = New-Object System.Drawing.Size(400, 225)
$form.StartPosition = "CenterScreen"

# Create the first dropdown menu
$dropdown1 = New-Object System.Windows.Forms.ComboBox
$dropdown1.Location = New-Object System.Drawing.Point(100, 30)
$dropdown1.Size = New-Object System.Drawing.Size(200, 25)
$dropdown1.Items.AddRange($headphoneDevices)
$dropdown1.SelectedIndex = 0
$form.Controls.Add($dropdown1)

# Create the label for the first dropdown menu
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(50, 30)
$label1.Size = New-Object System.Drawing.Size(40, 25)
$label1.Text = "Headphone:"
$form.Controls.Add($label1)

# Create the second dropdown menu
$dropdown2 = New-Object System.Windows.Forms.ComboBox
$dropdown2.Location = New-Object System.Drawing.Point(100, 70)
$dropdown2.Size = New-Object System.Drawing.Size(200, 25)
$dropdown2.Items.AddRange($microphoneDevices)
$dropdown2.SelectedIndex = 0
$form.Controls.Add($dropdown2)

# Create the label for the second dropdown menu
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(50, 70)
$label2.Size = New-Object System.Drawing.Size(40, 25)
$label2.Text = "Microphone:"
$form.Controls.Add($label2)

# Create the OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(150, 120)
$okButton.Size = New-Object System.Drawing.Size(60, 25)
$okButton.Text = "OK"

# Create the status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 160)
$statusLabel.Size = New-Object System.Drawing.Size(200, 25)
$form.Controls.Add($statusLabel)

# Event handler for the OK button's Click event
$okButton.Add_Click({
    $selectedHeadphoneDevice = $dropdown1.SelectedItem
    $selectedMicrophoneDevice = $dropdown2.SelectedItem
    $statusLabel.Text = "Processing: $selectedHeadphoneDevice, $selectedMicrophoneDevice"
    Write-Host "Selected Device: $selectedHeadphoneDevice"
    Write-Host "Selected Option: $selectedMicrophoneDevice"

    # Read the JSON file and display the command line friendly items
    $jsonAudioDevices = Get-Content $jsonFilePath | ConvertFrom-Json 
    $selectedHeadphoneForSoundVolumeView = $jsonAudioDevices | Where-Object { $_.'Device Name' -eq $selectedHeadphoneDevice -and $_.'Command-Line Friendly ID' -like '*Device*' -and $_.'Command-Line Friendly ID' -like '*Render*' } | Select-Object -ExpandProperty 'Command-Line Friendly ID' 
    $selectedMicrophoneForSoundVolumeView = $jsonAudioDevices | Where-Object { $_.'Device Name' -eq $selectedMicrophoneDevice -and $_.'Command-Line Friendly ID' -like '*Device*' -and $_.'Command-Line Friendly ID' -like '*Capture*' } | Select-Object -ExpandProperty 'Command-Line Friendly ID'

    # Set the audio device
    $statusLabel.Text = "Set the $selectedHeadphoneDevice, $selectedMicrophoneDevice as default"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetDefault "$selectedHeadphoneForSoundVolumeView" all
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetDefault "$selectedMicrophoneForSoundVolumeView" all

    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAppDefault "$selectedHeadphoneForSoundVolumeView" all "chrome.exe"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAppDefault "$selectedHeadphoneForSoundVolumeView" all "msedge.exe"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAppDefault "$selectedHeadphoneForSoundVolumeView" all "firefox.exe"

    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAppDefault "$selectedMicrophoneForSoundVolumeView" all "chrome.exe"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAppDefault "$selectedMicrophoneForSoundVolumeView" all "msedge.exe"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAppDefault "$selectedMicrophoneForSoundVolumeView" all "firefox.exe"
    Start-Sleep -Seconds 3

    # Ensure that the audio device is unmuted
    $statusLabel.Text = "Unmute $selectedHeadphoneDevice, $selectedMicrophoneDevice"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /Unmute "$selectedHeadphoneForSoundVolumeView" 
    & "$soundVolumeViewPath\SoundVolumeView.exe" /Unmute "$selectedMicrophoneForSoundVolumeView"
    Start-Sleep -Seconds 3

    # Disable exclusive mode 
    $statusLabel.Text = "Disabled exclusive mode for $selectedHeadphoneDevice, $selectedMicrophoneDevice"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAllowExclusive "$selectedHeadphoneForSoundVolumeView" 0
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetAllowExclusive "$selectedMicrophoneForSoundVolumeView" 0
    Start-Sleep -Seconds 3

    # Disable exclusive priority
    $statusLabel.Text = "Disabled exclusive priority for $selectedHeadphoneDevice, $selectedMicrophoneDevice"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetExclusivePriority "$selectedHeadphoneForSoundVolumeView" 0
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetExclusivePriority "$selectedMicrophoneForSoundVolumeView" 0
    Start-Sleep -Seconds 3

    # Set the volume to 100%
    $statusLabel.Text = "Maxed volume for $selectedHeadphoneDevice, $selectedMicrophoneDevice"
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetVolume "$selectedHeadphoneForSoundVolumeView" 100
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetVolume "$selectedMicrophoneForSoundVolumeView" 100

    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetVolume "chrome.exe" 100
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetVolume "msedge.exe" 100
    & "$soundVolumeViewPath\SoundVolumeView.exe" /SetVolume "firefox.exe" 100
    Start-Sleep -Seconds 3

    $statusLabel.Text = "Done!"
    Start-Sleep -Seconds 3
    $form.Close()
})

$form.Controls.Add($okButton)

# Show the form
$form.ShowDialog()

# Delete the JSON file
Remove-Item -Path $jsonFilePath -Force
Remove-Item -Path $soundVolumeViewPath -Recurse -Force
