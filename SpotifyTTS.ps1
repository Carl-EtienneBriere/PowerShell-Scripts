<#
.SYNOPSIS
    This PowerShell script adds an icon to the system tray that announces the currently playing track on Spotify.

.DESCRIPTION
    - It uses a custom icon in the Windows notification area.
    - A context menu allows enabling or disabling the voice announcement of tracks.
    - It retrieves the currently playing track on Spotify and reads it aloud using Windows Text-to-Speech.
    - A "Mute" option allows temporarily disabling track announcements.
    - An "Exit" option allows properly closing the application.

.PARAMETERS
    No parameters are required.

.REQUIREMENTS
    - Spotify must be installed and running.
    - PowerShell must support Windows Forms and System.Speech (available by default on Windows).

.NOTES
    Créé par / Created by : Carl-Étienne Brière    
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Speech

# Create a notification icon
$NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
#$NotifyIcon.Icon = [System.Drawing.SystemIcons]::Information
$NotifyIcon.Text = "SpotifyTTS"
$NotifyIcon.Visible = $false

$IconBase64 = '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'
$IconBytes = [Convert]::FromBase64String($IconBase64)
$Stream = [System.IO.MemoryStream]::new($IconBytes, 0, $IconBytes.Length)
$NotifyIcon.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Stream).GetHIcon()))

# Create the speech synthesizer
$SpeechBot = New-Object System.Speech.Synthesis.SpeechSynthesizer
$SpeechBot.Rate = 1

# Global variables
$global:Muted = $false
$global:NowPlaying = $null

# Create the context menu
$ContextMenu = New-Object System.Windows.Forms.ContextMenu
$CheckboxMenuItem = New-Object System.Windows.Forms.MenuItem "Mute"
$ExitMenuItem = New-Object System.Windows.Forms.MenuItem "Exit"

# Configure the context menu
$CheckboxMenuItem.Checked = $false  # Default unchecked
$ContextMenu.MenuItems.Add($CheckboxMenuItem)
$ContextMenu.MenuItems.Add("-")  # Separator line
$ContextMenu.MenuItems.Add($ExitMenuItem)
$NotifyIcon.ContextMenu = $ContextMenu

# Action for the checkbox
$CheckboxMenuItem.add_Click({
    $CheckboxMenuItem.Checked = -not $CheckboxMenuItem.Checked  # Toggle state
    $global:Muted = $CheckboxMenuItem.Checked
})

# Action for the "Exit" option
$ExitMenuItem.add_Click({
    $NotifyIcon.Visible = $false  # Hide the icon
    $NotifyIcon.Dispose()         # Cleanup resources
    [System.Windows.Forms.Application]::Exit()  # Close the application
    [System.GC]::Collect()  # Force garbage collection
    [System.GC]::WaitForPendingFinalizers()  # Ensure all finalizers have completed
})


# Configure a Timer to check Spotify's track title
$Timer = New-Object System.Windows.Forms.Timer
$Timer.Interval = 500  # 500 ms
$Timer.add_Tick({
    If (-not $global:Muted) {
        $CurrentTrack = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 -and $_.Name -eq "Spotify" } | Select-Object -ExpandProperty MainWindowTitle

        If ($CurrentTrack -and ($global:NowPlaying -ne $CurrentTrack)) {
            $global:NowPlaying = $CurrentTrack
            If ($CurrentTrack -notlike "*Spotify*") {
                $Artist, $Song = $CurrentTrack -split '-'
                $SpeechBot.Speak("$Song by $Artist")
            }
        }
    }
})

$Timer.Start()
Start-Sleep -Seconds 5
$NotifyIcon.Visible = $true
[System.Windows.Forms.Application]::Run()