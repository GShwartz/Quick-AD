# Get the directory of the script file
$scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
$passBuilderModuleFileName = "PasswordSettings.psm1"

# Combine the script directory and file names to get path
$moduleFileName = Join-Path -Path $scriptDirectory -ChildPath $passBuilderModuleFileName

# Import Module
Import-Module $moduleFileName -Force

# Function to create the About window
function ShowAboutWindow {
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "QuicK-AD"
    $aboutForm.Width = 300
    $aboutForm.Height = 200
    $aboutForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $aboutForm.MaximizeBox = $false
    $aboutForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Text = "Version: $($global:version)"
    $labelVersion.Location = New-Object System.Drawing.Point(10, 20)
    $labelVersion.AutoSize = $true

    $linkGitHub = New-Object System.Windows.Forms.LinkLabel
    $linkGitHub.Text = "GitHub"
    $linkGitHub.Location = New-Object System.Drawing.Point(10, 50)
    $linkGitHub.AutoSize = $true
    $linkGitHub.Add_LinkClicked({
        [System.Diagnostics.Process]::Start("https://github.com/GShwartz/Quick-AD/tree/main")
    })

    $labelCopyright = New-Object System.Windows.Forms.Label
    $labelCopyright.Text = "© 2024 Gil Shwartz"
    $labelCopyright.Location = New-Object System.Drawing.Point(10, 80)
    $labelCopyright.AutoSize = $true

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $buttonOK.Location = New-Object System.Drawing.Point(100, 120)

    # Add controls to the form
    $aboutForm.Controls.Add($labelVersion)
    $aboutForm.Controls.Add($labelScripter)
    $aboutForm.Controls.Add($linkGitHub)
    $aboutForm.Controls.Add($labelCopyright)
    $aboutForm.Controls.Add($buttonOK)

    # Show the form
    $aboutForm.ShowDialog()
}

# Function to show the top menu
function CreateMenuStrip {
    # Create a top menu strip
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    $menuStrip.Height = 4
    $menuStrip.BackColor = [System.Drawing.Color]::Wheat
    $menuStrip.Padding = New-Object System.Windows.Forms.Padding(1)

    # File menu
    $fileMenu = $menuStrip.Items.Add("File")
    $fileMenu.Padding = New-Object System.Windows.Forms.Padding(0)
    
    # Settings menu
    $settingsMenuItem = $fileMenu.DropDownItems.Add("Settings")
    $settingsMenuItem.Padding = New-Object System.Windows.Forms.Padding(0)

    # Password Strength sub-menu
    $passwordStrengthMenuItem = $settingsMenuItem.DropDownItems.Add("Password Strength")
    $passwordStrengthMenuItem.Add_Click({
        #Write-Host "Work in progress."
        SetPasswordStrength
    })

    $exitMenuItem = $fileMenu.DropDownItems.Add("Exit")
    $exitMenuItem.Add_Click({
        # Exit the application
        [Windows.Forms.Application]::Exit()
    })

    # Help menu
    $helpMenu = $menuStrip.Items.Add("Help")
    $helpMenu.Padding = New-Object System.Windows.Forms.Padding(0)
    $helpMenuAbout = $helpMenu.DropDownItems.Add("About")
    $helpMenuAbout.Add_Click({
        ShowAboutWindow
    })

    # Add a Paint event handler
    $menuStrip.Add_Paint({
        param($sender, $e)

        $borderColor = [System.Drawing.Color]::Black
        $borderWidth = 0.2

        $rect = $sender.ClientRectangle
        $rect.Width--
        $rect.Height--

        $pen = New-Object System.Drawing.Pen $borderColor, $borderWidth
        $e.Graphics.DrawRectangle($pen, $rect)

        $pen.Dispose()
    })

    return $menuStrip
}