# Function to create the About window
function ShowAboutWindow {
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "QuicK-AD"
    $aboutForm.Width = 300
    $aboutForm.Height = 200
    $aboutForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $aboutForm.MaximizeBox = $false
    $aboutForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $labelScripter = New-Object System.Windows.Forms.Label
    $labelScripter.Text = "Written by: Gil Shwartz"
    $labelScripter.Location = New-Object System.Drawing.Point(10, 10)
    $labelScripter.AutoSize = $true

    $labelVersion = New-Object System.Windows.Forms.Label
    $labelVersion.Text = "Version: $($global:version)"
    $labelVersion.Location = New-Object System.Drawing.Point(10, 40)

    $linkGitHub = New-Object System.Windows.Forms.LinkLabel
    $linkGitHub.Text = "GitHub"
    $linkGitHub.Location = New-Object System.Drawing.Point(10, 70)
    $linkGitHub.AutoSize = $true
    $linkGitHub.Add_LinkClicked({
        [System.Diagnostics.Process]::Start("https://github.com/GShwartz/Quick-AD/tree/main")
    })

    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $buttonOK.Location = New-Object System.Drawing.Point(100, 120)

    # Add controls to the form
    $aboutForm.Controls.Add($labelVersion)
    $aboutForm.Controls.Add($labelScripter)
    $aboutForm.Controls.Add($linkGitHub)
    $aboutForm.Controls.Add($buttonOK)

    # Show the form
    $aboutForm.ShowDialog()
}

# Function to show the top menu
function CreateMenuStrip {
    # Create a top menu strip
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    $menuStrip.Height = 8
    $menuStrip.BackColor = [System.Drawing.Color]::LightGray

    # File menu
    $fileMenu = $menuStrip.Items.Add("File")
    $exitMenuItem = $fileMenu.DropDownItems.Add("Exit")
    $exitMenuItem.Add_Click({
        # Exit the application
        [Windows.Forms.Application]::Exit()
    })

    # Help menu
    $helpMenu = $menuStrip.Items.Add("Help")
    $helpMenuAbout = $helpMenu.DropDownItems.Add("About")
    $helpMenuAbout.Add_Click({
        ShowAboutWindow
    })

    return $menuStrip
}