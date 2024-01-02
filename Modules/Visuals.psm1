# Function to draw a green V mark near the textbox
function DrawVmark ($Form, $x, $y, $name) {
    # Create a PictureBox for the checkmark sign
    $PictureBox = New-Object System.Windows.Forms.PictureBox
    $PictureBox.Name = "$($name)"
    $PictureBox.Location = New-Object System.Drawing.Point($x, $y)
    $PictureBox.Size = New-Object System.Drawing.Size(25, 25)
    $Form.Controls.Add($PictureBox)

    # Draw a green checkmark on the PictureBox with a bolder line
    $PictureBox.Add_Paint({
        $e = $_
        $g = $e.Graphics
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Green, 3)  # Adjust the line thickness as needed (3 in this case)
        $points = @(
            [System.Drawing.Point]::new(10, 15),
            [System.Drawing.Point]::new(12, 25),
            [System.Drawing.Point]::new(25, 5)
        )
        $g.DrawLines($pen, $points)
    })

    return
}

# Function to draw a red X mark near the textbox
function DrawXmark ($Form, $x, $y, $name) {
    # Create a PictureBox for the X mark
    $PictureBox = New-Object System.Windows.Forms.PictureBox
    $PictureBox.Name = "$($name)"
    $PictureBox.Location = New-Object System.Drawing.Point($x, $y)
    $PictureBox.Size = New-Object System.Drawing.Size(23, 23)
    $form.Controls.Add($PictureBox)

    # Draw a red X on the PictureBox with a bolder line
    $PictureBox.Add_Paint({
        $e = $_
        $g = $e.Graphics
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Red, 3)
        $g.DrawLine($pen, 5, 5, 20, 20)
        $g.DrawLine($pen, 20, 5, 5, 20)
    })

    return
}

# Function to hide the mark by X, Y, and name
function HideMark ($form, $name) {
    # Find the PictureBox control by its name and remove it from the form
    $pictureBox = $form.Controls | Where-Object { $_.Name -eq $name }
    if ($null -ne $pictureBox) {
        #$Form.Controls.Remove($pictureBox)
        $pictureBox.Dispose()  # Dispose of the PictureBox to release resources
    }

    return
}

# Function to hide all marks
function HideAllMarks($form, $prefixes) {
    # Iterate through the controls of the form
    foreach ($control in $form.Controls) {
        if ($prefixes -contains ($prefixes | Where-Object { $control.Name -like "$_*" })) {
            # Found a PictureBox with a matching prefix, dispose of it to hide
            $control.Dispose()
        }
    }
}

# Function to create the tooltips over the buttons
function CreateToolTip {
    param (
        [System.Windows.Forms.Control]$control,
        [string]$text
    )

    # Define the tooltip
    $tooltip = New-Object System.Windows.Forms.ToolTip
    $tooltip.SetToolTip($control, $text)

    # Define the timer
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 2000  # 2000 milliseconds (2 seconds)

    # Add mouse enter event handler to the buttons
    $control.Add_MouseEnter({
        if ($control.Text -match "Find User") {
           $findUserTimer = New-Object System.Windows.Forms.Timer
           $findUserTimer.Interval = 2000
           $findUserTimer.Start()
        }
        elseif ($control.Text -match "Find Computer") {
            $findComputerTimer = New-Object System.Windows.Forms.Timer
            $findComputerTimer.Interval = 2000
            $findComputerTimer.Start()
        }
        elseif ($control.Text -match "Browse") {
            $browseTimer = New-Object System.Windows.Forms.Timer
            $browseTimer.Interval = 2000
            $browseTimer.Start()
        }
        elseif ($control.Text -match "Generate Password") {
            $generatePasswordTimer = New-Object System.Windows.Forms.Timer
            $generatePasswordTimer.Interval = 2000
            $generatePasswordTimer.Start()
        }
        elseif ($control.Text -match "Reset Password") {
            $resetPasswordTimer = New-Object System.Windows.Forms.Timer
            $resetPasswordTimer.Interval = 2000
            $resetPasswordTimer.Start()
        }
        elseif ($control.Text -match "Re-Enable") {
            $reEnableTimer = New-Object System.Windows.Forms.Timer
            $reEnableTimer.Interval = 2000
            $reEnableTimer.Start()
        }
        elseif ($control.Text -match "Copy Groups") {
            $copyGroupsTimer = New-Object System.Windows.Forms.Timer
            $copyGroupsTimer.Interval = 2000
            $copyGroupsTimer.Start()
        }
        elseif ($control.Text -match "Remove Groups") {
            $removeGroupsTimer = New-Object System.Windows.Forms.Timer
            $removeGroupsTimer.Interval = 2000
            $removeGroupsTimer.Start()
        }
        elseif ($control.Text -match "Move OU") {
            $moveOUTimer = New-Object System.Windows.Forms.Timer
            $moveOUTimer.Interval = 2000
            $moveOUTimer.Start()
        }
    })

    # Add mouse leave event handler to the buttons
    $control.Add_MouseLeave({
        if ($control.Text -match "Find User") {
            $findUserTimer = New-Object System.Windows.Forms.Timer
            $findUserTimer.Interval = 2000
            $findUserTimer.Stop()
            $tooltip.Hide($control)
         }
         elseif ($control.Text -match "Find Computer") {
            $findComputerTimer = New-Object System.Windows.Forms.Timer
            $findComputerTimer.Interval = 2000
            $findComputerTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Browse") {
            $browseTimer = New-Object System.Windows.Forms.Timer
            $browseTimer.Interval = 2000
            $browseTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Generate Password") {
            $generatePasswordTimer = New-Object System.Windows.Forms.Timer
            $generatePasswordTimer.Interval = 2000
            $generatePasswordTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Reset Password") {
            $resetPasswordTimer = New-Object System.Windows.Forms.Timer
            $resetPasswordTimer.Interval = 2000
            $resetPasswordTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Re-Enable") {
            $reEnableTimer = New-Object System.Windows.Forms.Timer
            $reEnableTimer.Interval = 2000
            $reEnableTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Copy Groups") {
            $copyGroupsTimer = New-Object System.Windows.Forms.Timer
            $copyGroupsTimer.Interval = 2000
            $copyGroupsTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Remove Groups") {
            $removeGroupsTimer = New-Object System.Windows.Forms.Timer
            $removeGroupsTimer.Interval = 2000
            $removeGroupsTimer.Stop()
            $tooltip.Hide($control)
        }
        elseif ($control.Text -match "Move OU") {
            $moveOUTimer = New-Object System.Windows.Forms.Timer
            $moveOUTimer.Interval = 2000
            $moveOUTimer.Stop()
            $tooltip.Hide($control)
        }
    })

    $timer.Add_Tick({
        $timer.Stop()
    })
}

# Function to update the status bar text
function UpdateStatusBar($message, $color) {
    $global:statusBarTextBox.Text = $message
    $global:statusBarTextBox.ForeColor = $color
}



