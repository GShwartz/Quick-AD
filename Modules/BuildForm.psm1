# Function to create the Windows.Form
function CreateCanvas {
    param (
        [string]$formTitle,
        [int]$x,
        [int]$y
    )

    $form = New-Object Windows.Forms.Form
    $form.Text = $formTitle
    $form.Size = New-Object System.Drawing.Size($x, $y)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.Margin = '0,0,0,0'
    $form.Padding = '0,0,0,0'

    return $form
}

# Function to create a Windows.Forms labels
function CreateLabel {
    param(
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($x, $y)
    $label.Size = New-Object System.Drawing.Size($width, $height)
    $label.Text = $text

    return $label
}

# Function to create a Windows.Forms textboxes
function CreateTextbox {
    param(
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [bool]$readOnly
    )

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point($x, $y)
    $textbox.Size = New-Object System.Drawing.Size($width, $height)
    $textbox.ReadOnly = $readOnly

    return $textbox
}

# Function to create a Windows.Forms buttons
function CreateButton {
    param(
        [string]$text,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [bool]$enabled
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point($x, $y)
    $button.Size = New-Object System.Drawing.Size($width, $height)
    $button.Text = $text
    #$button.ForeColor = [System.Drawing.Color]::Black
    #$button.BackColor = [System.Drawing.Color]::Gray
    $button.Enabled = $enabled

    return $button
}

# Function to create a Windows.Forms List box
function CreateListBox {
    param (
        [string]$name,
        [int]$width,
        [int]$height,
        [int]$x,
        [int]$y,
        [string[]]$items
    )

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Name = $name
    $listBox.Size = New-Object System.Drawing.Size($width, $height)
    $listBox.Location = New-Object System.Drawing.Point($x, $y)
    $listBox.Items.AddRange($items)

    return $listBox
}

# Function to create a Windows.Forms checkboxes
function CreateCheckbox {
    param(
        [int]$x,
        [int]$y,
        [string]$text
    )

    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Location = New-Object System.Drawing.Point($x, $y)
    $checkbox.Text = $text

    return $checkbox
}

# Function to create a Windows.Forms dropdown
function CreateNumbersDropdown {
    param(
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height,
        [int]$minValue,
        [int]$maxValue
    )

    $dropdown = New-Object System.Windows.Forms.ComboBox
    $dropdown.Location = New-Object System.Drawing.Point($x, $y)
    $dropdown.Size = New-Object System.Drawing.Size($width, $height)

    # Populate the dropdown with numbers from $minValue to $maxValue
    $dropdown.Items.AddRange($minValue..$maxValue)

    # Set the dropdown style to DropDownList to make it read-only
    $dropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

    return $dropdown
}

# Function to create a status bar
function CreateStatusBar {
    # Create a status bar
    $global:statusBar = New-Object System.Windows.Forms.StatusBar
    $statusBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $statusBar.Padding = '0,0,0,0'

    # Create a textbox for status bar
    $global:statusBarTextBox = New-Object System.Windows.Forms.TextBox
    $statusBarTextBox.Multiline = $false
    $statusBarTextBox.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $statusBarTextBox.Text = "Ready"
    $statusBarTextBox.ForeColor = [System.Drawing.Color]::Black
    $statusBarTextBox.BackColor = [System.Drawing.Color]::Wheat
    $statusBarTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $statusBarTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Horizontal 
    $statusBarTextBox.ReadOnly = $true

    # Add an event handler for resizing the form
    $form.add_Resize({
        $statusBarTextBox.Width = $form.Width - 15
    })

    # Add the textbox to the form
    $form.Controls.Add($statusBarTextBox)
    $statusBar.Panels.Add((New-Object System.Windows.Forms.StatusBarPanel))
    $statusBar.Controls.Add($statusBarTextBox)

    # Allow marking the text to activate the horizontal scrollbar
    $statusBarTextBox.Add_MouseDown({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $statusBarTextBox.SelectAll()
        }
    })
}

# Function to create a listbox
function ShowListBoxForm {
    param (
        [string]$formTitle,
        [string]$listBoxName,
        [int]$formWidth,
        [int]$formHeight,
        [int]$listBoxWidth,
        [int]$listBoxHeight,
        [int]$cancelX,
        [int]$cancelY,
        [array]$items
    )

    $OUSelectionForm = CreateCanvas $formTitle -x $formWidth -y $formHeight
    $listBox = CreateListBox -name $listBoxName -width $listBoxWidth -height $listBoxHeight -x 20 -y 20 -items $items
    $buttonSelect = CreateButton -text "Select" -x 20 -y ($listBoxHeight + 30) -width 80 -height 25 -enabled $false
    $selectedItem = $null

    $buttonSelect.Add_Click({
        if ($null -ne $listBox.SelectedItem) {
            $script:selectedItem = $listBox.SelectedItem.ToString()
            $OUSelectionForm.Close()
        }
        else {
            # Update statusbar message
            UpdateStatusBar "No item selected." -color 'Yellow'

            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            Write-Host "$($dateTime) | " -NoNewline
            Write-Host "No item selected." -NoNewline -ForegroundColor Yellow
        }
    })

    # Add a Hover event for the ListBox - Enable 'Select' button only if an item is selected.
    $listBox.Add_MouseDown({
        try {
            $tempSelection = $listBox.SelectedItem.ToString()
        }
        catch {
            Write-Host "No item selected." -ForegroundColor Yellow
        }

        if (-not [string]::IsNullOrEmpty($tempSelection) -or -not [string]::IsNullOrWhiteSpace($tempSelection)) {
            $buttonSelect.Enabled = $true
        }
        else {
            $buttonSelect.Enabled = $false
        }
    })

    $buttonCancel = CreateButton -text "Cancel" -x $cancelX -y $cancelY -width 70 -height 25 -enabled $true
    $buttonCancel.Add_Click({
        $OUSelectionForm.Close()
    })

    # Add controls to the form
    $OUSelectionForm.Controls.Add($listBox)
    $OUSelectionForm.Controls.Add($buttonSelect)
    $OUSelectionForm.Controls.Add($buttonCancel)

    # Show the form
    $OUSelectionForm.ShowDialog()
    return $script:selectedItem
}






