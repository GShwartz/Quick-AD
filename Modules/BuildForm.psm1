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
        [int]$y
    )

    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Location = New-Object System.Drawing.Point($x, $y)

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

    return $dropdown
}
