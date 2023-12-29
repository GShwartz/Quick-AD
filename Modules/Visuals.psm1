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
