function SetPasswordStrength {
    # Define the form
    $SetPasswordForm = CreateCanvas "Password Settings" -x 400 -y 350
    $SetPasswordForm.MaximizeBox = $false

    # Create labels
    $labelUpper = CreateLabel -text "Min Uppercase:" -x 20 -y 20 -width 120 -height 20
    $labelLower = CreateLabel -text "Min Lowercase:" -x 20 -y 60 -width 120 -height 20
    $labelSpecial = CreateLabel -text "Min Special Characters:" -x 20 -y 100 -width 160 -height 20
    $labelDigits = CreateLabel -text "Number of Digits:" -x 20 -y 140 -width 120 -height 20

    # Create dropdown lists using the CreateNumbersDropdown function
    $dropUpper = CreateNumbersDropdown -x 250 -y 20 -width 40 -height 20 -minValue 1 -maxValue 3
    $dropLower = CreateNumbersDropdown -x 250 -y 60 -width 40 -height 20 -minValue 1 -maxValue 3
    $dropSpecial = CreateNumbersDropdown -x 250 -y 100 -width 40 -height 20 -minValue 1 -maxValue 3
    $dropDigits = CreateNumbersDropdown -x 250 -y 140 -width 40 -height 20 -minValue 6 -maxValue 10

    # Set the default selected index for each dropdown
    $dropUpper.SelectedIndex = 0
    $dropLower.SelectedIndex = 0
    $dropSpecial.SelectedIndex = 0
    $dropDigits.SelectedIndex = 0

    # Create checkbox for using initials
    $checkBoxUseInitials = CreateCheckbox -x 20 -y 180 -text "Include Initials"
    if ($global:passDefaultIncludeInitials) {$checkBoxUseInitials.Checked = $true} else {$checkBoxUseInitials.Checked = $false}

    # Create checkbox for using shuffle
    $checkBoxUseShuffle = CreateCheckbox -x 20 -y 220 -text "Shuffle"
    if ($global:passShuffle) {$checkBoxUseShuffle.Checked = $true} else {$checkBoxUseShuffle.Checked = $false}

    # Create a button to generate the password
    $buttonSavePasswordSettings = CreateButton -text "Save" -x 20 -y 260 -width 80 -height 25 -enabled $true
    $buttonSavePasswordSettings.Add_Click({
        $global:passDefaultUpperNum = [int]$dropUpper.SelectedItem
        $global:passDefaultLowerNum = [int]$dropLower.SelectedItem
        $global:passDefaultSpecialsNum = [int]$dropSpecial.SelectedItem
        $global:passDefaultDigitsNum = [int]$dropDigits.SelectedItem
        $global:passDefaultIncludeInitials = [bool]$checkBoxUseInitials.Checked
        $global:passShuffle = [bool]$checkBoxUseShuffle.Checked

        # Log the start of script execution
        LogScriptExecution -logPath $global:logFilePath -action "Password settings Saved." -userName $env:USERNAME

        $SetPasswordForm.Close()
    })

    # Create a 'Cancel' button
    $buttonCancelPassBuilder = CreateButton -text "Cancel" -x 280 -y 260 -width 70 -height 25 -enabled $true
    $buttonCancelPassBuilder.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $global:logFilePath -action "Password settings canceled" -userName $env:USERNAME

        $SetPasswordForm.Close()
    })

    # Add controls to the form
    $SetPasswordForm.Controls.Add($labelUpper)
    $SetPasswordForm.Controls.Add($dropUpper)
    $SetPasswordForm.Controls.Add($labelLower)
    $SetPasswordForm.Controls.Add($dropLower)
    $SetPasswordForm.Controls.Add($labelSpecial)
    $SetPasswordForm.Controls.Add($dropSpecial)
    $SetPasswordForm.Controls.Add($labelDigits)
    $SetPasswordForm.Controls.Add($dropDigits)
    $SetPasswordForm.Controls.Add($labelUseInitials)
    $SetPasswordForm.Controls.Add($checkBoxUseInitials)
    $SetPasswordForm.Controls.Add($labelShufflePassword)
    $SetPasswordForm.Controls.Add($checkBoxUseShuffle)
    $SetPasswordForm.Controls.Add($buttonSavePasswordSettings)
    $SetPasswordForm.Controls.Add($buttonCancelPassBuilder)

    # Display the form
    $SetPasswordForm.ShowDialog()
}
