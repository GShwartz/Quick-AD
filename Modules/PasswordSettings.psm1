# Function that sets the password strength
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
    if ($global:passDefaultUpperNum -ge 1 -and $global:passDefaultUpperNum -le 3) {
        $dropUpper.SelectedIndex = $global:passDefaultUpperNum - 1
    } 
    else {
        Write-Host "Maximum 3 Upper-Case letters. Check default settings." -ForegroundColor Red
        return $false
    }

    if ($global:passDefaultLowerNum -ge 1 -and $global:passDefaultLowerNum -le 3) {
        $dropLower.SelectedIndex = $global:passDefaultLowerNum - 1
    }
    else {
        Write-Host "Maximum 3 Lower-Case letters. Check default settings." -ForegroundColor Red
        return $false
    }

    if ($global:passDefaultSpecialsNum -ge 1 -and $global:passDefaultSpecialsNum -le 3) {
        $dropSpecial.SelectedIndex = $global:passDefaultSpecialsNum - 1
    }
    else {
        Write-Host "Minimum 1 & Maximum 3 Special characters. Check default settings." -ForegroundColor Red
        return $false
    }

    if ($global:passDefaultDigitsNum -ge 6 -and $global:passDefaultDigitsNum -le 10) {
        if ($global:passDefaultDigitsNum -eq 6) {$dropDigits.SelectedIndex = 0}
        elseif ($global:passDefaultDigitsNum -eq 7) {$dropDigits.SelectedIndex = 1}
        elseif ($global:passDefaultDigitsNum -eq 8) {$dropDigits.SelectedIndex = 2}
        elseif ($global:passDefaultDigitsNum -eq 9) {$dropDigits.SelectedIndex = 3}
        elseif ($global:passDefaultDigitsNum -eq 10) {$dropDigits.SelectedIndex = 4}
    }
    else {
        Write-Host "Minimum 6 & Maximum 10 Digits. Check default settings." -ForegroundColor Red
        return $false
    }

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

    if ($global:showTooltips -eq $true) {
        # Create tooltips for controls in SetPasswordStrength function
        CreateToolTip -control $labelUpper -text "Minimum number of uppercase letters."
        CreateToolTip -control $labelLower -text "Minimum number of lowercase letters."
        CreateToolTip -control $labelSpecial -text "Minimum number of special characters."
        CreateToolTip -control $labelDigits -text "Number of digits in the password."
        CreateToolTip -control $dropUpper -text "Select the minimum number of uppercase letters."
        CreateToolTip -control $dropLower -text "Select the minimum number of lowercase letters."
        CreateToolTip -control $dropSpecial -text "Select the minimum number of special characters."
        CreateToolTip -control $dropDigits -text "Select the number of digits in the password."
        CreateToolTip -control $checkBoxUseInitials -text "Use AD user's initials in the password."
        CreateToolTip -control $checkBoxUseShuffle -text "Shuffle the characters in the password."
        CreateToolTip -control $buttonSavePasswordSettings -text "Save the password settings."
        CreateToolTip -control $buttonCancelPassBuilder -text "Cancel and close the form."
    }

    # Display the form
    $SetPasswordForm.ShowDialog()
}
