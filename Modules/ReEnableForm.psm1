# Import local modules
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

$moduleNames = @("BuildForm", "CsvHandler", "Logger", "Visuals")

foreach ($moduleName in $moduleNames) {
    $modulePath = Join-Path $scriptDirectory "$moduleName.psm1"
    Import-Module $modulePath -Force
}

# Function to display the Re-Enable form
function ShowReEnableForm {
    # Disable the main ReEnable button
    $global:buttonReEnable.Enabled = $false

    # Create a new form for re-enabling a user
    $reEnableForm = CreateCanvas "Re-Enable" -x 250 -y 150
    $reEnableForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create a label for entering the AD username
    $labelUsername = CreateLabel -text "Example User Name" -x 10 -y 10 -width 120 -height 20

    # Create a textbox for AD username input
    $textboxUsername = CreateTextbox -x 10 -y 30 -width 120 -height 10 -readOnly $false

    # Create a 'Re-Enable' button
    $buttonReEnableForm = CreateButton -text "Re-Enable" -x 10 -y 60 -width 70 -height 25 -enabled $false
    $buttonReEnableForm.Add_Click({
        $global:buttonFindADUser.Enabled = $false

        try {
            # Get the example AD user from the textbox
            $exampleADuser = $textboxUsername.Text
            
            # Disable the Re-Enable & Cancel buttons while working
            $buttonReEnableForm.Enabled = $false
            $buttonCancelReEnable.Enabled = $false

            # Acquire the lock
            if (AcquireLock) {
                try {
                    # Send AD request to find and check the example user
                    $userCheckup = FindADUser $exampleADuser
                
                } finally {
                    # Release the lock when done (even if an error occurs)
                    ReleaseLock
                }
            }
        
            if ($null -ne $userCheckup) {
                if ($userCheckup.Enabled -eq $true) {
                    # User is enabled, check if it's locked out
                    if (IsUserLockedOut $userCheckup.SamAccountName) {
                        # Update statusbar message
                        UpdateStatusBar "User '$exampleADuser' is locked-out." -color 'DarkOrange'

                        # User is locked out
                        $unlockAccountResult = [System.Windows.Forms.MessageBox]::Show("User '$exampleADuser' is locked. Do you want to unlock the account?", "User Locked", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
                        if ($unlockAccountResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                            # Unlock the user account
                            Unlock-ADAccount -Identity $exampleADuser
                            
                            # Update statusbar message
                            UpdateStatusBar "User '$exampleADuser' has been unlocked." -color 'White'

                            # Display Summery dialog box
                            [System.Windows.Forms.MessageBox]::Show("User '$exampleADuser' has been unlocked.", "Account Unlocked", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                            # Enable the Re-Enable button
                            $buttonReEnableForm.Enabled = $true

                            # Log the start of script execution
                            LogScriptExecution -logPath $global:logFilePath -action "Unlocked user: '$($userCheckup.SamAccountName)'" -userName $env:USERNAME

                        } elseif ($unlockAccountResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                            # Enable the Re-Enable & Cancel buttons
                            $buttonReEnableForm.Enabled = $true
                            $buttonCancelReEnable.Enabled = $true

                            # Update statusbar message
                            UpdateStatusBar "Unlocked skipped on user: '$exampleADuser'." -color 'DarkOrange'

                            # Log the start of script execution
                            LogScriptExecution -logPath $global:logFilePath -action "Cancel Unlock for '$($userCheckup.SamAccountName)'." -userName $env:USERNAME
                        }
                    }
                    
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host ""
                    Write-Host "======================================================"
                    Write-Host "$($dateTime) | Re-Enabling '$($global:primaryUser.SamAccountName)' with '$($exampleADuser)'..."
                    Write-Host "======================================================"

                    # Perform the re-enable action
                    Enable-AdAccount -Identity $global:primaryUser.SamAccountName

                    # Copy groups from example user
                    CopyGroups -adUsername $global:primaryUser.SamAccountName -exampleADuser $exampleADuser
                    
                    # Move the AD user to Example user OU
                    $isMoved = MoveUserToOU -exampleDisName $userCheckup.DistinguishedName -primaryDisName $global:primaryUser.DistinguishedName
                    if ($isMoved -eq $true) {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        HideMark $form "ADUsername"

                        # Show summery in console
                        Write-Host "======================================================"
                        Write-Host "($dateTime) | User '$($global:primaryUser.SamAccountName)' Re-Enabled with '$($exampleADuser)'."
                        Write-Host "======================================================"

                        # Log action
                        LogScriptExecution -logPath $global:logFilePath -action "Re-Enabled user: '$($global:primaryUser.SamAccountName)'." -userName $env:USERNAME

                        # Manage buttons
                        $buttonFindADUser.Enabled = $true
                        $buttonResetPassword.Enabled = $false
                        $buttonReEnableForm.Enabled = $false
                        $buttonGeneratePassword.Enabled = $false
                        $buttonCopyGroups.Enabled = $false
                        $buttonRemoveGroups.Enabled = $false
                        $buttonMoveOU.Enabled = $false
                        
                        HideMark $global:form "ADUsername"

                        # Update statusbar message
                        UpdateStatusBar "User '$($global:primaryUser.SamAccountName)' has been re-enabled." -color 'White'

                        # Show Summery dialog box
                        [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' has been re-enabled.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        
                        # Close the Re-Enable form
                        $reEnableForm.Close()

                        return $true
                    }
                    else {
                        return $isMoved
                    }

                } else {
                    $buttonReEnableForm.Enabled = $false
                    $buttonCancelReEnable.Enabled = $true
                    HideMark $global:form "ADUsername"

                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "The user '$($userCheckup.SamAccountName)' is disabled." -userName $env:USERNAME

                    # Update statusbar message
                    UpdateStatusBar "The user '$($userCheckup.SamAccountName)' is disabled." -color 'Red'

                    # User is disabled
                    [System.Windows.Forms.MessageBox]::Show("User '$($userCheckup.SamAccountName)' is disabled.", "User Disabled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    
                    # Manage buttons
                    $buttonFindADUser.Enabled = $false
                    $buttonResetPassword.Enabled = $false
                    $buttonReEnableForm.Enabled = $false
                    $buttonGeneratePassword.Enabled = $false
                    $buttonCopyGroups.Enabled = $false
                }

            } else {
                # Update statusbar message
                UpdateStatusBar "The user '$($exampleADuser) was not found." -color 'Red'

                $global:buttonFindADUser.Enabled = $false
                [System.Windows.Forms.MessageBox]::Show("The user '$($exampleADuser) was not found.", "ShowReEnableForm: User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                
                # Disable the Re-Enable button and enable the cancel button
                $buttonReEnableForm.Enabled = $false
                $buttonCancelReEnable.Enabled = $true

                return 
            }

        } catch {
            # Update statusbar message
            UpdateStatusBar "Error: $($_.Exception.Message)." -color 'Red'

            # Handle any other errors that may occur
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            
            # Disable the Re-Enable button and close the re-enable form
            $global:buttonReEnableForm.Enabled = $false
            $reEnableForm.Close()
            
            return
        }
    })

    # Add TextChanged event handler to the textbox
    $textboxUsername.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxUsername.Text)) {
            $buttonReEnableForm.Enabled = $false

        } else {
            $buttonReEnableForm.Enabled = $true
        }
    })

    # Create a 'Cancel' button
    $buttonCancelReEnable = CreateButton -text "Cancel" -x 140 -y 60 -width 70 -height 25 -enabled $true
    $buttonCancelReEnable.Add_Click({
        # Enable the main ReEnable button
        $global:buttonReEnable.Enabled = $true
        
        $reEnableForm.Close()
        $global:buttonFindADUser.Focus()

        # Log the start of script execution
        LogScriptExecution -logPath $global:logFilePath -action "Cancel Re-Enable" -userName $env:USERNAME
    })

    # Add the form controls
    $reEnableForm.Controls.Add($labelUsername)
    $reEnableForm.Controls.Add($textboxUsername)
    $reEnableForm.Controls.Add($buttonReEnableForm)
    $reEnableForm.Controls.Add($buttonCancelReEnable)

    # Show the re-enable form
    $reEnableForm.ShowDialog()
}