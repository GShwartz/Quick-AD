# Import and load local modules
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

    # Create the re-enable form
    $reEnableForm = CreateCanvas "Re-Enable" -x 250 -y 150
    $reEnableForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $labelUsername = CreateLabel -text "Example User Name" -x 10 -y 10 -width 120 -height 20
    $textboxUsername = CreateTextbox -x 10 -y 30 -width 120 -height 10 -readOnly $false
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
                try {$userCheckup = FindADUser $exampleADuser} 
                finally {ReleaseLock}
            }
        
            if ($null -ne $userCheckup) {
                if ($userCheckup.Enabled -eq $true) {
                    # User is enabled, check if it's locked out
                    if (IsUserLockedOut $userCheckup.SamAccountName) {
                        # Update statusbar message
                        UpdateStatusBar "User '$exampleADuser' is locked-out." -color 'Black'

                        # User is locked out
                        $unlockAccountResult = [System.Windows.Forms.MessageBox]::Show("User '$exampleADuser' is locked. Do you want to unlock the account?", "User Locked", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
                        if ($unlockAccountResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                            # Unlock the user account
                            Unlock-ADAccount -Identity $exampleADuser
                            
                            # Update statusbar message
                            UpdateStatusBar "User '$exampleADuser' has been unlocked." -color 'Black'

                            # Display Summery dialog box
                            [System.Windows.Forms.MessageBox]::Show("User '$exampleADuser' has been unlocked.", "Account Unlocked", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                            # Enable the Re-Enable button
                            $buttonReEnableForm.Enabled = $true

                            # Log the start of script execution
                            LogScriptExecution -logPath $global:logFilePath -action "Unlocked user: '$($userCheckup.SamAccountName)'" -userName $env:USERNAME

                        } 
                        elseif ($unlockAccountResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                            # Enable the Re-Enable & Cancel buttons
                            $buttonReEnableForm.Enabled = $true
                            $buttonCancelReEnable.Enabled = $true

                            # Update statusbar message
                            UpdateStatusBar "Unlocked skipped on user: '$exampleADuser'." -color 'Black'

                            # Log the start of script execution
                            LogScriptExecution -logPath $global:logFilePath -action "Cancel Unlock for '$($userCheckup.SamAccountName)'." -userName $env:USERNAME
                        }
                    }
                    
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "Re-Enabling " -NoNewline -ForegroundColor Cyan
                    Write-Host "'$($global:primaryUser.SamAccountName)' " -NoNewline
                    Write-Host "with " -NoNewline -ForegroundColor Cyan
                    Write-Host "'$($exampleADuser)'" -NoNewline
                    Write-Host "..." -ForegroundColor Cyan

                    # Perform the re-enable action
                    Enable-AdAccount -Identity $global:primaryUser.SamAccountName

                    # Copy groups from example user
                    CopyGroups -adUsername $global:primaryUser.SamAccountName -exampleADuser $exampleADuser
                    
                    # Move the AD user to Example user OU
                    $isMoved = MoveUserToOU -exampleDisName $userCheckup.DistinguishedName -primaryDisName $global:primaryUser.DistinguishedName
                    if ($isMoved -eq $true) {
                        # Manage visuals
                        HideMark $global:form "ADUsername"

                        # Log action
                        LogScriptExecution -logPath $global:logFilePath -action "Re-Enabled user: '$($global:primaryUser.SamAccountName)'." -userName $env:USERNAME
                        
                        # Update statusbar message
                        UpdateStatusBar "User '$($global:primaryUser.SamAccountName)' has been re-enabled." -color 'Black'

                        # Display results
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "($dateTime) | " -NoNewline
                        Write-Host "User " -NoNewline -ForegroundColor Green
                        Write-Host "'$($global:primaryUser.SamAccountName)' " -NoNewline 
                        Write-Host "Re-Enabled with " -NoNewline -ForegroundColor Green
                        Write-Host "'$($exampleADuser)'."
                        [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' has been re-enabled.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                        
                        # Manage buttons
                        $buttonFindADUser.Enabled = $true
                        $buttonResetPassword.Enabled = $false
                        $buttonReEnableForm.Enabled = $false
                        $buttonGeneratePassword.Enabled = $false
                        $buttonCopyGroups.Enabled = $false
                        $buttonRemoveGroups.Enabled = $false
                        $buttonMoveOU.Enabled = $false

                        # Close the Re-Enable form
                        $reEnableForm.Close()
                        return $true
                    }
                    else {
                        return $isMoved
                    }
                } 
                else {
                    $buttonReEnableForm.Enabled = $false
                    $buttonCancelReEnable.Enabled = $true
                    HideMark $global:form "ADUsername"

                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "The user '$($userCheckup.SamAccountName)' is disabled." -userName $env:USERNAME

                    # Update statusbar message
                    UpdateStatusBar "The user '$($userCheckup.SamAccountName)' is disabled." -color 'Red'

                    # Display results
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "The user " -NoNewline -ForegroundColor Red
                    Write-Host "'$($exampleADuser)' " -NoNewline
                    Write-Host "is disabled." -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("User '$($userCheckup.SamAccountName)' is disabled.", "User Disabled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    
                    # Manage buttons
                    $global:buttonFindADUser.Enabled = $false
                    $global:buttonResetPassword.Enabled = $false
                    $global:buttonReEnableForm.Enabled = $false
                    $global:buttonGeneratePassword.Enabled = $false
                    $global:buttonCopyGroups.Enabled = $false
                    $global:buttonRemoveGroups.Enabled = $false
                    $global:buttonMoveOU.Enabled = $false

                    return $false
                }
            } 
            else {
                # Update statusbar message
                UpdateStatusBar "The user '$($exampleADuser) was not found." -color 'Red'

                # Display results
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "The user " -NoNewline -ForegroundColor Red
                Write-Host "'$($exampleADuser)' " -NoNewline
                Write-Host "was not found." -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show("The user '$($exampleADuser) was not found.", "ShowReEnableForm: User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                
                # Manage buttons
                $global:buttonFindADUser.Enabled = $false
                $buttonReEnableForm.Enabled = $false
                $buttonCancelReEnable.Enabled = $true

                return $false
            }
        } 
        catch {
            # Update statusbar message
            UpdateStatusBar "Error: $($_.Exception.Message)." -color 'Red'

            # Display results
            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            Write-Host "$($dateTime) | " -NoNewline
            Write-Host "Error: " -ForegroundColor Red
            Write-Host "$($_.Exception.Message)" -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            
            # Manage buttons
            $global:buttonReEnableForm.Enabled = $false
            $reEnableForm.Close()
            
            return $false
        }
    })

    # Add TextChanged event handler to the textbox
    $textboxUsername.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxUsername.Text)) {$buttonReEnableForm.Enabled = $false} else {$buttonReEnableForm.Enabled = $true}
    })

    # Create a 'Cancel' button
    $buttonCancelReEnable = CreateButton -text "Cancel" -x 140 -y 60 -width 70 -height 25 -enabled $true
    $buttonCancelReEnable.Add_Click({
        # Enable the main ReEnable button
        $global:buttonReEnable.Enabled = $true
        
        # Update statusbar message
        UpdateStatusBar "Re-Enable canceled." -color 'Black'

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

    # Show the form
    $reEnableForm.ShowDialog()
}