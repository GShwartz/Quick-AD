# Import local modules
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

$csvHandlerModulePath = Join-Path $scriptDirectory "CsvHandler.psm1"
$loggerModulePath = Join-Path $scriptDirectory "Logger.psm1"
$buildFormModulePath = Join-Path $scriptDirectory "BuildForm.psm1"

Import-Module $csvHandlerModulePath -Force
Import-Module $loggerModulePath -Force
Import-Module $buildFormModulePath -Force

# Function to display the Copy Groups form
function ShowCopyGroupsForm {
    # Create the Copy Groups form components
    $copyGroupsForm = CreateCanvas "Copy Groups" -x 280 -y 140
    $copyGroupsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $labelUsername = CreateLabel -text "Example AD Username" -x 10 -y 10 -width 200 -height 20
    $textboxUsername = CreateTextbox -x 10 -y 30 -width 200 -height 20 -readOnly $false
    $buttonCopy = CreateButton -text "Copy" -x 10 -y 60 -width 90 -height 25 -enabled $false
    
    # Functionality for button click
    $buttonCopy.Add_Click({
        # Disable the Copy Groups & Cancel buttons while working
        $buttonCopy.Enabled = $false
        $buttonCancelCopy.Enabled = $false

        try {
            # Get the example AD user from the textbox
            $exampleADuser = $textboxUsername.Text
            if ($exampleADuser -eq $global:primaryUser.SamAccountName) {
                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "The user '$($exampleADuser)' is the same as the primary." -userName $env:USERNAME

                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "The user " -NoNewline -ForegroundColor Red
                Write-Host "'$($exampleADuser)' " -NoNewline
                Write-Host "is the same as the primary." -ForegroundColor Red

                # Update statusbar message
                UpdateStatusBar "The user '$($exampleADuser)' is the same as the primary." -color 'Red'

                # Display Not-Found dialog
                [System.Windows.Forms.MessageBox]::Show("The user '$($exampleADuser) is the same as the primary.", "Duplicated Entry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                # Disable the Copy Groups button and enable the Cancel button
                $buttonCancelCopy.Enabled = $true

                return $false
            }

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
                # Perform the copy groups action
                $isCopied, $message = CopyGroups -adUsername $global:primaryUser.SamAccountName -exampleADuser $exampleADuser
                if (-not $isCopied) {
                    # Log the start of script execution
                    LogScriptExecution -logPath $global:logFilePath -action "Failed to copy groups from $($exampleADuser). $($message)" -userName $env:USERNAME

                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "Failed to copy groups from " -NoNewline -ForegroundColor Red
                    Write-Host "'$($exampleADuser)'. " -NoNewline
                    Write-Host "$($message)"

                    # Update statusbar message
                    UpdateStatusBar "Failed to copy groups from $($exampleADuser): $($message)" -color 'DarkRed'

                    # Show Summary dialog box
                    [System.Windows.Forms.MessageBox]::Show("Failed to copy groups from $($exampleADuser).", "Copy Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                    $buttonCancelCopy.Enabled = $true
                    return $false
                }
                
                # Enable buttons
                $global:buttonFindADUser.Enabled = $false
                $global:buttonGeneratePassword.Enabled = $true
                $global:buttonCopyGroups.Enabled = $true
                
                $message = "Groups have been copied from '$exampleADuser' to '$($global:primaryUser.SamAccountName)'."

                # Log the start of script execution
                LogScriptExecution -logPath $global:logFilePath -action "$($message)" -userName $env:USERNAME

                # Update statusbar message
                UpdateStatusBar "$($message)" -color 'Black'

                # Show Summary dialog box
                [System.Windows.Forms.MessageBox]::Show("$($message)", "Copy Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Close the Copy Groups form
                $copyGroupsForm.Close()

                return $true, $message
            } 
            else {
                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "The user '$($exampleADuser) was not found." -userName $env:USERNAME

                # Update statusbar message
                UpdateStatusBar "The user '$($exampleADuser) was not found." -color 'Red'

                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "The user " -NoNewline -ForegroundColor Red
                Write-Host "'$($exampleADuser)' " -NoNewline
                Write-Host "was not found." -ForegroundColor Red

                # Display Not-Found dialog
                [System.Windows.Forms.MessageBox]::Show("The user '$($exampleADuser) was not found.", "User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                # Disable the Copy Groups button and enable the Cancel button
                $buttonCopy.Enabled = $false
                $buttonCancelCopy.Enabled = $true
                return $false
            }
        } 
        catch {
            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Error: $($_.Exception.Message)" -userName $env:USERNAME

            # Update statusbar message
            UpdateStatusBar "Error: $($_.Exception.Message)" -color 'Red'

            # Disable the Copy Groups button and close the form
            $buttonCopy.Enabled = $false
            $copyGroupsForm.Close()
            
            Write-Error "$($_)"
            return $false
        }
    })

    # Add TextChanged event handler to the textbox
    $textboxUsername.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxUsername.Text)) {
            $buttonCopy.Enabled = $false

        } else {
            $buttonCopy.Enabled = $true
        }
    })

    # Create a 'Cancel' button
    $buttonCancelCopy = CreateButton -text "Cancel" -x 140 -y 60 -width 70 -height 25 -enabled $true
    $buttonCancelCopy.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $global:logFilePath -action "Copy Groups canceled." -userName $env:USERNAME

        $copyGroupsForm.Close()
        $global:buttonFindADUser.Focus()
    })

    # Add the Cancel button
    $copyGroupsForm.Controls.Add($labelUsername)
    $copyGroupsForm.Controls.Add($textboxUsername)
    $copyGroupsForm.Controls.Add($buttonCopy)
    $copyGroupsForm.Controls.Add($buttonCancelCopy)

    # Show the Copy Groups form
    $copyGroupsForm.ShowDialog()
}