# Import local modules
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

$csvHandlerModulePath = Join-Path $scriptDirectory "CsvHandler.psm1"
$loggerModulePath = Join-Path $scriptDirectory "Logger.psm1"

Import-Module $csvHandlerModulePath -Force
Import-Module $loggerModulePath -Force

# Get the current username
$currentUserName = $env:USERNAME

# Function to display the Copy Groups form
function ShowCopyGroupsForm {
    param (
        [string]$adUsername,
        [string]$logFilePath,
        [System.Windows.Forms.Button]$buttonFindADUser,
        [System.Windows.Forms.Button]$buttonResetPassword,
        [System.Windows.Forms.Button]$buttonReEnable,
        [System.Windows.Forms.Button]$buttonGeneratePassword,
        [System.Windows.Forms.Button]$buttonCopyGroups,
        [System.Windows.Forms.Button]$buttonMoveOU
    )

    # Create a new form for copying groups
    $copyGroupsForm = New-Object System.Windows.Forms.Form
    $copyGroupsForm.Text = "Copy Groups"
    $copyGroupsForm.Size = New-Object System.Drawing.Size(280, 140)
    $copyGroupsForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    # Create a label for entering the AD username
    $labelUsername = New-Object System.Windows.Forms.Label
    $labelUsername.Location = New-Object System.Drawing.Point(10, 10)
    $labelUsername.Size = New-Object System.Drawing.Size(200, 20)
    $labelUsername.Text = "Example AD username"
    $copyGroupsForm.Controls.Add($labelUsername)

    # Create a textbox for AD username input
    $textboxUsername = New-Object System.Windows.Forms.TextBox
    $textboxUsername.Location = New-Object System.Drawing.Point(10, 30)
    $textboxUsername.Size = New-Object System.Drawing.Size(200, 20)
    $textboxUsername.Text = ""
    $copyGroupsForm.Controls.Add($textboxUsername)

    # Create a 'Copy Groups' button
    $buttonCopy = New-Object System.Windows.Forms.Button
    $buttonCopy.Location = New-Object System.Drawing.Point(10, 60)
    $buttonCopy.Size = New-Object System.Drawing.Size(90, 25)
    $buttonCopy.Text = "Copy"
    $buttonCopy.Enabled = $false
    $buttonCopy.Add_Click({
        # Disable the Copy Groups & Cancel buttons while working
        $buttonCopy.Enabled = $false
        $buttonCancelCopy.Enabled = $false

        try {
            # Get the example AD user from the textbox
            $exampleADuser = $textboxUsername.Text
            if ($exampleADuser -eq $global:primaryUser.SamAccountName) {
                # Log action
                LogScriptExecution -logPath $logFilePath -action "The user '$($exampleADuser) is the same as the primary." -userName $currentUserName

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
                CopyGroups -adUsername $global:primaryUser.SamAccountName -exampleADuser $exampleADuser
                
                # Enable buttons
                $buttonFindADUser.Enabled = $false
                $buttonGeneratePassword.Enabled = $true
                $buttonCopyGroups.Enabled = $true
                
                # Log the start of script execution
                LogScriptExecution -logPath $logFilePath -action "Groups have been copied from '$exampleADuser' to '$($global:primaryUser.SamAccountName)'." -userName $currentUserName

                # Show Summary dialog box
                [System.Windows.Forms.MessageBox]::Show("Groups have been copied to '$($global:primaryUser.SamAccountName)'.", "Copy Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
                # Close the Copy Groups form
                $copyGroupsForm.Close()

            } else {
                # Log action
                LogScriptExecution -logPath $logFilePath -action "The user '$($exampleADuser) was not found." -userName $currentUserName

                # Display Not-Found dialog
                [System.Windows.Forms.MessageBox]::Show("The user '$($exampleADuser) was not found.", "Copy Groups: User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                # Disable the Copy Groups button and enable the Cancel button
                $buttonCopy.Enabled = $false
                $buttonCancelCopy.Enabled = $true
            }

        } catch {
            # Log action
            LogScriptExecution -logPath $logFilePath -action "Error: $($_.Exception.Message)" -userName $currentUserName

            # Disable the Copy Groups button and close the form
            $buttonCopy.Enabled = $false
            $copyGroupsForm.Close()
            
            Write-Error "$($_)"
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

    # Add the Copy Groups button to the form
    $copyGroupsForm.Controls.Add($buttonCopy)

    # Create a 'Cancel' button
    $buttonCancelCopy = New-Object System.Windows.Forms.Button
    $buttonCancelCopy.Location = New-Object System.Drawing.Point(140, 60)
    $buttonCancelCopy.Size = New-Object System.Drawing.Size(70, 25)
    $buttonCancelCopy.Text = "Cancel"
    $buttonCancelCopy.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Copy Groups canceled." -userName $currentUserName

        $copyGroupsForm.Close()
        $buttonFindADUser.Focus()
    })

    # Add the Cancel button
    $copyGroupsForm.Controls.Add($buttonCancelCopy)

    # Show the Copy Groups form
    $copyGroupsForm.ShowDialog()
}