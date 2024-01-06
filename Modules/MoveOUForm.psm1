# Import local modules
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import local modules
$modulePaths = @(
    "Logger",           # Handles logging
    "CsvHandler",       # Handles CSV file operations
    "ADHandler",        # Handles Active Directory operations
    "BuildForm"        # Builds and configures forms
)

foreach ($moduleName in $modulePaths) {
    $modulePath = Join-Path $scriptDirectory "$moduleName.psm1"
    Import-Module $modulePath -Force
}

# Function to handle the AD User OU move
function HandleUser {
    # Get the example AD user from the textbox
    $exampleADuser = $textboxExample.Text
    # Disable the Copy Groups & Cancel buttons while working
    $buttonMove.Enabled = $false
    $buttonCancelMoveOU.Enabled = $false

    if ($global:primaryUser.SamAccountName -eq $exampleADuser) {
        # Update statusbar message
        UpdateStatusBar "'$($exampleADuser) is the same as the primary." -color 'Red'

        # Display Error dialog
        [System.Windows.Forms.MessageBox]::Show("'$($exampleADuser) is the same as the primary.", "Duplicated Entry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        $buttonCancelMoveOU.Enabled = $true

        return $false
    }

    # Acquire the lock
    if (AcquireLock) {
        try {
            # Send AD request to find and check the example user
            $userCheckup = FindADUser $exampleADuser
            if ($null -ne $userCheckup) {
                # Relocate user to designated OU
                MoveUserToOU -exampleDisName $userCheckup.DistinguishedName -primaryDisName $global:primaryUser.DistinguishedName

                # Enable buttons
                $buttonFindADUser.Enabled = $true
                $buttonGeneratePassword.Enabled = $true
                $buttonMove.Enabled = $false
                $buttonMoveOU.Enabled = $false
                
                # Log the start of script execution
                LogScriptExecution -logPath $logFilePath -action "'$($global:textboxADUsername.Text)' has been relocated to $($userCheckup.DistinguishedName)." -userName $env:USERNAME

                # Update statusbar message
                UpdateStatusBar "'$($global:textboxADUsername.Text)' has been relocated to '$($userCheckup.DistinguishedName)'." -color 'Black'

                # Close the Copy Groups form
                $moveOUForm.Close()

                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "The user '" -NoNewline -ForegroundColor Green
                Write-Host "$($global:primaryUser.SamAccountName)' " -NoNewline 
                Write-Host "was reloacted to " -NoNewline -ForegroundColor Green
                Write-Host "$($userCheckup.DistinguishedName)"
                    
                # Show Summary dialog box
                [System.Windows.Forms.MessageBox]::Show("'$($global:textboxADUsername.Text)' relocated successfully.", "Move OU", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            else {
                # Log the start of script execution
                LogScriptExecution -logPath $logFilePath -action "The user '$($exampleADuser) was not found." -userName $env:USERNAME
                
                # Update statusbar message
                UpdateStatusBar "The user '$($exampleADuser) was not found." -color 'Red'

                # Display Not-Found dialog
                [System.Windows.Forms.MessageBox]::Show("The user '$($exampleADuser) was not found.", "MoveOU: User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
                # Disable the Copy Groups button and enable the Cancel button
                $buttonMove.Enabled = $false
                $buttonCancelMoveOU.Enabled = $true
            }

        } finally {
            # Release the lock when done (even if an error occurs)
            ReleaseLock
        }
    }
} 

# Function to handle the AD Computer OU move
function HandleComputer {
    # Get the example AD computer from the textbox
    $exampleADComputer = $textboxExample.Text

    # Disable the Copy Groups & Cancel buttons while working
    $buttonMove.Enabled = $false
    $buttonCancelMoveOU.Enabled = $false

    if ($global:textboxADComputer.Text -eq $exampleADComputer) {
        # Update statusbar message
        UpdateStatusBar "'$($exampleADComputer) is the same as the primary." -color 'Red'

        # Display Not-Found dialog
        [System.Windows.Forms.MessageBox]::Show("'$($exampleADComputer) is the same as the primary.", "Duplicated Entry", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        $buttonCancelMoveOU.Enabled = $true

        return $false
    }

    # Acquire the lock
    if (AcquireLock) {
        try {
            # Send AD request to find and check the example user
            $computer = FindADComputer $exampleADComputer
            if ($computer) {
                # Relocate user to designated OU
                MoveUserToOU -exampleDisName $computer.DistinguishedName -primaryDisName $global:primaryComputer.DistinguishedName
                
                # Disable the main Move OU button
                $buttonMoveOU.Enabled = $false

                # Log Action
                LogScriptExecution -logPath $logFilePath -action "'$($global:primaryComputer.Name)' has been relocated to $($computer.DistinguishedName)." -userName $env:USERNAME
                
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline 
                Write-Host "Computer '" -NoNewline -ForegroundColor Green
                Write-Host "$($global:primaryComputer.Name)' " -NoNewline 
                Write-Host "was reloacted to " -NoNewline -ForegroundColor Green
                Write-Host "$($computer.DistinguishedName)"

                # Update statusbar message
                UpdateStatusBar "'$($global:primaryComputer.Name)' relocated successfully." -color 'Black'

                # Show Summary dialog box
                [System.Windows.Forms.MessageBox]::Show("'$($global:primaryComputer.Name)' relocated successfully.", "Move OU", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Close the Copy Groups form
                $moveOUForm.Close()

            }
            else {
                # Update statusbar message
                UpdateStatusBar "'$($exampleADComputer) not found." -color 'Red'

                # Display Not-Found dialog
                [System.Windows.Forms.MessageBox]::Show("'$($exampleADComputer) not found.", "Computer not found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                
                $buttonMove.Enabled = $false
                $buttonCancelMoveOU.Enabled = $true

                return $false   
            }

        } finally {
            # Release the lock when done (even if an error occurs)
            ReleaseLock
        }

        $buttonCancelMoveOU.Enabled = $true
    }
}

# Function to display the Move OU form
function ShowMoveOUForm {
    # Create a new form for copying groups
    $moveOUForm = CreateCanvas "Move OU" -x 280 -y 220

    if ([string]::IsNullOrEmpty($global:textboxADComputer.Text)) {
        # Create a label for entering the AD username
        $labelUsername = CreateLabel -text "Example AD username" -x 10 -y 10 -width 150 -height 20
        $moveOUForm.Controls.Add($labelUsername)
    }
    elseif ([string]::IsNullOrEmpty($global:textboxADUsername.Text)) {
        # Create a label for entering the AD Computer
        $labelExampleComputer = CreateLabel -text "Example AD Computer" -x 10 -y 10 -width 200 -height 20
        $moveOUForm.Controls.Add($labelExampleComputer)
    }
    
    # Create a textbox for the Example input
    $textboxExample = CreateTextbox -x 10 -y 30 -width 200 -height 20 -readOnly $false

    # Create a label for OU Name
    $labelOUName = CreateLabel -text "OU Name" -x 10 -y 60 -width 200 -height 20

    # Create OU name textbox
    $textboxOUName = CreateTextbox -x 10 -y 80 -width 200 -height 20 -readOnly $false

    # Add TextChanged event handler to the textbox
    $textboxExample.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMove.Enabled = $false

        } else {
            $buttonMove.Enabled = $true
        }

        if (-not [string]::IsNullOrEmpty($textboxOUName.Text)) {
            $textboxOUName.Text = ""
        }
    })

    # Event handler for mouse down on Example textbox
    $textboxExample.Add_MouseDown({
        $textboxOUName.Text = ""

    })

    # Add TextChanged event handler to the textbox
    $textboxOUName.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMove.Enabled = $false

        } else {
            $buttonMove.Enabled = $true
        }

        if (-not [string]::IsNullOrEmpty($textboxExample.Text)) {
            $textboxExample.Text = ""
        }
    })

    # Event handler for mouse down on OU Name textbox
    $textboxOUName.Add_MouseDown({
        $textboxExample.Text = ""
    })

    # Create the Move button
    $buttonMove = CreateButton -text "Move OU" -x 10 -y 130 -width 70 -height 25 -enabled $false
    $buttonMove.Add_Click({
        $selectedItem = $null
        $ouName = $textboxOUName.Text

        try {
            if ([string]::IsNullOrEmpty($global:textboxADComputer.Text)) {
                if (-not [string]::IsNullOrEmpty($textboxOUName.Text)) {
                    # Fetch OUs
                    $ouNames = Get-ADOrganizationalUnit -Filter {(Name -like $ouName)} -Properties Name
                    if ($null -ne $ouNames) {
                        # Create a form
                        $formList = CreateCanvas "OU Selection" -x 400 -y 300
                        $listBox = CreateListBox -name "OUListBox" -width 350 -height 200 -x 20 -y 20 -items $ouNames
                        $buttonSelect = CreateButton -text "Select" -x 20 -y 230 -width 80 -height 25 -enabled $false
                        $buttonSelect.Add_Click({
                            if ($null -ne $listBox.SelectedItem) {
                                # Perform data validation
                                if (-not [string]::IsNullOrEmpty($global:textboxADUsername.Text)) {
                                    $currentPrimary = FindADUser $global:textboxADUsername.Text
                                    if (-not $currentPrimary -eq $global:primaryUser) {
                                        $currentPrimary = $global:primaryUser
                                    }
                                }

                                $selectedItem = $listBox.SelectedItem.ToString()
                                if ($selectedItem) {
                                    $oldPrimary = $currentPrimary

                                    # Relocate user to designated OU
                                    Move-ADObject -Identity $oldPrimary -TargetPath $selectedItem

                                    $currentPrimary = FindADUser $oldPrimary.SamAccountName
                                    if (-not [string]::IsNullOrEmpty($global:primaryUser)) {
                                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                        Write-Host "$($dateTime) | " -NoNewline
                                        Write-Host "The user " -NoNewline -ForegroundColor Green
                                        Write-Host "'$($oldPrimary.SamAccountName)' " -NoNewline 
                                        Write-Host "was reloacted to " -NoNewline -ForegroundColor Green
                                        Write-Host "$($selectedItem)"

                                        # Update statusbar message
                                        UpdateStatusBar "User '$($oldPrimary.SamAccountName)' was relocated to $($selectedItem)" -color 'Black'
                                    }
                                }

                                $formList.Close()
                                $moveOUForm.Close()
                                $global:buttonMoveOU.Focus()

                            } 
                            else {
                                # Update statusbar message
                                UpdateStatusBar "No item selected." -color 'Yellow'
                                Write-Host "No item selected." -ForegroundColor Yellow
                            }
                        })

                        # Add a Hover event for the ListBox
                        $listBox.Add_MouseDown({
                            $tempSelection = $listBox.SelectedItem.ToString()
                            if (-not [string]::IsNullOrEmpty($tempSelection) -or -not [string]::IsNullOrWhiteSpace($tempSelection)) {$buttonSelect.Enabled = $true}
                            else {$buttonSelect.Enabled = $false}
                        })

                        $buttonCancelList = CreateButton -text "Cancel" -x 300 -y 230 -width 70 -height 25 -enabled $true
                        $buttonCancelList.Add_Click({
                            $formList.Close()
                            $formList.Dispose()
                        })

                        # Add controls to the form
                        $formList.Controls.Add($listBox)
                        $formList.Controls.Add($buttonSelect)
                        $formList.Controls.Add($buttonCancelList)

                        # Show the form
                        $formList.ShowDialog()
                    }
                    else {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline
                        Write-Host "The OU " -NoNewline -ForegroundColor Red
                        Write-Host "'$($ouName)' " -NoNewline
                        Write-Host "was not found." -ForegroundColor Red

                        # Update statusbar message
                        UpdateStatusBar "OU '$($ouName)' was not found." -color 'Red'

                        # Log action
                        LogScriptExecution -logPath $logFilePath -action "OU '$($ouName)' not found." -userName $env:USERNAME

                        # Display Error dialog
                        [System.Windows.Forms.MessageBox]::Show("OU '$($ouName)' not found.", "Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                        return $false
                    }
                }
                else {
                    HandleUser
                }
            }
            elseif ([string]::IsNullOrEmpty($global:textboxADUsername.Text)) {
                if (-not [string]::IsNullOrEmpty($textboxOUName.Text)) {
                    # Fetch all OUs with the name 'Disabled' or containing 'Disabled'
                    $ouNames = Get-ADOrganizationalUnit -Filter {(Name -like $ouName)} -Properties Name
                    if ($null -ne $ouNames) {
                        # Create a form
                        $formList = CreateCanvas "OU Selection" -x 400 -y 300

                        # Create ListBox
                        $listBox = CreateListBox -name "OUListBox" -width 350 -height 200 -x 20 -y 20 -items $ouNames

                        # Create Select button
                        $buttonSelect = CreateButton -text "Select" -x 20 -y 230 -width 80 -height 25 -enabled $false
                        $buttonSelect.Add_Click({
                            if ($null -ne $listBox.SelectedItem) {
                                # Perform data validation
                                if (-not [string]::IsNullOrEmpty($global:textboxADComputer.Text)) {
                                    $currentPrimary = FindADComputer $global:textboxADComputer.Text
                                    if (-not $currentPrimary -eq $global:primaryComputer) {
                                        $currentPrimary = $global:primaryComputer
                                    }
                                }

                                $selectedItem = $listBox.SelectedItem.ToString()
                                if ($selectedItem) {
                                    $oldPrimary = $currentPrimary

                                    # Relocate user to designated OU
                                    Move-ADObject -Identity $oldPrimary -TargetPath $selectedItem

                                    $currentPrimary = FindADComputer $oldPrimary.Name
                                    if (-not [string]::IsNullOrEmpty($global:primaryComputer)) {
                                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                        Write-Host "$($dateTime) | " -NoNewline
                                        Write-Host "Computer " -NoNewline -ForegroundColor Green
                                        Write-Host "'$($currentPrimary.Name)' " -NoNewline 
                                        Write-Host "was reloacted to " -NoNewline -ForegroundColor Green
                                        Write-Host "$($selectedItem)"
                                        
                                        # Update statusbar message
                                        UpdateStatusBar "Computer '$($global:primaryComputer.Name)' was relocated to $($selectedItem)." -color 'Black'

                                        # Log action
                                        LogScriptExecution -logPath $logFilePath -action "Computer '$($global:primaryComputer.Name)' was relocated to $($selectedItem)" -userName $env:USERNAME
                                    }
                                }

                                $formList.Close()
                                $moveOUForm.Close()

                            } else {
                                # Update statusbar message
                                UpdateStatusBar "No item selected." -color 'DarkOrange'

                                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                Write-Host "$($dateTime) | " -NoNewline
                                Write-Host "No item selected." -ForegroundColor Yellow
                            }
                        })

                        # Add a Hover event for the ListBox
                        $listBox.Add_MouseDown({
                            $buttonSelect.Enabled = $true
                        })

                        # Add controls to the form
                        $formList.Controls.Add($listBox)
                        $formList.Controls.Add($buttonSelect)

                        # Show the form
                        $formList.ShowDialog()
                    }
                    else {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline
                        Write-Host "The OU " -NoNewline -ForegroundColor Red
                        Write-Host "'$($ouName)' " -NoNewline
                        Write-Host "was not found." -ForegroundColor Red

                        # Update statusbar message
                        UpdateStatusBar "OU '$($ouName)' not found." -color 'Red'

                        # Log the start of script execution
                        LogScriptExecution -logPath $logFilePath -action "OU '$($ouName)' not found." -userName $env:USERNAME

                        # Display Error dialog
                        [System.Windows.Forms.MessageBox]::Show("OU '$($ouName)' not found.", "Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                        return $false
                    }
                }
                else {HandleComputer}
            }
        } 
        catch {
            # Log the start of script execution
            LogScriptExecution -logPath $global:logFilePath -action "Error: $($_)" -userName $env:USERNAME

            # Update statusbar message
            UpdateStatusBar "An error occured. Check log." -color 'Red'

            # Disable the Copy Groups button and close the form
            $buttonMove.Enabled = $false
            $moveOUForm.Close()

            return $_.Exception.Message
        }
    })

    # Create a 'Cancel' button
    $buttonCancelMoveOU = CreateButton -text "Cancel" -x 180 -y 130 -width 70 -height 25 -enabled $true
    $buttonCancelMoveOU.Add_Click({
        $buttonMoveOU.Enabled = $true
        $buttonFindADUser.Enabled = $false
        
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Move OU canceled." -userName $env:USERNAME

        $buttonMoveOU.Focus()
        $moveOUForm.Close()
        $moveOUForm.Dispose()
    })

    # Add the form controllers
    $moveOUForm.Controls.Add($labelOUName)
    $moveOUForm.Controls.Add($textboxOUName)
    $moveOUForm.Controls.Add($textboxExample)
    $moveOUForm.Controls.Add($buttonMove)
    $moveOUForm.Controls.Add($buttonCancelMoveOU)

    # Focus on the example textbox
    $textboxExample.Select()

    # Monitor window close button (X)
    $moveOUForm.Add_FormClosing({
        $global:buttonMoveOU.Enabled = $true

        # Log Action
        LogScriptExecution -logPath $logFilePath -action "Closed Move OU window." -userName $env:USERNAME

    })

    # Show the Copy Groups form
    $moveOUForm.ShowDialog()

}

# Function to display the CSV User OU relocation form
function ShowCSVMoveUserOUForm {
    # Create a new form for copying groups
    $csvUserMoveOUForm = CreateCanvas "Move OU" -x 280 -y 220
    $csvUserMoveOUForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $csvUserMoveOUForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    if ([string]::IsNullOrEmpty($global:textboxADComputer.Text)) {
        # Create a label for entering the AD username
        $labelUsername = CreateLabel -text "Example AD username" -x 10 -y 10 -width 150 -height 20
        $csvUserMoveOUForm.Controls.Add($labelUsername)
    }
    elseif ([string]::IsNullOrEmpty($global:textboxADUsername.Text)) {
        # Create a label for entering the AD Computer
        $labelExampleComputer = CreateLabel -text "Example AD Computer" -x 10 -y 10 -width 200 -height 20
        $csvUserMoveOUForm.Controls.Add($labelExampleComputer)
    }
    
    # Create a label for entering the AD username
    $labelUsername = CreateLabel -text "Example AD username" -x 10 -y 10 -width 150 -height 20

    # Create a textbox for the Example input
    $textboxExample = CreateTextbox -x 10 -y 30 -width 200 -height 20 -readOnly $false
    
    # Create a label for OU Name
    $labelOUName = CreateLabel -text "OU Name" -x 10 -y 60 -width 200 -height 20

    # Create OU name textbox
    $textboxOUName = CreateTextbox -x 10 -y 80 -width 200 -height 20 -readOnly $false

    # Create the Move button
    $buttonMoveCSVUser = CreateButton -text "Move OU" -x 10 -y 130 -width 70 -height 25 -enabled $false
    $buttonMoveCSVUser.Add_Click({
        if (AcquireLock) {
            try {
                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "Move OU on users CSV file" -userName $env:USERNAME

                $buttonMoveCSVUser.Enabled = $false
                $buttonCancelMoveOU.Enabled = $false

                $selectedItem = $null
                $ouName = $textboxOUName.Text
                if (-not [string]::IsNullOrEmpty($ouName)) {
                    # Fetch OUs
                    $ouNames = Get-ADOrganizationalUnit -Filter {(Name -like $ouName)} -Properties Name
                    if ($null -ne $ouNames) {
                        # Create a form
                        $formList = CreateCanvas "OU Selection" -x 400 -y 300
                        $listBox = CreateListBox -name "OUListBox" -width 350 -height 200 -x 20 -y 20 -items $ouNames
                        $buttonSelect = CreateButton -text "Select" -x 20 -y 230 -width 80 -height 25 -enabled $false
                        $buttonSelect.Add_Click({
                            $buttonSelect.Enabled = $false

                            if (-not [string]::IsNullOrEmpty($listBox.SelectedItem.ToString())) {
                                $selectedItem = $listBox.SelectedItem.ToString()
                                
                                $csvPath = $textboxCSVFilePath.Text
                                $csvData = Import-Csv -Path $csvPath

                                foreach ($row in $csvData) {
                                    # Access the "Username" column from each row
                                    $username = $row.Username
                                    if (-not [string]::IsNullOrEmpty($username)) {
                                        # Get the AD user parameters
                                        $user = FindADUser $username
                            
                                        # Check if the user was found
                                        if (-not [string]::IsNullOrEmpty($user)) {
                                            if ($user.Enabled) {
                                                # Retrieve the distinguished name
                                                $userDistinguishedName = $user.DistinguishedName
                                                
                                                # Relocate user to designated OU
                                                Move-ADObject -Identity $userDistinguishedName -TargetPath $selectedItem
                            
                                                # Display summery in console
                                                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                                Write-Host "$($dateTime) | " -NoNewline 
                                                Write-Host "User " -NoNewline -ForegroundColor Green
                                                Write-Host "'$($username)' " -NoNewline 
                                                Write-Host "relocated to " -NoNewline -ForegroundColor Green
                                                Write-Host "$($selectedItem)." 
                            
                                                # Log action
                                                LogScriptExecution -logPath $global:logFilePath -action "User $($username) relocated to $($selectedItem)." -userName $env:USERNAME
                                            }
                                            else {
                                                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                                Write-Host "$($dateTime) | " -NoNewline 
                                                Write-Host "User " -NoNewline -ForegroundColor Red
                                                Write-Host "'$($username)' " -NoNewline
                                                Write-Host "is disabled. " -NoNewline -ForegroundColor Red
                                                Write-Host "skipping..." -ForegroundColor Yellow
                            
                                                # Log action
                                                LogScriptExecution -logPath $global:logFilePath -action "User '$($username)' is disabled. skipping" -userName $env:USERNAME
                                            }
                                        }
                            
                                        # Apply timer to prevent DOS
                                        Start-Sleep -Milliseconds 300
                                    }
                                }
                            }

                            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                            Write-Host "$($dateTime) | " -NoNewline
                            Write-Host "CSV users relocated successfully." -ForegroundColor Green

                            # Log action
                            LogScriptExecution -logPath $global:logFilePath -action "CSV users relocated successfully." -userName $env:USERNAME

                            # Update statusbar message
                            UpdateStatusBar "CSV users relocated successfully." -color 'Black'

                            $formList.Close()
                            $formList.Dispose()
                            $csvUserMoveOUForm.Close()
                            $csvUserMoveOUForm.Dispose()
                        })

                        # Add a Hover event for the ListBox
                        $listBox.Add_MouseDown({
                            $buttonSelect.Enabled = $true
                        })

                        $buttonCancelList = CreateButton -text "Cancel" -x 300 -y 230 -width 70 -height 25 -enabled $true
                        $buttonCancelList.Add_Click({
                            $formList.Close()
                            $formList.Dispose()
                        })

                        # Add controls to the form
                        $formList.Controls.Add($listBox)
                        $formList.Controls.Add($buttonSelect)
                        $formList.Controls.Add($buttonCancelList)

                        # Show the form
                        $formList.ShowDialog()
                    }
                }
                else {
                    $example = $textboxExample.Text
                    $exampleUser = FindADUser $example

                    if (-not [string]::IsNullOrEmpty($exampleUser)) {
                        # Retrieve the DistinguishedName
                        $exampleDistinguishedName = $exampleUser.DistinguishedName
                        
                        # Handle user CSV data with the example DistinguishedName
                        HandleMoveOUCSV $exampleDistinguishedName

                        $csvUserMoveOUForm.Close()
                        return $true
                    }
                    else {
                        # Write error to console
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline 
                        Write-Host "Example user " -NoNewline -ForegroundColor Red
                        Write-Host "'$($example)' " -NoNewline -ForegroundColor White
                        Write-Host "not found." -ForegroundColor Red

                        # Disable the move button
                        $buttonMove.Enabled = $false
                        $buttonCancelMoveOU.Enabled = $true

                        # Update statusbar message
                        UpdateStatusBar "Example user '$($example)' not found." -color 'Red'

                        # Display Error dialog
                        [System.Windows.Forms.MessageBox]::Show("Example user '$($example)' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                        return $false
                    }
                }
            }
            finally {
                ReleaseLock
            }
        }
    })

    # Create a 'Cancel' button
    $buttonCancelMoveOU = CreateButton -text "Cancel" -x 180 -y 130 -width 70 -height 25 -enabled $true
    $buttonCancelMoveOU.Add_Click({
        $buttonMoveCSVUser.Enabled = $true
        $buttonFindADUser.Enabled = $false
        
        $csvUserMoveOUForm.Close()
        $buttonMoveCSVUser.Focus()

        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Move OU canceled." -userName $env:USERNAME
    })

    # Add TextChanged event handler to the textbox
    $textboxExample.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMoveCSVUser.Enabled = $false
        } else {
            $buttonMoveCSVUser.Enabled = $true
        }
    })

    # Add TextChanged event handler to the textbox
    $textboxOUName.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMoveCSVUser.Enabled = $false

        } else {
            $buttonMoveCSVUser.Enabled = $true
        }

        if (-not [string]::IsNullOrEmpty($textboxExample.Text)) {
            $textboxExample.Text = ""
        }
    })

    # Event handler for mouse down on OU Name textbox
    $textboxOUName.Add_MouseDown({
        $textboxExample.Text = ""
    })

    $csvUserMoveOUForm.Controls.Add($labelUsername)
    $csvUserMoveOUForm.Controls.Add($textboxExample)
    $csvUserMoveOUForm.Controls.Add($labelOUName)
    $csvUserMoveOUForm.Controls.Add($textboxOUName)
    $csvUserMoveOUForm.Controls.Add($buttonMoveCSVUser)
    $csvUserMoveOUForm.Controls.Add($buttonCancelMoveOU)

    # Show the Copy Groups form
    $csvUserMoveOUForm.ShowDialog()
}

# Function to display the CSV User OU relocation form
function ShowCSVMoveComputerOUForm {
    # Create a new form for User CSV
    $csvComputerMoveOUForm = New-Object System.Windows.Forms.Form
    $csvComputerMoveOUForm.Text = "Relocate AD Computer"
    $csvComputerMoveOUForm.Size = New-Object System.Drawing.Size(280, 150)
    $csvComputerMoveOUForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $csvComputerMoveOUForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    # Create a label for entering the AD username
    $labelExampleComputerName = CreateLabel -text "Example Computer Name" -x 10 -y 10 -width 180 -height 20

    # Create a textbox for the Example input
    $textboxExample = CreateTextbox -x 10 -y 30 -width 200 -height 20 -readOnly $false
    
    # Add TextChanged event handler to the textbox
    $textboxExample.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMove.Enabled = $false
        } else {
            $buttonMove.Enabled = $true
        }
    })

    # Create the Move button
    $buttonMove = CreateButton -text "Move OU" -x 10 -y 70 -width 70 -height 25 -enabled $false
    $buttonMove.Add_Click({
        if (AcquireLock) {
            try {
                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "Move OU on computers CSV file" -userName $env:USERNAME

                $buttonMove.Enabled = $false
                $buttonCancelMoveOU.Enabled = $false

                $example = $textboxExample.Text
                $exampleUser = FindADComputer $example

                if (-not [string]::IsNullOrEmpty($exampleUser)) {
                    # Retrieve the DistinguishedName
                    $exampleDistinguishedName = $exampleUser.DistinguishedName
                    
                    # Handle user CSV data with the example DistinguishedName
                    HandleMoveOUCSV $exampleDistinguishedName

                    $csvComputerMoveOUForm.Close()
                    return $true
                }
                else {
                    # Write error to console
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline 
                    Write-Host "Example user " -NoNewline -ForegroundColor Red
                    Write-Host "'$($example)' " -NoNewline -ForegroundColor White
                    Write-Host "not found." -ForegroundColor Red

                    # Update statusbar message
                    UpdateStatusBar "Example computer '$($example)' not found." -color 'Red'

                    # Disable the move button
                    $buttonMove.Enabled = $false
                    $buttonCancelMoveOU.Enabled = $true

                    # Display Error dialog
                    [System.Windows.Forms.MessageBox]::Show("Example computer '$($example)' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                    return $false
                }
            }
            finally {
                ReleaseLock
            }
        }
    })

    # Create a 'Cancel' button
    $buttonCancelMoveOU = CreateButton -text "Cancel" -x 180 -y 70 -width 70 -height 25 -enabled $true
    $buttonCancelMoveOU.Add_Click({
        $buttonMoveOU.Enabled = $true
        $buttonFindADUser.Enabled = $false
        
        $csvComputerMoveOUForm.Close()
        $buttonMoveOU.Focus()

        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Move OU canceled." -userName $env:USERNAME
    })

    $csvComputerMoveOUForm.Controls.Add($labelExampleComputerName)
    $csvComputerMoveOUForm.Controls.Add($textboxExample)
    $csvComputerMoveOUForm.Controls.Add($buttonMove)
    $csvComputerMoveOUForm.Controls.Add($buttonCancelMoveOU)

    # Show the Copy Groups form
    $csvComputerMoveOUForm.ShowDialog()

}
