# Import local modules
$modulePaths = @(
    "Logger",           # Handles logging
    "Lock",             # Manages lock file
    "CsvHandler",       # Handles CSV file operations
    "ADHandler",        # Handles Active Directory operations
    "BuildForm"         # Builds and configures forms
)
foreach ($moduleName in $modulePaths) {
    $modulePath = Join-Path $PSScriptRoot "$($moduleName).psm1"
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
                
                # Log the start of script execution
                LogScriptExecution -logPath $logFilePath -action "'$($global:textboxADUsername.Text)' has been relocated to $($userCheckup.DistinguishedName)." -userName $env:USERNAME

                # Update statusbar message
                UpdateStatusBar "'$($global:textboxADUsername.Text)' has been relocated to '$($userCheckup.DistinguishedName)'." -color 'Black'
                
                # Display results
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "The user '" -NoNewline -ForegroundColor Green
                Write-Host "$($global:primaryUser.SamAccountName)' " -NoNewline 
                Write-Host "was reloacted to " -NoNewline -ForegroundColor Green
                Write-Host "$($userCheckup.DistinguishedName)"
                [System.Windows.Forms.MessageBox]::Show("'$($global:textboxADUsername.Text)' relocated successfully.", "Move OU", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                # Enable buttons
                $buttonFindADUser.Enabled = $true
                $buttonGeneratePassword.Enabled = $true
                $buttonMove.Enabled = $false
                $global:buttonMoveOU.Enabled = $true

                # Close the Copy Groups form
                $moveOUForm.Close()
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
                # Disable the main Move OU button
                $buttonMoveOU.Enabled = $false

                # Relocate user to designated OU
                MoveUserToOU -exampleDisName $computer.DistinguishedName -primaryDisName $global:primaryComputer.DistinguishedName

                # Log Action
                LogScriptExecution -logPath $logFilePath -action "'$($global:primaryComputer.Name)' has been relocated to $($computer.DistinguishedName)." -userName $env:USERNAME
                
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline 
                Write-Host "Computer '" -NoNewline -ForegroundColor Green
                Write-Host "$($global:primaryComputer.Name)' " -NoNewline 
                Write-Host "has been reloacted to " -NoNewline -ForegroundColor Green
                Write-Host "$($computer.DistinguishedName)"
                [System.Windows.Forms.MessageBox]::Show("'$($global:primaryComputer.Name)' relocated successfully.", "Move OU", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                # Update statusbar message
                UpdateStatusBar "'$($global:primaryComputer.Name)' relocated successfully." -color 'Black'

                # Manage buttons
                $global:buttonMoveOU.Enabled = $true
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

    # Add TextChanged event handler to the textbox
    $textboxExample.add_TextChanged({
        if ([string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMove.Enabled = $false

        } else {
            $buttonMove.Enabled = $true
        }

        if (-not [string]::IsNullOrEmpty($textboxOUName.Text)) {
            $textboxOUName.Text = ""
        }
    })

    # Create the Move button
    $buttonMove = CreateButton -text "Move OU" -x 10 -y 130 -width 70 -height 25 -enabled $false
    $buttonMove.Add_Click({
        if (-not [string]::IsNullOrEmpty($global:textboxADUsername.Text)) {HandleUser}
        else {HandleComputer}
    
        $global:buttonGeneratePassword.Enabled = $false
        $global:buttonCopyGroups.Enabled = $false
        $global:buttonRemoveGroups.Enabled = $false
    })

    # Create a 'Cancel' button
    $buttonCancelMoveOU = CreateButton -text "Cancel" -x 180 -y 130 -width 70 -height 25 -enabled $true
    $buttonCancelMoveOU.Add_Click({
        $compText = $textboxADComputer.Text
        $usertext = $textboxADUsername.Text
        $csvText = $textboxCSVFilePath.Text

        if (-not [string]::IsNullOrEmpty($compText)) {
            $buttonFindComputer.Enabled = $false
            $buttonGeneratePassword.Enabled = $false
            $buttonCopyGroups.Enabled = $false
            $buttonRemoveGroups.Enabled = $false
            $buttonMoveOU.Enabled = $true
        }
        elseif (-not [string]::IsNullOrEmpty($usertext)){
            $buttonFindADUser.Enabled = $false
            $buttonGeneratePassword.Enabled = $true
            $buttonCopyGroups.Enabled = $true
            $buttonMoveOU.Enabled = $true
        }
        elseif (-not [string]::IsNullOrEmpty($csvText)) {
            # Read the first line of the CSV file to check the header
            $firstLine = Get-Content -Path $csvText -TotalCount 1
            if ($firstLine -match "Username") {
                $buttonGeneratePassword.Enabled = $true
                $buttonCopyGroups.Enabled = $true
                $buttonMoveOU.Enabled = $true
            }
            elseif ($firstLine -match "ComputerName") {
                $buttonGeneratePassword.Enabled = $false
                $buttonCopyGroups.Enabled = $false
                $buttonRemoveGroups.Enabled = $false
                $buttonMoveOU.Enabled = $true
            }
            else {
                $buttonGeneratePassword.Enabled = $false
                $buttonCopyGroups.Enabled = $false
                $buttonRemoveGroups.Enabled = $false
                $buttonMoveOU.Enabled = $false
                $buttonFindADUser.Enabled = $false
                $buttonFindComputer.Enabled = $false
            }
        }

        # Log Action
        LogScriptExecution -logPath $logFilePath -action "Closed Move OU window." -userName $env:USERNAME

        $buttonMoveOU.Focus()
        $moveOUForm.Close()
        $moveOUForm.Dispose()
    })

    # Add the form controllers
    $moveOUForm.Controls.Add($textboxExample)
    $moveOUForm.Controls.Add($buttonMove)
    $moveOUForm.Controls.Add($buttonCancelMoveOU)

    # Focus on the example textbox
    $textboxExample.Select()

    # Show the Move OU form
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

                # Manage buttons
                $buttonMoveCSVUser.Enabled = $false
                $buttonCancelMoveOU.Enabled = $false

                $selectedItem = $null
                $ouName = $textboxOUName.Text
                if (-not [string]::IsNullOrEmpty($ouName)) {
                    # Fetch OUs
                    $ouNames = Get-ADOrganizationalUnit -Filter {(Name -like $ouName)} -Properties DistinguishedName
                    if ($null -ne $ouNames -or [string]::IsNullOrEmpty($ouNames)) {
                        $selectedItem = ShowListBoxForm -formTitle "OU Selection" -listBoxName "OUListBox" -formWidth 400 -formHeight 300 -listBoxWidth 350 -listBoxHeight 200 -cancelX 300 -cancelY 230 -items $ouNames
                        $selectedOU = $selectedItem[1]
                                
                        # Import CSV
                        $csvPath = $textboxCSVFilePath.Text
                        $csvData = Import-Csv -Path $csvPath

                        # Perform operations on each CSV user
                        foreach ($row in $csvData) {
                            # Access the "Username" column from each row
                            $username = $row.Username
                            if (-not [string]::IsNullOrEmpty($username)) {
                                # Get the AD user parameters
                                $user = FindADUser $username
                                if (-not [string]::IsNullOrEmpty($user)) {
                                    if ($user.Enabled) {
                                        # Retrieve the distinguished name
                                        $userDistinguishedName = $user.DistinguishedName
                                        
                                        # Relocate user to designated OU
                                        Move-ADObject -Identity $userDistinguishedName -TargetPath $selectedOU
                    
                                        # Display summery in console
                                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                        Write-Host "$($dateTime) | " -NoNewline 
                                        Write-Host "User " -NoNewline -ForegroundColor Green
                                        Write-Host "'$($username)' " -NoNewline 
                                        Write-Host "relocated to " -NoNewline -ForegroundColor Green
                                        Write-Host "$($selectedOU)." 
                    
                                        # Log action
                                        LogScriptExecution -logPath $global:logFilePath -action "User $($username) relocated to $($selectedOU)." -userName $env:USERNAME
                                    }
                                    else {
                                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                        Write-Host "$($dateTime) | " -NoNewline 
                                        Write-Host "User " -NoNewline -ForegroundColor Yellow
                                        Write-Host "'$($username)' " -NoNewline
                                        Write-Host "is disabled. " -NoNewline -ForegroundColor Yellow
                                        Write-Host "skipping..." -ForegroundColor Yellow
                    
                                        # Log action
                                        LogScriptExecution -logPath $global:logFilePath -action "User '$($username)' is disabled. skipping" -userName $env:USERNAME
                                    }
                                }
                                else {
                                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                    Write-Host "$($dateTime) | " -NoNewline 
                                    Write-Host "User " -NoNewline -ForegroundColor Red
                                    Write-Host "'$($username)' " -NoNewline
                                    Write-Host "was not found. " -ForegroundColor Red

                                    # Log action
                                    LogScriptExecution -logPath $global:logFilePath -action "User '$($username)' was not found." -userName $env:USERNAME
                                }
                    
                                # Apply timer to prevent DOS
                                Start-Sleep -Milliseconds 300
                            }
                        }
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline
                        Write-Host "CSV users relocated successfully." -ForegroundColor Green

                        # Log action
                        LogScriptExecution -logPath $global:logFilePath -action "CSV users relocated successfully." -userName $env:USERNAME

                        # Update statusbar message
                        UpdateStatusBar "CSV users relocated successfully." -color 'Black'

                        $csvUserMoveOUForm.Close()
                        $csvUserMoveOUForm.Dispose()
                    }
                    else {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline
                        Write-Host "The OU " -NoNewline -ForegroundColor Red
                        Write-Host "'$($ouName)' " -NoNewline
                        Write-Host "was not found." -ForegroundColor Red

                        UpdateStatusBar "'$($ouName)' was not found." -color 'Red'

                        # Log action
                        LogScriptExecution -logPath $global:logFilePath -action "'$($ouName)' was not found." -userName $env:USERNAME

                        # Display Error dialog
                        [System.Windows.Forms.MessageBox]::Show("'$($ouName)' was not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                        $buttonCancelMoveOU.Enabled = $true
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

                        # Manage buttons
                        $buttonMoveCSVUser.Enabled = $false
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
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {$buttonMoveCSVUser.Enabled = $false} 
        else {$buttonMoveCSVUser.Enabled = $true}
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
    # Create form
    $csvComputerMoveOUForm = CreateCanvas "Relocate AD Computer" -x 320 -y 220
    $labelExampleComputerName = CreateLabel -text "Example Computer Name" -x 20 -y 10 -width 250 -height 20
    $textboxExample = CreateTextbox -x 20 -y 30 -width 200 -height 20 -readOnly $false
    $labelOUName = CreateLabel -text "OU Name" -x 20 -y 60 -width 200 -height 20
    $textboxOUName = CreateTextbox -x 20 -y 80 -width 200 -height 20 -readOnly $false
    $buttonMoveComputer = CreateButton -text "Move OU" -x 20 -y 140 -width 70 -height 25 -enabled $false
    $buttonCancelMoveOU = CreateButton -text "Cancel" -x 220 -y 140 -width 70 -height 25 -enabled $true

    # Manage form controls
    $buttonMoveComputer.Add_Click({
        if (AcquireLock) {
            try {
                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "Move OU on computers CSV file" -userName $env:USERNAME

                $buttonMoveComputer.Enabled = $false
                $buttonCancelMoveOU.Enabled = $false

                $selectedItem = $null
                $computerOUName = $textboxOUName.Text
                if (-not [string]::IsNullOrEmpty($computerOUName)) {
                    # Fetch OUs
                    $ouNames = Get-ADOrganizationalUnit -Filter {(Name -like $computerOUName)} -Properties DistinguishedName
                    if ($null -ne $ouNames) {
                        $selectedItem = ShowListBoxForm -formTitle "OU Selection" -listBoxName "OUListBox" -formWidth 400 -formHeight 300 -listBoxWidth 350 -listBoxHeight 200 -cancelX 300 -cancelY 230 -items $ouNames
                        $selectedOU = $selectedItem[1]
                                
                        # Import CSV
                        $csvPath = $textboxCSVFilePath.Text
                        $csvData = Import-Csv -Path $csvPath

                        foreach ($row in $csvData) {
                            # Access the "ComputerName" column from each row
                            $computerName = $row.ComputerName
                            if (-not [string]::IsNullOrEmpty($computerName)) {
                                # Get the AD user parameters
                                $computer = FindADComputer $computerName
                    
                                # Check if the user was found
                                if (-not [string]::IsNullOrEmpty($computer)) {
                                    if ($computer.Enabled) {
                                        # Retrieve the distinguished name
                                        $computerDistinguishedName = $computer.DistinguishedName
                                        
                                        # Relocate user to designated OU
                                        Move-ADObject -Identity $computerDistinguishedName -TargetPath $selectedOU
                    
                                        # Display summery in console
                                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                        Write-Host "$($dateTime) | " -NoNewline 
                                        Write-Host "Computer " -NoNewline -ForegroundColor Green
                                        Write-Host "'$($computer.Name)' " -NoNewline 
                                        Write-Host "has been relocated to " -NoNewline -ForegroundColor Green
                                        Write-Host "$($selectedOU)." 
                    
                                        # Log action
                                        LogScriptExecution -logPath $global:logFilePath -action "Computer $($computer.Name) relocated to $($selectedOU)." -userName $env:USERNAME

                                        Start-Sleep -Milliseconds 300
                                    }
                                    else {
                                        # Log action
                                        LogScriptExecution -logPath $global:logFilePath -action "Computer '$($computer.Name)' is disabled. skipping" -userName $env:USERNAME

                                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                                        Write-Host "$($dateTime) | " -NoNewline 
                                        Write-Host "User " -NoNewline -ForegroundColor Red
                                        Write-Host "'$($computer.Name)' " -NoNewline
                                        Write-Host "is disabled. " -NoNewline -ForegroundColor Red
                                        Write-Host "skipping..." -ForegroundColor Yellow
                                    }
                                }
                                Start-Sleep -Milliseconds 300
                            }
                        }
                    }

                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "CSV computers relocated successfully." -userName $env:USERNAME

                    # Display results
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "CSV computers relocation completed." -ForegroundColor Green

                    # Update statusbar message
                    UpdateStatusBar "CSV computers relocation completed." -color 'Black'

                    # Manage buttons
                    $global:buttonGeneratePassword.Enabled = $false
                    $global:buttonCopyGroups.Enabled = $false
                    $global:buttonRemoveGroups.Enabled = $false

                    $csvComputerMoveOUForm.Close()
                    $csvComputerMoveOUForm.Dispose()
                    }
                else {
                    $example = $textboxExample.Text
                    $exampleComputer = FindADComputer $example

                    if (-not [string]::IsNullOrEmpty($exampleComputer)) {
                        # Retrieve the DistinguishedName
                        $exampleDistinguishedName = $exampleComputer.DistinguishedName
                        
                        # Handle user CSV data with the example DistinguishedName
                        HandleMoveOUCSV $exampleDistinguishedName

                        $csvComputerMoveOUForm.Close()
                        return $true
                    }
                    else {
                        # Display results
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline 
                        Write-Host "Example computer " -NoNewline -ForegroundColor Red
                        Write-Host "'$($example)' " -NoNewline -ForegroundColor White
                        Write-Host "was not found." -ForegroundColor Red
                        [System.Windows.Forms.MessageBox]::Show("Example computer '$($example)' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                        # Update statusbar message
                        UpdateStatusBar "Example computer '$($example)' not found." -color 'Red'

                        # Manage buttons
                        $buttonMoveComputer.Enabled = $false
                        $buttonCancelMoveOU.Enabled = $true

                        return $false
                    }
                }
            }
            finally {ReleaseLock}
        }
    })

    $buttonCancelMoveOU.Add_Click({
        $buttonFindADUser.Enabled = $false
        $buttonGeneratePassword.Enabled = $false
        $buttonCopyGroups.Enabled = $false
        $buttonRemoveGroups.Enabled = $false
        $buttonMoveOU.Enabled = $true

        $csvComputerMoveOUForm.Close()
        $csvComputerMoveOUForm.Dispose()
        $buttonMoveOU.Focus()

        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Move OU canceled." -userName $env:USERNAME
    })

    $textboxExample.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMoveComputer.Enabled = $false
        } else {
            $buttonMoveComputer.Enabled = $true
        }
    })

    $textboxOUName.add_TextChanged({
        if ([string]::IsNullOrWhiteSpace($textboxOUName.Text) -and [string]::IsNullOrEmpty($textboxExample.Text)) {
            $buttonMoveComputer.Enabled = $false

        } else {
            $buttonMoveComputer.Enabled = $true
        }

        if (-not [string]::IsNullOrEmpty($textboxExample.Text)) {
            $textboxExample.Text = ""
        }
    })

    $csvComputerMoveOUForm.Controls.Add($labelExampleComputerName)
    $csvComputerMoveOUForm.Controls.Add($textboxExample)
    $csvComputerMoveOUForm.Controls.Add($labelOUName)
    $csvComputerMoveOUForm.Controls.Add($textboxOUName)
    $csvComputerMoveOUForm.Controls.Add($buttonMoveComputer)
    $csvComputerMoveOUForm.Controls.Add($buttonCancelMoveOU)

    # Show the Copy Groups form
    $csvComputerMoveOUForm.ShowDialog()

}
