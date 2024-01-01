# Import local modules
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import local modules
$modulePaths = @(
    "Logger",           # Handles logging
    "Lock",             # Provides functions for locking and unlocking resources
    "Visuals",          # Handles visual components and UI elements
    "BuildForm"         # Builds and configures forms
)

foreach ($moduleName in $modulePaths) {
    $modulePath = Join-Path $scriptDirectory "$moduleName.psm1"
    Import-Module $modulePath -Force
}

# Function to perform Checkups & csv updates
function PerformComputerActions($computerNames, $csvData) {
    for ($i = 0; $i -lt $computerNames.Count; $i++) {
        $computerName = $computerNames[$i]

        # Check if computer exists in AD
        $computer = FindADComputer $computerName
        if ([string]::IsNullOrEmpty($computer)) {
            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            Write-Host "$($dateTime) | " -NoNewline
            Write-Host "Computer " -NoNewline -ForegroundColor Yellow
            Write-Host "'$($computerName)' " -NoNewline
            Write-Host "not found." -ForegroundColor Yellow

            # Update statusbar message
            UpdateStatusBar "'$($computerName)' not found." -color 'Red'

            $csvData[$i].Results = "Not Found"
            continue  # Skip to the next iteration
        }

        # If computer is disabled, enable it
        if (-not $computer.Enabled) {
            try {
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "Computer " -NoNewline -ForegroundColor Yellow
                Write-Host "'$($computerName)' " -NoNewline
                Write-Host "is disabled." -ForegroundColor Yellow

                # Update statusbar message
                UpdateStatusBar "'$($computerName)' is disabled." -color 'Red'

                # Confirm re-enable action with user
                $confirmReEnable = [System.Windows.Forms.MessageBox]::Show("Computer '$($computer.Name)' is disabled. Re-Enable?", "Re-Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($confirmReEnable -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Enable-AdAccount -Identity $computer
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "Computer " -NoNewline -ForegroundColor Green
                    Write-Host "'$($computer.Name)' " -NoNewline
                    Write-Host "re-enabled successfully." -ForegroundColor Green

                    # Update statusbar message
                    UpdateStatusBar "'$($computerName)' re-enabled successfully." -color 'White'

                    # Update the 'Disabled' value in the CSV to 'TRUE'
                    $csvData[$i].ReEnabled = "TRUE"
                }
                else {
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "Re-Enable skipped on: " -NoNewline -ForegroundColor Yellow
                    Write-Host "'$($computer.Name)'."

                    # Update statusbar message
                    UpdateStatusBar "Re-Enable skipped on: '$($computerName)'." -color 'White'

                    $csvData[$i].ReEnabled = "FALSE"
                }

                Start-Sleep -Milliseconds 500
            }
            catch {
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "Failed to enable " -NoNewline -ForegroundColor Red
                Write-Host "'$($computer.Name)'."
                Write-Host "$($_)" -ForegroundColor Red

                # Update statusbar message
                UpdateStatusBar "Failed to enable '$($computerName)'." -color 'Red'

                # Update the 'Disabled' value in the CSV to 'FALSE'
                $csvData[$i].ReEnabled = "FALSE"
                $csvData[$i].Results = "Failed to Re-Enable"
            }
        }

        # Check if the computer account is locked out
        $isLockedOut = ($computer."msDS-User-Account-Control-Computed" -band 0x00100000) -eq 0x00100000

        # Output the result
        if ($isLockedOut) {
            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            Write-Host "$($dateTime) | " -NoNewline
            Write-Host "Computer account " -NoNewline -ForegroundColor Yellow
            Write-Host "'$($computer.Name)' " -NoNewline
            Write-Host "is locked out." -ForegroundColor Yellow

            # Update statusbar message
            UpdateStatusBar "Computer account '$($computerName)' is locked out." -color 'DarkOrange'

            # Confirm re-enable action with user
            $confirmReEnable = [System.Windows.Forms.MessageBox]::Show("Computer '$($computer.Name)' is disabled. Re-Enable?", "Re-Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirmReEnable -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Unlock the locked-out computer account
                Unlock-AdAccount -Identity $computer

                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "Computer account " -NoNewline -ForegroundColor Green
                Write-Host "'$($computer.Name)' " -NoNewline
                Write-Host "has been unlocked." -ForegroundColor Green

                # Update statusbar message
                UpdateStatusBar "Computer account '$($computerName)' has been unlocked." -color 'White'

                $csvData[$i].Unlocked = "TRUE"
            }
            else {
                $csvData[$i].Unlocked = "FALSE"
            }
        } 

        # Update CSV file
        $csvData[$i].Results = "OK"
        $components = $computer.DistinguishedName -split ','
        $filteredComponents = $components | Where-Object { $_ -notmatch 'CN=' }
        $disName = $filteredComponents -join ','
        $csvData[$i].OldLocation = $disName
    }
}

# Function to Browse and Load the CSV file
function BrowseAndLoadCSV {
    try {
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "CSV Files (*.csv)|*.csv"
        $fileDialog.Title = "Select a CSV File"

        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFilePath = $fileDialog.FileName

            if ([string]::IsNullOrWhiteSpace($selectedFilePath)) {
                Write-Error "ValidateCSVFile: Error: Empty file path field."
                # Draw X mark near the CSV filepath textbox
                HideMark $global:form "CSVPath"
                HideMark $global:form "ADUsername"
                HideMark $global:form "ADComputer"
                DrawXmark $global:form 210 235 "CSVPath"

                # Update statusbar message
                UpdateStatusBar "Error: Empty file path field." -color 'Red'

                $global:textboxCSVFilePath.Text = $selectedFilePath

                $global:buttonGeneratePassword.Enabled = $false
                $global:buttonCopyGroups.Enabled = $false
                $global:buttonMoveOU.Enabled = $false

                return $false
            }
        
            if (-not (Test-Path -Path $selectedFilePath -PathType Leaf)) {
                Write-Error "ValidateCSVFile: File not found: $selectedFilePath"
                # Draw X mark near the CSV filepath textbox
                HideMark $global:form "CSVPath"
                HideMark $global:form "ADUsername"
                HideMark $global:form "ADComputer"
                DrawXmark $global:form 210 235 "CSVPath"

                # Update statusbar message
                UpdateStatusBar "File not found." -color 'Red'

                $global:textboxCSVFilePath.Text = $selectedFilePath
                
                $global:buttonGeneratePassword.Enabled = $false
                $global:buttonCopyGroups.Enabled = $false
                $global:buttonMoveOU.Enabled = $false

                return $false
            }
        
            $fileExtension = [System.IO.Path]::GetExtension($selectedFilePath)
            if ($fileExtension -eq ".csv") {
                # Read the first line of the CSV file to check for the "Username" header
                $firstLine = Get-Content -Path $selectedFilePath -TotalCount 1
                if ($firstLine -match 'Username') {
                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "Loaded users CSV: $($selectedFilePath)" -userName $env:USERNAME

                    # Update statusbar message
                    UpdateStatusBar "Loaded users CSV: $($selectedFilePath)" -color 'White'

                    $global:textboxADUsername.Text = ""
                    $global:textboxADComputer.Text = ""

                    # Display the CSV filepath inside the Path textbox
                    $global:textboxCSVFilePath.Text = $selectedFilePath

                    # Draw a green V mark near the CSV filepath textbox
                    HideMark $global:form "CSVPath"
                    HideMark $global:form "ADUsername"
                    HideMark $global:form "ADComputer"
                    DrawVmark $global:form 210 235 "CSVPath"
                    
                    $global:buttonGeneratePassword.Enabled = $true
                    $global:buttonCopyGroups.Enabled = $false
                    $global:buttonMoveOU.Enabled = $true
                    $global:isCSV = $true

                    return $true
                }
                elseif ($firstLine -match 'ComputerName') {
                    # Read the CSV file
                    $csvData = Import-Csv -Path $selectedFilePath
                    
                    # Define the csv headers that should be included at the end of the process
                    $headers = "ComputerName", "ReEnabled", "UnLocked", "Results", "OldLocation", "NewLocation"

                    # Ensure the headers are present
                    foreach ($header in $headers) {
                        if ($csvData[0].PSObject.Properties.Match($header).Count -eq 0) {
                            $csvData | Add-Member -MemberType NoteProperty -Name $header -Value $null
                        }
                    }

                    # Initialize ArrayLists to store data from each column
                    $computerNames = New-Object System.Collections.ArrayList
        
                    # Process each row in the CSV and populate the ArrayLists
                    $csvData | ForEach-Object {
                        # Check if the value is not empty before adding to ArrayList
                        if (-not [string]::IsNullOrWhiteSpace($_.ComputerName)) {
                            $computerNames.Add($_.ComputerName.ToString())
                        }
                    }
                    # Perform computer actions
                    PerformComputerActions $computerNames $csvData
                
                    # Export the modified CSV data back to the original CSV file
                    $csvData | Export-Csv -Path $selectedFilePath -NoTypeInformation
        
                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "Loaded CSV: $($selectedFilePath)" -userName $env:USERNAME
                    
                    # Update statusbar message
                    UpdateStatusBar "Loaded Computer CSV: $($selectedFilePath)" -color 'White'

                    $textboxADUsername.Text = ""
                    $textboxADComputer.Text = ""

                    # Display the CSV filepath inside the Path textbox
                    $global:textboxCSVFilePath.Text = $selectedFilePath

                    # Draw a green V mark near the CSV filepath textbox
                    HideMark $global:form "CSVPath"
                    HideMark $global:form "ADUsername"
                    HideMark $global:form "ADComputer"
                    DrawVmark $global:form 210 235 "CSVPath"
                    
                    $global:buttonGeneratePassword.Enabled = $false
                    $global:buttonReEnable.Enabled = $false
                    $global:buttonCopyGroups.Enabled = $false
                    $global:buttonMoveOU.Enabled = $true
                    $global:isCSV = $true

                    Start-Process $selectedFilePath
                    return $true
                }
                else {
                    $global:textboxCSVFilePath.Text = $selectedFilePath
                    
                    $global:buttonGeneratePassword.Enabled = $false
                    $global:buttonCopyGroups.Enabled = $false
                    $global:buttonMoveOU.Enabled = $false

                    # Update statusbar message
                    UpdateStatusBar "File or file data Error." -color 'Red'

                    # Draw X mark near the CSV filepath textbox
                    HideMark $global:form "CSVPath"
                    HideMark $global:form "ADUsername"
                    HideMark $global:form "ADComputer"
                    DrawXmark $global:form 210 235 "CSVPath"

                    $global:isCSV = $false
                    return $false
                }
        
            } else {
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "Invalid file format. Only CSV files are allowed." -ForegroundColor Red

                # Update statusbar message
                UpdateStatusBar "Invalid file format. Only CSV files are allowed." -color 'Red'

                # Draw X mark near the CSV filepath textbox
                HideMark $global:form "CSVPath"
                HideMark $global:form "ADUsername"
                HideMark $global:form "ADComputer"
                DrawXmark $global:form 210 235 "CSVPath"

                $global:textboxCSVFilePath.Text = $selectedFilePath

                $global:buttonGeneratePassword.Enabled = $false
                $global:buttonCopyGroups.Enabled = $false
                $global:buttonMoveOU.Enabled = $false

                $global:isCSV = $false
                return $false
            }

        } else {
            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Cenceled CSV file Browse" -userName $env:USERNAME
            $global:isCSV = $false
            return $false
            
        }

    } catch {
        # Log the start of script execution
        LogScriptExecution -logPath $global:logFilePath -action "Catch error: $($_.Exception.Message)" -userName $env:USERNAME

        # Update statusbar message
        UpdateStatusBar "Error: $($_.Exception.Message)." -color 'Red'

        $global:textboxCSVFilePath.Text = $selectedFilePath

        $global:buttonGeneratePassword.Enabled = $false
        $global:buttonCopyGroups.Enabled = $false
        $global:buttonMoveOU.Enabled = $false

        # Draw X mark near the CSV filepath textbox
        HideMark $global:form "CSVPath"
        HideMark $global:form "ADUsername"
        HideMark $global:form "ADComputer"
        DrawXmark $global:form 210 235 "CSVPath"

        # Display Error dialog
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        return $false
    }
}

# Function to export data to csv
function ExportUserData($user, $cacheFilePath, $isExample, $isLocked, $groups) {
    $firstName = $user.GivenName
    $lastName = $user.Surname
    $userDistinguishedName = $user.DistinguishedName
    $userOU = "OU=" + ($userDistinguishedName -split ",OU=", 2)[1]

    # Build a dataset for the csv information
    $userData = @{
        Username = $user.SamAccountName
        FirstName = $firstName
        LastName = $lastName
        Enabled = $user.Enabled
        DistinguishedName = $userOU
        Example = $isExample
        Locked = $isLocked
        Groups = $groups
    }

    # Define dataset
    $csvLine = """$($userData.Username)"", ""$($userData.FirstName)"", ""$($userData.LastName)"", ""$($userData.Enabled)"", ""$($userData.DistinguishedName)"", ""$($userData.Example)"", ""$($userData.Locked)"", ""$($userData.Groups)"""

    # Add dataset to the csv file
    Add-Content -Path $cacheFilePath -Value $csvLine
}

# Function to handle CSV Headers
function HandleHeaders($csvD) {
    # Check if the "Password" header is missing
    if ("Password" -notin $csvD[0].PSObject.Properties.Name) {
        # Add the "Password" header next to the "Username" header
        $csvData | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "Password" -Value $null
        } 2>$null  # Suppress errors in the console

        # Export the modified CSV data back to a temporary CSV file
        $tempPath = [System.IO.Path]::GetTempFileName()
        $csvD | Export-Csv -Path $tempPath -NoTypeInformation

        # Replace the original file with the modified CSV
        Copy-Item -Path $tempPath -Destination $path -Force
        Remove-Item -Path $tempPath
    }

    # Check if the "First Name" header is missing
    if ("FirstName" -notin $csvD[1].PSObject.Properties.Name) {
        # Add the "Password" header next to the "Username" header
        $csvD | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $null
        } 2>$null  # Suppress errors in the console

        # Export the modified CSV data back to a temporary CSV file
        $tempPath = [System.IO.Path]::GetTempFileName()
        $csvD | Export-Csv -Path $tempPath -NoTypeInformation

        # Replace the original file with the modified CSV
        Copy-Item -Path $tempPath -Destination $path -Force
        Remove-Item -Path $tempPath
    }

    # Check if the "Last Name" header is missing
    if ("LastName" -notin $csvD[2].PSObject.Properties.Name) {
        # Add the "Password" header next to the "Username" header
        $csvD | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "LastName" -Value $null
        } 2>$null  # Suppress errors in the console

        # Export the modified CSV data back to a temporary CSV file
        $tempPath = [System.IO.Path]::GetTempFileName()
        $csvD | Export-Csv -Path $tempPath -NoTypeInformation

        # Replace the original file with the modified CSV
        Copy-Item -Path $tempPath -Destination $path -Force
        Remove-Item -Path $tempPath
    }

    # Check if the "Enabled" header is missing
    if ("Disabled" -notin $csvD[3].PSObject.Properties.Name) {
        # Add the "Password" header next to the "Username" header
        $csvD | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "Disabled" -Value $null
        } 2>$null  # Suppress errors in the console

        # Export the modified CSV data back to a temporary CSV file
        $tempPath = [System.IO.Path]::GetTempFileName()
        $csvD | Export-Csv -Path $tempPath -NoTypeInformation

        # Replace the original file with the modified CSV
        Copy-Item -Path $tempPath -Destination $path -Force
        Remove-Item -Path $tempPath
    }

    # Check if the "Enabled" header is missing
    if ("ReEnabled" -notin $csvD[4].PSObject.Properties.Name) {
        # Add the "Password" header next to the "Username" header
        $csvD | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "ReEnabled" -Value $null
        } 2>$null  # Suppress errors in the console

        # Export the modified CSV data back to a temporary CSV file
        $tempPath = [System.IO.Path]::GetTempFileName()
        $csvD | Export-Csv -Path $tempPath -NoTypeInformation

        # Replace the original file with the modified CSV
        Copy-Item -Path $tempPath -Destination $path -Force
        Remove-Item -Path $tempPath
    }
}

# Function to Generate Passwords in the CSV
function GenerateCSVPasswords($path) {
    try {
        # Read the CSV file
        $csvData = Import-Csv -Path $path

        # Validate existence of the additional CSV headers
        HandleHeaders $csvData

        # Display console information
        Write-Host ""
        for ($i = 1; $i -le 40; $i++) {
            $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
            Write-Host "=" -ForegroundColor $currentColor -NoNewline
            Write-Host " " -NoNewline
        }
        Write-Host ""
        Write-Host "Generating Passwords in CSV file: $($path)..."
        for ($i = 1; $i -le 40; $i++) {
            $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
            Write-Host "=" -ForegroundColor $currentColor -NoNewline
            Write-Host " " -NoNewline
        }
        Write-Host ""

        # Access the usernames and update the passwords
        $csvData | ForEach-Object {
            $username = $_.Username
            try {
                # Search user in AD
                $csvUser = FindADUser $username
                if ($csvUser) {
                    $_.FirstName = $csvUser.GivenName
                    $_.LastName = $csvUser.Surname

                    # Check if the user is disabled
                    if (IsUserDisabled $csvUser) {
                        # Add the Disabled status to the "Enabled" header's column
                        $_.Disabled = "Disabled"
                        
                        # Update console
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline 
                        Write-Host "User " -NoNewline -ForegroundColor Yellow
                        Write-Host "'$($username)' " -NoNewline -ForegroundColor White
                        Write-Host "is disabled." -ForegroundColor Yellow
                        
                        # Log action
                        LogScriptExecution -logPath $global:logFilePath -action "User '$($username)' is disabled." -userName $env:USERNAME

                        # Confirm action with user
                        $confirmResult = [System.Windows.Forms.MessageBox]::Show("User '$($username)' is disabled. Re-Enable?", "Re-Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                            $global:primaryUser = $csvUser

                            # Release the script lock to allow the re-enable process
                            ReleaseLock

                            # Display the Re-Enable form
                            ShowReEnableForm

                            # Search user in AD and Update the CSV ReEnabled column
                            $csvPostUser = FindADUser $username
                            if (IsUserDisabled $csvPostUser) {
                                $_.ReEnabled = "FALSE"
                            }
                            else {
                                $_.ReEnabled = "TRUE"
                            }
                        }
                        else {
                            # Update the CSV ReEnabled column
                            $_.ReEnabled = "FALSE"
                        }
                    }

                    # Set user as primary
                    $primary = $csvUser

                    # Extract the initials from the AD user account
                    $initials = ($primary.GivenName.Substring(0, 1).ToUpper()) + ($primary.Surname.Substring(0, 1).ToLower())

                    # Generate a random 6-digit number
                    $randomNumber = GenerateRandomNumber

                    # Define a list of special characters
                    $specialCharacters = '!', '@', '#', '$'

                    # Choose a random special character
                    $randomSpecialChar = $specialCharacters | Get-Random

                    # Generate a password
                    $password = $initials + $randomSpecialChar + $randomNumber

                    # Update the 'Password' column in the CSV data
                    $_.Password = $password
                    
                    # Update console
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline 
                    Write-Host "Password generated for " -NoNewline -ForegroundColor Green
                    Write-Host "'$($username)'." -ForegroundColor White
                    
                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "Password generated for '$($username)'." -userName $env:USERNAME

                    Start-Sleep -Milliseconds 300
                    
                } else {
                    # Update console
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "User " -NoNewline -ForegroundColor Yellow
                    Write-Host "'$($username)' " -NoNewline -ForegroundColor White
                    Write-Host "not found." -ForegroundColor Yellow

                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "User '$($username)' not found." -userName $env:USERNAME

                    Start-Sleep -Milliseconds 300
                }
            } 
            catch {
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                write-Host "An error occurred while processing user " -NoNewline -ForegroundColor Red
                write-Host "'$($username)'. "
                Write-Host "$($_)" -ForegroundColor Red

                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "Error: $($_.Exception.Message)." -userName $env:USERNAME

                # Update statusbar message
                UpdateStatusBar "An error occurred while processing '$($username)'." -color 'Red'
            }
        }

        # Save the modified CSV data back to the file
        $csvData | Export-Csv -Path $path -NoTypeInformation

        # Update statusbar message
        UpdateStatusBar "Passwords have been generated successfully." -color 'White'

        # Log action
        LogScriptExecution -logPath $global:logFilePath -action "Passwords have been generated successfully." -userName $env:USERNAME

        # Open the CSV file
        Start-Process -FilePath $path
    } 
    catch {
        Write-Error "An error occurred while processing the CSV file: $_"

        # Log action
        LogScriptExecution -logPath $global:logFilePath -action "User '$($username)' not found." -userName $env:USERNAME

        # Update statusbar message
        UpdateStatusBar "Error  while processing the CSV file: $_." -color 'Red'
    }
}

# Function to reset the CSV passwords
function ResetCSV {
    $isReset = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reset?", "User Locked", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($isReset -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            # Read the CSV file
            $csvData = Import-Csv -Path $global:textboxCSVFilePath.Text
            
            Write-Host ""
            for ($i = 1; $i -le 40; $i++) {
                $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
                Write-Host "=" -ForegroundColor $currentColor -NoNewline
                Write-Host " " -NoNewline
            }
            Write-Host ""
            Write-Host "Resetting Passwords in CSV file: $($global:textboxCSVFilePath.Text)..."
            for ($i = 1; $i -le 40; $i++) {
                $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
                Write-Host "=" -ForegroundColor $currentColor -NoNewline
                Write-Host " " -NoNewline
            }
            Write-Host ""

            # Access the usernames and update the passwords
            $csvData | ForEach-Object {
                try {
                    $username = $_.Username
                    $password = $_.Password

                    # Perform password reset for the selected user
                    Set-ADAccountPassword -Identity "$username" -NewPassword (ConvertTo-SecureString -AsPlainText "$password" -Force) -Reset

                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline
                    Write-Host "Reset password for " -NoNewline -ForegroundColor Green
                    Write-Host "'$($username)' " -NoNewline -ForegroundColor White
                    Write-Host "completed." -ForegroundColor Green
                    
                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "Reset password completed for '$($username)'." -userName $env:USERNAME

                    # Define delay to prevent AD overloading
                    Start-Sleep -Milliseconds 300
                }
                catch {
                    if ($_.Exception.Message -match "Cannot bind argument to parameter 'String' because it is an empty string") {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline 
                        Write-Host "Password for user " -NoNewline -ForegroundColor Yellow
                        Write-Host "'$($username)' " -NoNewline -ForegroundColor White
                        Write-Host "not found." -ForegroundColor Yellow
                    }
                    else {
                        Write-Warning "Error resetting password for $($username): $_"
                        Write-Host ""
                    }
                }
            }

            $global:buttonResetPassword.Enabled = $false
            $global:buttonGeneratePassword.Enabled = $false

            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Resetting passwords completed successfully." -userName $env:USERNAME

            # Clear the password cell in the csv file
            $clearPass = [System.Windows.Forms.MessageBox]::Show("Do you wish to clear the passwords from the csv?", "Clear Passwords", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($clearPass -eq [System.Windows.Forms.DialogResult]::Yes) {
                $csvData | ForEach-Object {
                    $_.Password = ""
                }
                
                # Save the modified CSV data back to the file
                $csvData | Export-Csv -Path $global:textboxCSVFilePath.Text -NoTypeInformation
                
                # Update statusbar message
                UpdateStatusBar "Passwords cleared from the CSV file." -color 'White'

                # Log action
                LogScriptExecution -logPath $global:logFilePath -action "Passwords cleared from the CSV file." -userName $env:USERNAME

                # Show Clear Password summery dialogbox
                [System.Windows.Forms.MessageBox]::Show("Passwords cleared from the CSV file.", "Clear Passwords", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                $global:buttonGeneratePassword.Enabled = $true
                $global:buttonResetPassword.Enabled = $false
                $global:buttonCopyGroups.Enabled = $false
            }

            # Update statusbar message
            UpdateStatusBar "Resetting passwords completed successfully." -color 'White'

            # Open the CSV file
            Start-Process -FilePath $global:textboxCSVFilePath.Text

            # Show summery dialogbox
            [System.Windows.Forms.MessageBox]::Show("Reset password completed.", "Password Reset", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $global:buttonGeneratePassword.Enabled = $true

            return $true
            }
        catch {
            Write-Host "Error: $_"

            # Update statusbar message
            UpdateStatusBar "Error: $_" -color 'Red'

            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Error: $_" -userName $env:USERNAME

            return $false
        }
    } else {
        return $false
    }
}

# Function to handle CSV file for OU relocating
function HandleMoveOUCSV($exampleDistinguishedName) {
    $csvPath = $textboxCSVFilePath.Text
    $csvData = Import-Csv -Path $csvPath

    Write-Host ""
    for ($i = 1; $i -le 40; $i++) {
        $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
        Write-Host "=" -ForegroundColor $currentColor -NoNewline
        Write-Host " " -NoNewline
    }
    Write-Host ""
    Write-Host "Performing OU Relocation on CSV: $($csvPath)..."
    for ($i = 1; $i -le 40; $i++) {
        $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
        Write-Host "=" -ForegroundColor $currentColor -NoNewline
        Write-Host " " -NoNewline
    }
    Write-Host ""

    foreach ($row in $csvData) {
        # Access the "Username" column from each row
        $username = $row.Username
        if (-not [string]::IsNullOrEmpty($username)) {
            # Get the AD user parameters
            $user = FindADUser $username

            # Check if the user was found
            if ($null -ne $user) {
                if ($user.Enabled) {
                    # Retrieve the distinguished name
                    $userDistinguishedName = $user.DistinguishedName

                    # Perform OU relocation
                    MoveUserToOU -primaryDisName $userDistinguishedName -exampleDisName $exampleDistinguishedName

                    # Display summery in console
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline 
                    Write-Host "User " -NoNewline -ForegroundColor Green
                    Write-Host "'$($username)' " -NoNewline 
                    Write-Host "relocated to " -NoNewline -ForegroundColor Green
                    Write-Host "$($exampleDistinguishedName)." 

                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "User $($username) relocated to $($exampleDistinguishedName)." -userName $env:USERNAME
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

                    continue
                }
            } 

            # Apply timer to prevent DOS
            Start-Sleep -Milliseconds 300
        }
        else {
            # Retrieve computer name from csv row
            $computerName = $row.ComputerName
            $newLocation = $row.NewLocation

            if (-not [string]::IsNullOrEmpty($computerName)) {
                $computer = FindADComputer $computerName
                if ($null -ne $computer) {
                    # Get the computer's Distinguished Name
                    $computerDis = $computer.DistinguishedName
                    
                    # Perform OU relocation
                    MoveUserToOU -primaryDisName $computerDis -exampleDisName $exampleDistinguishedName
                    
                    $components = $exampleDistinguishedName -split ','
                    $filteredComponents = $components | Where-Object { $_ -notmatch 'CN=' }
                    $exampleDisName = $filteredComponents -join ','

                    # Display summery in console
                    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                    Write-Host "$($dateTime) | " -NoNewline 
                    Write-Host "Computer " -NoNewline -ForegroundColor Green
                    Write-Host "'$($computerName)' " -NoNewline 
                    Write-Host "relocated to " -NoNewline -ForegroundColor Green
                    Write-Host "$($exampleDisName)." 

                    # Log action
                    LogScriptExecution -logPath $global:logFilePath -action "Computer $($computerName) relocated to $($exampleDistinguishedName)." -userName $env:USERNAME
                    
                    $row.NewLocation = $exampleDisName

                    # Apply timer to prevent DOS
                    Start-Sleep -Milliseconds 300
                }
            }
        }
    }

    # Export the modified data to the CSV file
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Force

    # Display summery in console
    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    Write-Host "$($dateTime) | " -NoNewline 
    Write-Host "Relocation process completed." -ForegroundColor Green

    # Update statusbar message
    UpdateStatusBar "Relocation process completed successfully." -color 'White'

    # Open CSV file
    #Start-Process $csvPath
}