# Get the directory of the script file
$scriptDirectory = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
$loggerModuleFileName = "Logger.psm1"
$visualsModule = "Visuals.psm1"

# Combine the script directory and file names to get path
$moduleFileName = Join-Path -Path $scriptDirectory -ChildPath $loggerModuleFileName
$visualsFileName = Join-Path -Path $scriptDirectory -ChildPath $visualsModule

# Import Module
Import-Module $moduleFileName -Force
Import-Module $visualsFileName -Force

# Function to generate a random 6-digit number without consecutive or repeated digits
function GenerateRandomNumber {
    do {
        $random = Get-Random -Minimum 100000 -Maximum 999999
        $randomString = $random.ToString()
        $valid = $true

        for ($i = 0; $i -lt 5; $i++) {
            if ($randomString[$i] -eq $randomString[$i + 1]) {
                $valid = $false
                break
            }
        }
    } until ($valid)

    return $random
}

# Function to generate a random password based on initials and a random special character
function GeneratePassword() {
    # Extract the initials from the AD user account
    $initials = ($global:primaryUser.GivenName.Substring(0, 1).ToUpper()) + ($global:primaryUser.Surname.Substring(0, 1).ToLower())
    
    # Generate a random 6 digit number
    $randomNumber = GenerateRandomNumber

    # Define a list of special characters
    $specialCharacters = '!', '@', '#', '$'

    # Choose a random special character
    $randomSpecialChar = $specialCharacters | Get-Random

    # Generate a password
    $password = $initials + $randomSpecialChar + $randomNumber

    # Copy the generated password to the clipboard
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    [System.Windows.Forms.Clipboard]::SetText([System.Runtime.InteropServices.Marshal]::PtrToStringUni([System.Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($securePassword)))

    # Display the generated password
    [System.Windows.Forms.MessageBox]::Show("Generated Password: $password`n`nThe password has been copied to the clipboard.", "Password Generated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    # Build and return the password
    return ($securePassword)

}

# Function to reset the password and disable "Change password at next logon" and clear "Password Never Expires"
function ResetPassword {
    if ($null -ne $global:primaryUser) {
        try {
            # Retrieve the generated password from the clipboard
            $clipboardPassword = [System.Windows.Forms.Clipboard]::GetText()

            # Reset the AD password without requiring a change at next logon
            Set-AdAccountPassword -Identity $global:primaryUser.SamAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $clipboardPassword -Force)
            
            # Disable "Change password at next logon"
            Set-AdUser -Identity $global:primaryUser.SamAccountName -ChangePasswordAtLogon $false

            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Password reset for '$($global:primaryUser.SamAccountName)' completed." -userName $env:USERNAME

            return $true

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error while resetting the password for user '$($global:primary.SamAccountName)': $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            
            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "$($_)" -userName $env:USERNAME
            
            return $false
        }
    }

    else {
        [System.Windows.Forms.MessageBox]::Show("AD user is missing.", "Generate Password: User Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
}

# Function to check if a user is disabled
function IsUserDisabled($user) {
    return (-not $user.Enabled)
}

# Function to check if a user is locked out
function IsUserLockedOut($username) {
    return (Search-ADAccount -LockedOut | Where-Object {$_.SamAccountName -eq $username})
}

# Function to find AD user
function FindADUser($username) {
    $user = Get-AdUser -Filter {SamAccountName -eq $username} -Properties GivenName, Surname, Enabled, DistinguishedName
    if ($user) {
        return $user
    }

    return $null
}

# Function to find AD computer
function FindADComputer($computer) {
    $adComputer = Get-ADComputer -Filter {Name -eq $computer} -Properties *
    if ($adComputer) {
        return $adComputer
    }

    return $null
}

# Function to copy groups from exampleADUser to Re-Enabled user
function CopyGroups {
    param (
        [string]$adUsername,
        [string]$exampleADuser
    )

    $userGroups = Get-AdPrincipalGroupMembership -Identity $exampleADuser |
        Select-Object -Property Name, SamAccountName |
        ForEach-Object {
            $_.Name
            $_.SamAccountName
        } | Select-Object -Unique
    
    # Filter out "Domain Users"
    $userGroups = $userGroups | Where-Object { $_ -ne 'Domain Users' }
    
    Write-Host ""
    for ($i = 1; $i -le 40; $i++) {
        $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
        Write-Host "=" -ForegroundColor $currentColor -NoNewline
        Write-Host " " -NoNewline
    }
    Write-Host ""
    Write-Host "Copying groups from $($exampleADuser)..."
    for ($i = 1; $i -le 40; $i++) {
        $currentColor = if ($i % 2 -eq 0) { 'Blue' } else { 'White' }
        Write-Host "=" -ForegroundColor $currentColor -NoNewline
        Write-Host " " -NoNewline
    }
    Write-Host ""

    if ($userGroups.Count -eq 0) {
        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        Write-Host "$($dateTime) | " -NoNewline
        Write-Host "No Groups." -ForegroundColor Yellow
        return $false
    }
    
    # Add the user to the groups of the specified $adUsername
    foreach ($group in $userGroups) {
        try {
            Add-ADGroupMember -Identity $group -Members $global:primaryUser.SamAccountName -ErrorAction Continue # SilentlyContinue # Stop ||

            # Set a timer to avoid request flooding | DOS
            Start-Sleep -Milliseconds 300

            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            Write-Host "$($dateTime) | " -NoNewLine 
            Write-Host "User " -NoNewline -ForegroundColor Green
            Write-Host "'$($global:primaryUser.SamAccountName)' " -NoNewline
            Write-Host "joined " -NoNewline -ForegroundColor Green
            Write-Host "$($group)."

            # Log action
            LogScriptExecution -logPath $logFilePath -action "$($global:primaryUser.SamAccountName) joined $($group)." -userName $env:USERNAME
        } 

        catch [Microsoft.ActiveDirectory.Management.ADException] {
            # Check if the error message contains the string indicating that the member already exists.
            if ($_.Exception.Message -match "specified user account is already a member of the specified group") {
                # Set a timer to avoid request flooding | DOS
                Start-Sleep -Milliseconds 300

            } else {
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"

                # Log action
                LogScriptExecution -logPath $logFilePath -action "$($_.Exception.Message)" -userName $env:USERNAME

                # Handle other errors (if needed).
                Write-Error "An error occurred: $($_.Exception.Message)"
            }
        }
    }

    # Retrieve and print the groups the user is a member of (only group names)
    $userGroups = Get-ADUser -Identity $global:primaryUser.SamAccountName -Properties MemberOf | 
    Select-Object -ExpandProperty MemberOf | 
    ForEach-Object { (Get-ADGroup $_).Name }

    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    Write-Host ""
    Write-Host "-------------------------------------------------------------------------------"
    Write-Host "$($dateTime) | User $($global:primaryUser.SamAccountName) is a member of the following groups:"
    foreach ($group in $userGroups) {
        Write-Host " *** $group" -ForegroundColor Cyan
    }
    Write-Host ""
}

# Function to remove groups from user
function RemoveGroups($username) {
    try {
        # Get the user's groups
        $userGroups = Get-ADUser -Identity $username -Properties MemberOf | Select-Object -ExpandProperty MemberOf

        if ($userGroups.Count -eq 0) {
            $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            Write-Host "$($dateTime) | " -NoNewline
            Write-Host "User " -NoNewline -ForegroundColor Yellow
            Write-Host "'$($username)' " -NoNewline
            Write-Host "is not a member of any group." -ForegroundColor Yellow
            return $null
        }

        # Filter out the "Domain Users" group
        $groupsToRemove = $userGroups | Where-Object { $_ -ne "Domain Users" }

        # Remove the user from each group
        foreach ($group in $groupsToRemove) {
            try {
                Remove-ADGroupMember -Identity $group -Members $username -Confirm:$false -ErrorAction Stop
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "User " -NoNewline -ForegroundColor Green
                Write-Host "'$($username)' " -NoNewline
                Write-Host "has been removed from " -NoNewline -ForegroundColor Green
                Write-Host "$($group)."

                # Log action
                LogScriptExecution -logPath $logFilePath -action "$($username) has been removed from $($group)." -userName $env:USERNAME

                Start-Sleep -Milliseconds 300

            } catch {
                $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                Write-Host "$($dateTime) | " -NoNewline
                Write-Host "Error removing user " -NoNewline -ForegroundColor Red
                Write-Host "'$($username)'" -NoNewline
                Write-Host "from group " -NoNewline -ForegroundColor Red
                Write-Host "'$($group)'."
                Write-Host "$_" -ForegroundColor Red
                
                # Log action
                LogScriptExecution -logPath $logFilePath -action "Error removing user $username from group $group. $_." -userName $env:USERNAME

                # Skip to the next entry
                continue
            }
        }
        return $true

    } catch {
        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        Write-Host "$($dateTime) | " -NoNewline
        Write-Host "Error retrieving groups for user " -NoNewline -ForegroundColor Red 
        Write-Host "'$($username)'."
        Write-Host "$_" 
        
        # Log action
        LogScriptExecution -logPath $logFilePath -action "Error retrieving groups for user $username. $_" -userName $env:USERNAME

        return $false
    }
}

# Function to move users to OU
function MoveUserToOU {
    param (
        [string]$exampleDisName,
        [string]$primaryDisName,
        [bool]$isUser
    )

    $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    try {
        # Extract the OU from the DistinguishedName
        $exampleOU = "OU=" + ($exampleDisName -split ",OU=",2)[1]

        # Move the user to the same OU as exampleUser
        Move-ADObject -Identity $primaryDisName -TargetPath $exampleOU

        return $true

    } catch {
        Write-Host "$($dateTime) | An error occurred: $_"

        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "$($_)" -userName $env:USERNAME

        return $_
    }
}

# Function to handle a locked-out account
Function HandleLockedOut {
    # Log action
    LogScriptExecution -logPath $logFilePath -action "$($global:primaryUser.SamAccountName) is locked-out." -userName $env:USERNAME

    # Display Unlock dialog box
    $unlockAccountResult = [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' is locked. Do you want to unlock the account?", "User Locked", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($unlockAccountResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Log action
        LogScriptExecution -logPath $logFilePath -action "$($global:primaryUser.SamAccountName) was unlocked." -userName $env:USERNAME

        # Unlock the user account
        Unlock-ADAccount -Identity $global:primaryUser.SamAccountName
        [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' has been unlocked.", "Account Unlocked", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Show a green V checkmark near the textbox
        DrawVmark $global:form 125 47 "ADUsername"
    }

    # Update buttons
    $global:buttonGeneratePassword.Enabled = $true
    $global:buttonCopyGroups.Enabled = $true
    $global:buttonMoveOU.Enabled = $true
}

# Function to manage the FindADuser click
function ManageFindADUserEvent {
    $global:buttonFindADUser.Enabled = $false

    # Get the AD user from the textbox
    $adUsername = $global:textboxADUsername.Text
    if (-not [string]::IsNullOrWhiteSpace($adUsername)) {
        # Send AD request to find the user
        $global:primaryUser = FindADUser $adUsername
        if ($global:primaryUser) {
            # Check if the user is Disabled
            if (-not $global:primaryUser.Enabled) {
                # Log action
                LogScriptExecution -logPath $logFilePath -action "User '$($global:primaryUser.SamAccountName)' is disabled.", "User Disabled" -userName $env:USERNAME

                # Update buttons
                $global:buttonFindADUser.Enabled = $false
                $global:buttonReEnable.Enabled = $true

                # Hide the green V checkmark
                HideMark $form "ADUsername"
                HideMark $form "ADComputer"
                HideMark $form "CSVPath"
                DrawXmark $Form 130 46 "ADUsername"

                # Show the User is disabled dialog box
                [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' is disabled.", "User Disabled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)

                return $false
            }

            # Check if the user is Locked-Out
            $locked = IsUserLockedOut $global:primaryUser.SamAccountName
            if ($locked) {
                HandleLockedOut
            } 
            else {
                # Show a green V checkmark near the textbox
                HideMark $form "ADUsername"
                DrawVmark $form 125 47 "ADUsername"
    
                # Manage buttons
                $global:buttonGeneratePassword.Enabled = $true
                $global:buttonCopyGroups.Enabled = $true
                $global:buttonRemoveGroups.Enabled = $true
                $global:buttonMoveOU.Enabled = $true
                
                # Set focus on the Find and Generate button
                $global:buttonGeneratePassword.Focus()

                return $true
            }
        } 
        else {
            # Draw red X mark near the username textbox
            HideMark $form "ADUsername"
            DrawXmark $form 130 46 "ADUsername"

            # Display dialog box
            [System.Windows.Forms.MessageBox]::Show("'$($adUsername)' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

            # Disable the FindADUser button
            $global:buttonFindADUser.Enabled = $false

            return $false
        }

    } else {
        # Display dialog box
        [System.Windows.Forms.MessageBox]::Show("Please enter an AD username.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        
        # Draw red X mark near the username textbox
        HideMark $form "ADUsername"
        DrawXmark $form 130 44 "ADUsername"
        
        return $false
    }
}

# Function to process the Reset Password validation
function ManageResetPasswordValidationEvent {
    $userTextBox = $global:textboxADUsername.Text
    if (-not [string]::IsNullOrWhiteSpace($userTextBox)) {
        # Display a validation dialog box
        $confirmDialogResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reset the password for user '$userTextBox'?", "Confirm Password Reset", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmDialogResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                # Reset the AD user's password
                $reset = ResetPassword
                if ($reset) {
                    # Log the start of script execution
                    LogScriptExecution -logPath $global:logFilePath -action "Password reset for user '$($userTextBox)'" -userName $env:USERNAME

                    # Display a summery information dialog box
                    [System.Windows.Forms.MessageBox]::Show("Password for user '$userTextBox' has been reset.", "Password Reset", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                    
                    $global:buttonResetPassword.Enabled = $false
                    $global:buttonFindADUser.Enabled = $false
                    if (-not $global:buttonMoveOU.Enabled) {
                        $global:buttonMoveOU.Enabled = $false
                    }

                    return $true
                }

            } catch {
                return $_.Exception.Message
            }
        }
    } 
    else {
        $global:buttonResetPassword.Enabled = $false
        $global:buttonFindADUser.Enabled = $false
        if (-not $global:buttonMoveOU.Enabled) {
            $global:buttonMoveOU.Enabled = $false
        }

        # Display dialog box
        [System.Windows.Forms.MessageBox]::Show("No AD username.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Function to manage the Find Computer click
function ManageFindComputerEvent {
    # Get the computer object from Active Directory
    $global:primaryComputer = FindADComputer $global:textboxADComputer.Text
    if ($global:primaryComputer) {
        # Check if the user is Disabled
        if (-not $global:primaryComputer.Enabled) {
            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Computer '$($global:primaryComputer.Name)' is disabled." -userName $env:USERNAME

            # Update buttons
            $global:buttonFindComputer.Enabled = $false
            $global:buttonMoveOU.Enabled = $false
            $global:buttonReEnable.Enabled = $true
            $global:buttonRemoveGroups.Enabled = $false

            # Hide the green V checkmark
            HideMark $form "ADUsername"
            HideMark $form "ADComputer"
            HideMark $form "CSVPath"
            DrawXmark $form 130 140 "ADComputer" 

            # Show the User is disabled dialog box
            [System.Windows.Forms.MessageBox]::Show("Computer '$($global:primaryComputer.Name)' is disabled", "Computer Disabled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)

            $global:buttonReEnable.Focus()
            return $false
        }

        # Show a green V checkmark near the textbox
        HideMark $form "ADComputer"
        DrawVmark $form 127 135 "ADComputer"

        # Manage buttons
        $global:buttonMoveOU.Enabled = $true
        $global:buttonRemoveGroups.Enabled = $false
        
        # Set focus on the Find and Generate button
        $global:buttonMoveOU.Focus()

        return $true

    }
    else {
        # Draw X
        HideMark $form "ADUsername"
        HideMark $form "ADComputer"
        HideMark $form "CSVPath"
        DrawXmark $form 130 140 "ADComputer" 

        # Log action
        LogScriptExecution -logPath $global:logFilePath -action "Computer '$($global:textboxADComputer.Text)' not found" -userName $env:USERNAME

        # Show the User is disabled dialog box
        [System.Windows.Forms.MessageBox]::Show("Computer '$($global:textboxADComputer.Text)' not found", "Computer Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

        $global:buttonRemoveGroups.Enabled = $false
        $global:textboxADComputer.Focus()
        return $false
    }
}