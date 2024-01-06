# 1. Complete MoveOU for CSV by listBox selection.

# Import form modules
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System

# Set the console code page to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Reset the console to its default encoding using Out-Default
$null = Out-Default

# Get the script's dir
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import local modules
$modulePaths = @(
    "Logger",               # Handles logging
    "Lock",                 # Provides functions for locking and unlocking resources
    "Visuals",              # Handles visual components
    "ReEnableForm",         # Manages re-enabling functionality
    "CopyGroups",           # Handles copying groups
    "moveOUForm",           # Manages the Move OU forms
    "CsvHandler",           # Handles CSV file operations
    "ADHandler",            # Handles Active Directory operations
    "BuildForm",            # Builds and configures forms
    "MenuStrip",            # Builds the top menu
    "PasswordSettings"      # Set the password's strength
)

foreach ($moduleName in $modulePaths) {
    $modulePath = Join-Path $PSScriptRoot "Modules\$moduleName.psm1"
    Import-Module $modulePath -Force
}

# Display loaded modules
#Get-Module

# Specify the file names
$logFileName = "Quick-AD.log"

# Combine the script directory and file names to get path
$global:logFilePath = Join-Path -Path $scriptDirectory -ChildPath $logFileName

# Ensure the log file exists or create it
if (-not (Test-Path -Path $logFilePath)) {
    New-Item -Path $logFilePath -ItemType File -Force | Out-Null
}

# ============== FUNCTIONS =================== #
# Function to manage the close event handler
function ManageFormClose() {
    $choice = [System.Windows.Forms.MessageBox]::Show("Confirm Exit", "Confirm Exit", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    elseif ($choice -eq [System.Windows.Forms.DialogResult]::No) {
        # Cancel the form closing
        $_.Cancel = $true
    }
}

function Main {
    # Define vars
    $global:version = "1.0.0"           # Application version
    $global:primaryUser = $null         # Store information about the primary user
    $global:primaryComputer = $null     # Store information about the primary computer
    $global:isComputer = $false         # Flag indicating whether the primary user is a computer
    $global:isCSV = $false              # Flag indicating whether the input is a CSV file
    $global:totalDigits = 6             # Default number of digits for random number generation
    $global:statusBarMessage = $null    # Message for the status bar
    $global:showTooltips = $true        # Flag to show/hide tooltips

    # Set the Password Settings default variables
    $global:specialChars = '!@#$'.ToCharArray()     # Define special characters for passwords
    $global:passDefaultUpperNum = 1                 # Set the default number of uppercase letters in the password
    $global:passDefaultLowerNum = 1                 # Set the default number of lowercase letters in the password
    $global:passDefaultSpecialsNum = 1              # Set the default number of special characters in the password
    $global:passDefaultDigitsNum = 6                # Set the default number of digits in the password
    $global:passDefaultIncludeInitials = $true      # Specify whether to include initials in the password
    $global:passShuffle = $false                    # Specify whether to shuffle the characters in the password

    # Set names for visual validations
    $adUserNamePictureName = "ADUsername"
    $adComputerPictureName = "ADComputer"
    $csvPathPictureName = "CSVPath"
    $prefixes = @($adUserNamePictureName, $adComputerPictureName, $csvPathPictureName)

    # Create the main form
    $global:form = CreateCanvas "QUICK-AD" -x 470 -y 380
    $form.MaximizeBox = $false
    $form.Margin = '0,0,0,0'
    $form.Padding = '0,0,0,0'

    # Create a top menu strip
    $global:menuStrip = CreateMenuStrip

    # Create labels
    $labelADUsername = CreateLabel -text "User Name" -x 10 -y 35 -width 120 -height 20
    $labelADComputerName = CreateLabel -text "Computer Name" -x 10 -y 125 -width 120 -height 20
    $labelCSVFilePath = CreateLabel -text "CSV file path" -x 10 -y 215 -width 70 -height 20

    # Create textboxes
    $global:textboxADUsername = CreateTextbox -x 10 -y 55 -width 120 -height 10 -readOnly $false
    $global:textboxADComputer = CreateTextbox -x 10 -y 145 -width 120 -height 10 -readOnly $false
    $global:textboxCSVFilePath = CreateTextbox -x 10 -y 240 -width 200 -height 10 -readOnly $true

    # Create buttons
    $global:buttonFindADUser = CreateButton -text "Find User" -x 10 -y 85 -width 80 -height 25 -enabled $false
    $global:buttonFindComputer = CreateButton -text "Find Computer" -x 10 -y 175 -width 100 -height 25 -enabled $false
    $global:buttonBrowseForCSVFile = CreateButton -text "Browse" -x 10 -y 270 -width 60 -height 25 -enabled $true
    $global:buttonGeneratePassword = CreateButton -text "Generate Password" -x 310 -y 40 -width 120 -height 25 -enabled $false
    $global:buttonResetPassword = CreateButton -text "Reset Password" -x 310 -y 80 -width 120 -height 25 -enabled $false
    $global:buttonReEnable = CreateButton -text "Re-Enable" -x 310 -y 120 -width 120 -height 25 -enabled $false
    $global:buttonCopyGroups = CreateButton -text "Copy Groups" -x 310 -y 160 -width 120 -height 25 -enabled $false
    $global:buttonRemoveGroups = CreateButton -text "Remove Groups" -x 310 -y 200 -width 120 -height 25 -enabled $false
    $global:buttonMoveOU = CreateButton -text "Move OU" -x 310 -y 240 -width 120 -height 25 -enabled $false

    if ($global:showTooltips -eq $true) {
        # Create tooltips
        CreateToolTip -control $buttonFindADUser -text "Search for an Active Directory User account"
        CreateToolTip -control $buttonFindComputer -text "Search for an Active Directory Computer account"
        CreateToolTip -control $buttonBrowseForCSVFile -text "Browse for a CSV file with the 'Username' or 'ComputerName' headers"
        CreateToolTip -control $buttonGeneratePassword -text "Generate passwords for selected user or CSV file"
        CreateToolTip -control $buttonResetPassword -text "Reset password for selected user or CSV file"
        CreateToolTip -control $buttonReEnable -text "Re-Enable User/Computer with an example AD account"
        CreateToolTip -control $buttonCopyGroups -text "Copy groups from an example AD user account"
        CreateToolTip -control $buttonRemoveGroups -text "Remove single/CSV user's group memberships"
        CreateToolTip -control $buttonMoveOU -text "Relocate user/computer to OU by example account or from CSV file"
    }

    # Call the function to create the status bar
    CreateStatusBar

    # Update statusbar message
    UpdateStatusBar "Ready" -color 'Black'

    # = = = = = = = EVENT HANDLERS = = = = = = = #
    $textboxADUsername.Add_MouseDown({
        $textboxADComputer.Text = ""
        $textboxCSVFilePath.Text = ""

        # AD Username controller's conditions - buttons and textboxes management
        if ([string]::IsNullOrEmpty($textboxADUsername.Text)) {HideAllMarks $form $prefixes}
        else {
            HideMark $form $adComputerPictureName
            HideMark $form $csvPathPictureName
        }

        # Manage FindComputer button
        if (-not [string]::IsNullOrEmpty($global:primaryUser) -and -not [string]::IsNullOrEmpty($textboxADUsername.Text)) {$buttonFindComputer.Enabled = $false}
        
        # Manage Generate Password button
        if (-not $buttonGeneratePassword.Enabled) {$buttonGeneratePassword.Enabled = $false} else {$buttonGeneratePassword.Enabled = $true}
        
        # Manage the rest of the controllers
        if (-not [string]::IsNullOrEmpty($global:primaryUser)) {
            if ($buttonGeneratePassword.Enabled) {$buttonGeneratePassword.Enabled = $true} else {$buttonGeneratePassword.Enabled = $false }
            if ($buttonCopyGroups.Enabled) {$buttonCopyGroups.Enabled = $true} else {$buttonCopyGroups.Enabled = $false}
            if ($buttonRemoveGroups.Enabled) {$buttonRemoveGroups.Enabled = $true} else {$buttonRemoveGroups.Enabled = $false}
            if ($buttonMoveOU.Enabled) {$buttonMoveOU.Enabled = $true} else {$buttonMoveOU.Enabled = $false}
        }
        else {
            $buttonGeneratePassword.Enabled = $false
            $buttonResetPassword.Enabled = $false
            $buttonReEnable.Enabled = $false
            $buttonCopyGroups.Enabled = $false
            $buttonRemoveGroups.Enabled = $false
            $buttonMoveOU.Enabled = $false
        }

        if ([string]::IsNullOrEmpty($textboxADUsername.Text)) {UpdateStatusBar "Ready" -color 'Black'}
    })

    $textboxADComputer.Add_MouseDown({
        HideMark $form $adUserNamePictureName
        HideMark $form $csvPathPictureName

        $textboxADUsername.Text = ""
        $textboxCSVFilePath.Text = ""
        $buttonFindADUser.Enabled = $false

        if ([string]::IsNullOrEmpty($textboxADComputer.Text)) {
            $buttonFindComputer.Enabled = $false
            $buttonGeneratePassword.Enabled = $false
            $buttonResetPassword.Enabled = $false
            $buttonCopyGroups.Enabled = $false
            $buttonRemoveGroups.Enabled = $false
            $buttonReEnable.Enabled = $false
            $buttonMoveOU.Enabled = $false
        }
        else {
            if (-not $buttonFindComputer.Enabled) {$buttonFindComputer.Enabled = $false} else {$buttonFindComputer.Enabled = $true}
            if (-not $buttonMoveOU.Enabled) {$buttonMoveOU.Enabled = $false} else {$buttonMoveOU.Enabled = $true}
            if ($buttonRemoveGroups.Enabled) {$buttonRemoveGroups.Enabled = $false}
            if ($buttonCopyGroups.Enabled) {$buttonCopyGroups.Enabled = $false}
        }

        if ([string]::IsNullOrEmpty($textboxADComputer.Text)) {UpdateStatusBar "Ready" -color 'Black'}
    })

    $textboxADUsername.add_TextChanged({
        if (-not [string]::IsNullOrEmpty($textboxCSVFilePath.Text)) {
            HideMark $form $csvPathPictureName
            $textboxCSVFilePath.Text = ""
        }
        
        # Manage Find button
        if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)) {
            $buttonFindADUser.Enabled = $true
        }
        else {
            $buttonFindADUser.Enabled = $false
            if (-not [string]::IsNullOrEmpty($textboxADComputer.Text)) {
                $buttonFindComputer.Enabled = $true
            }
        }
        
        # Disable controller's action buttons
        $buttonGeneratePassword.Enabled = $false
        $buttonResetPassword.Enabled = $false
        $buttonReEnable.Enabled = $false
        $buttonCopyGroups.Enabled = $false
        $buttonRemoveGroups.Enabled = $false
        $buttonMoveOU.Enabled = $false
        
        # Hide result mark
        HideMark $form $adUserNamePictureName
    })

    $textboxADComputer.add_TextChanged({
        # Clear CSV textbox & hide CSV mark
        if (-not [string]::IsNullOrEmpty($textboxCSVFilePath.Text)) {
            HideMark $form $csvPathPictureName
            $textboxCSVFilePath.Text = ""
        }
        
        # Manage Find button
        if (-not [string]::IsNullOrEmpty($textboxADComputer.Text)) {
            $buttonFindComputer.Enabled = $true
        }
        else {
            $buttonFindComputer.Enabled = $false
            if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)) {
                $buttonFindADUser.Enabled = $true
            }
        }

        # Disable controller's action buttons
        $buttonGeneratePassword.Enabled = $false
        $buttonResetPassword.Enabled = $false
        $buttonReEnable.Enabled = $false
        $buttonCopyGroups.Enabled = $false
        $buttonMoveOU.Enabled = $false

        # Hide result mark
        HideMark $form $adComputerPictureName
    })

    $buttonFindADUser.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "FindADUser: '$($textboxADUsername.Text)'" -userName $env:USERNAME
        
        $global:isComputer = $false
        $global:isCSV = $false

        if ($buttonFindComputer.Enabled) {
            $buttonFindComputer.Enabled = $false
        }

        # Acquire the lock
        if (AcquireLock) {
            try {
                ManageFindADUserEvent
            
            } finally {
                # Release the lock when done (even if an error occurs)
                ReleaseLock
            }
        }
    })

    $buttonFindComputer.Add_Click({
        $buttonFindComputer.Enabled = $false
        $global:isComputer = $true
        $global:isCSV = $false

        if ($buttonFindADUser.Enabled) {
            $buttonFindADUser.Enabled = $false
        }

        $buttonGeneratePassword.Enabled = $false
        $buttonCopyGroups.Enabled = $false

        # Manage Event
        ManageFindComputerEvent 
    })

    $buttonBrowseForCSVFile.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Browse" -userName $env:USERNAME

        # Browse & Load CSV file
        BrowseAndLoadCSV

    })

    $buttonGeneratePassword.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Generate Password" -userName $env:USERNAME

        # Grab the text from the AD Username input field
        $userText = $textboxADUsername.Text
        if ($userText -eq $global:primaryUser.SamAccountName -and $textboxCSVFilePath.Text -eq "") {
            # Generate password
            GeneratePassword $global:totalDigits

            # Manage buttons
            if (-not $buttonMoveOU.Enabled) {
                $buttonResetPassword.Enabled = $true
                $buttonGeneratePassword.Enabled = $true
                $buttonCopyGroups.Enabled = $true
                $buttonMoveOU.Enabled = $false
            }
            else {
                $buttonResetPassword.Enabled = $true
                $buttonGeneratePassword.Enabled = $true
                $buttonCopyGroups.Enabled = $true
                $buttonMoveOU.Enabled = $true
            }

        } else {
            # Grab the text from the csv filepath input field
            $filePathToProcess = $textboxCSVFilePath.Text

            # Acquire the lock
            if (AcquireLock) {
                try {
                    # Generate password
                    GenerateCSVPasswords $filePathToProcess
                
                } finally {
                    # Release the lock when done (even if an error occurs)
                    ReleaseLock
                }
            }

            # Manage buttons
            $buttonResetPassword.Enabled = $true
            $buttonGeneratePassword.Enabled = $true
            $buttonCopyGroups.Enabled = $false
            if (-not $buttonMoveOU.Enabled) {
                $buttonMoveOU.Enabled = $false
            }
            else {
                $buttonMoveOU.Enabled = $true
            }
        }
    })

    $buttonResetPassword.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Reset Password" -userName $env:USERNAME

        # Grab the text from the CSV filepath textbox
        $filepath = $textboxCSVFilePath.Text

        if ($filepath -eq "") {
            # CSV file path is empty. Work on a single user.
            ManageResetPasswordValidationEvent @parameters
        }

        else {
            # Acquire the lock
            if (AcquireLock) {
                try {
                    ResetCSV
                }
                finally {
                    # Release the lock when done (even if an error occurs)
                    ReleaseLock
                }
            }
        }

        # Manage buttons
        $buttonFindADUser.Enabled = $false
        $buttonGeneratePassword.Enabled = $true
        $buttonCopyGroups.Enabled = $true
        if ($buttonRemoveGroups.Enabled) {$buttonRemoveGroups.Enabled = $true} else {$buttonRemoveGroups.Enabled = $false}
        $buttonMoveOU.Enabled = $true
    })

    $buttonReEnable.Add_Click({
        $buttonFindADUser.Enabled = $false

        # Log action
        LogScriptExecution -logPath $logFilePath -action "Re-Enable" -userName $env:USERNAME

        # Get the AD username from the textbox
        $adUsername = $textboxADUsername.Text
        $compName = $textboxADComputer.Text

        if (-not [string]::IsNullOrEmpty($textboxADUsername.Text) -and -not $global:isComputer) {
            ShowReEnableForm
            $global:isComputer = $false
        } 
        elseif (-not [string]::IsNullOrEmpty($textboxADComputer.Text) -and $global:isComputer) {
            # Confirm action with user
            $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to re-enable $($compName)?", "Confirm Re-enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                try {
                    # Get the computer account
                    $compAccount = Get-ADComputer -Identity $compName
                
                    # Enable the computer account
                    Enable-AdAccount -Identity $compAccount
                
                    # Re-retrieve the computer account after enabling
                    $compAccount = Get-ADComputer -Identity $compName
                
                    # Check if the account is enabled
                    if ($compAccount.Enabled) {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline
                        Write-Host "'$($compName)' " -NoNewline
                        Write-Host "has been re-enabled." -ForegroundColor Green
                        
                        # Update statusbar message
                        UpdateStatusBar "Computer '$($compName)' has been re-enabled." -color 'Black'

                        # Display a success dialog box
                        [System.Windows.Forms.MessageBox]::Show("$($compName) has been re-enabled.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                        $buttonFindComputer.Enabled = $true
                        $buttonFindComputer.Focus()

                        # Draw V
                        HideMark $form $adComputerPictureName
                        DrawVmark $form 127 135 "ADComputer"
                    } 
                    else {
                        $dateTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
                        Write-Host "$($dateTime) | " -NoNewline
                        Write-Host "Failed to re-enable " -NoNewline -ForegroundColor Red
                        Write-Host "'$($compName)'."

                        # Update statusbar message
                        UpdateStatusBar "Failed to Re-Enable '$($compName)'." -color 'Red'
                        
                        # Draw X
                        HideMark $form $adComputerPictureName
                        DrawXmark $form 130 140 "ADComputer" 

                        # Display a warning dialog box
                        [System.Windows.Forms.MessageBox]::Show("Failed to re-enable $($compName).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                        
                        $global:isComputer = $false
                        $buttonReEnable.Enabled = $false
                        return $false
                    }

                    # Show a green V checkmark near the textbox
                    HideMark $form $adComputerPictureName
                    DrawVmark $form 127 135 "ADComputer"

                    $buttonFindComputer.Enabled = $false
                    $buttonReEnable.Enabled = $false
                    $buttonMoveOU.Enabled = $true
                    $global:isComputer = $false
                    return $true

                } catch {
                    # Handle exceptions
                    Write-Error "An error occurred: $_"
                    
                    # Display an error dialog box
                    [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    
                    $buttonReEnable.Enabled = $false
                    $global:isComputer = $false
                    return $false
                }
            }
            else {
                # User chose not to proceed
                Write-Warning "Re-enable operation canceled by the user."
                return $false
            }
        }
        else {
            if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)) {
                # Log action
                LogScriptExecution -logPath $logFilePath -action "'$($adUsername)' not found." -userName $env:USERNAME

                # Display error dialog box
                [System.Windows.Forms.MessageBox]::Show("'$($adUsername)' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

                # Disable the FindADUser button
                $buttonReEnable.Enabled = $false
                $buttonFindADUser.Enabled = $false
                $global:isComputer = $false

                return $false
            }
            elseif ([string]::IsNullOrEmpty($compAccount)) {
            # Log action
            LogScriptExecution -logPath $logFilePath -action "'$($compName)' not found." -userName $env:USERNAME

            # Display error dialog box
            [System.Windows.Forms.MessageBox]::Show("'$($compName)' not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

            # Disable the FindADUser button
            $buttonFindADUser.Enabled = $false
            $global:isComputer = $false

            return $false 
            }
        }
    })

    $buttonCopyGroups.Add_Click({
        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Copy Groups" -userName $env:USERNAME

        # Disable the FindADUser button
        $buttonFindADUser.Enabled = $false

        if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)) {ShowCopyGroupsForm}
        elseif (-not [string]::IsNullOrEmpty($textboxCSVFilePath.Text)) {ProcessCSVCopyGroups}

    })

    $buttonRemoveGroups.Add_Click({
        if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)) {
            # Confirm action with user
            $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove '$($textboxADUsername.Text)' from all groups?", "Confirm Group Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                $isRemoved = RemoveGroups $textboxADUsername.Text
                if ($isRemoved) {
                    $buttonRemoveGroups.Enabled = $false
                    [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' has been removed from all groups.", "Remove Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                }

                
            }
        }
        elseif (-not [string]::IsNullOrEmpty($textboxCSVFilePath.Text)) {
            # Process Group Removal for CSV users
            ProccessCSVGroupRemoval
        }
    })

    $buttonMoveOU.Add_Click({
        $buttonFindADUser.Enabled = $false
        $buttonMoveOU.Enabled = $false

        # Log the start of script execution
        LogScriptExecution -logPath $logFilePath -action "Move OU" -userName $env:USERNAME

        $userText = $textboxADUsername.Text
        $computerText = $textboxADComputer.Text
        $csvText = $textboxCSVFilePath.Text

        if (-not [string]::IsNullOrWhiteSpace($userText) -and -not [string]::IsNullOrWhiteSpace($computerText) -and [string]::IsNullOrWhiteSpace($csvText)) {
            HideMark $form "ADUsername"
            HideMark $form "ADComputer"

            # Display Error dialog
            [System.Windows.Forms.MessageBox]::Show("Can't be both User & Computer. Choose one.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)

            $textboxADUsername.Text = ""
            $textboxADComputer.Text = ""
            return $false
        }
        
        if (-not [string]::IsNullOrEmpty($csvText) -or -not [string]::IsNullOrWhiteSpace($csvText)) {
            # Build a form for the CSV operation
            $csvUserPath = $textboxCSVFilePath.Text

            # Log action
            LogScriptExecution -logPath $global:logFilePath -action "Loaded Users CSV: $($csvUserPath)" -userName $env:USERNAME

            # Read the first line of the CSV file to check for the first header
            $firstLine = Get-Content -Path $csvUserPath -TotalCount 1
            if ($firstLine -match 'Username') {
                ShowCSVMoveUserOUForm
                $buttonMoveOU.Enabled = $true
            }
            elseif ($firstLine -match 'ComputerName') {
                ShowCSVMoveComputerOUForm
                $buttonMoveOU.Enabled = $true
            }
        }
        else {
            $userCondition = -not [string]::IsNullOrWhiteSpace($userText) -and -not [string]::IsNullOrEmpty($userText)
            $computerCondition = -not [string]::IsNullOrEmpty($computerText) -or -not [string]::IsNullOrWhiteSpace($computerText)

            if ($userCondition -or $computerCondition) {
                ShowMoveOUForm
                $buttonMoveOU.Enabled = $true
            }
        }

        # Manage buttons
        $buttonFindADUser.Enabled = $false
        $buttonGeneratePassword.Enabled = $true
        $buttonCopyGroups.Enabled = $true
        $buttonMoveOU.Enabled = $true
    })

    # Add controls to form
    $form.Controls.Add($menuStrip)
    $form.Controls.Add($labelADUsername)
    $form.Controls.Add($textboxADUsername)
    $form.Controls.Add($buttonFindADUser)
    $form.Controls.Add($labelADComputerName)
    $form.Controls.Add($textboxADComputer)
    $form.Controls.Add($buttonFindComputer)
    $form.Controls.Add($labelCSVFilePath)
    $form.Controls.Add($textboxCSVFilePath)
    $form.Controls.Add($buttonBrowseForCSVFile)
    $form.Controls.Add($buttonGeneratePassword)
    $form.Controls.Add($buttonResetPassword)
    $form.Controls.Add($buttonReEnable)
    $form.Controls.Add($buttonCopyGroups)
    $form.Controls.Add($buttonRemoveGroups)
    $form.Controls.Add($buttonMoveOU)
    $form.Controls.Add($statusBar)

    $form.Add_FormClosing({
        ManageFormClose
    })

    # Show the form
    [Windows.Forms.Application]::Run($form)
}
# ============== MAIN SECTION =================== #
# Log Action
LogScriptExecution -logPath $logFilePath -action ("= " * 35) -userName $env:USERNAME
LogScriptExecution -logPath $logFilePath -action ("{0,42}" -f "Script Start") -userName $env:USERNAME
LogScriptExecution -logPath $logFilePath -action ("= " * 35) -userName $env:USERNAME

# Run main function
Main
