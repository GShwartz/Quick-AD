# 1. Finish the tooltips
# 2. Add settings with option to change the tooltip popup time.
# 3. Add statusbar updates


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
    "Logger",           # Handles logging
    "Lock",             # Provides functions for locking and unlocking resources
    "Visuals",          # Handles visual components and UI elements
    "ReEnableForm",     # Manages re-enabling functionality
    "CopyGroups",       # Handles copying groups
    "moveOUForm",       # Manages the Move OU forms
    "CsvHandler",       # Handles CSV file operations
    "ADHandler",        # Handles Active Directory operations
    "BuildForm",        # Builds and configures forms
    "MenuStrip"         # Builds the top menu
)

foreach ($moduleName in $modulePaths) {
    $modulePath = Join-Path $scriptDirectory "Modules\$moduleName.psm1"
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

# ============== MAIN SECTION =================== #
# Log Action
LogScriptExecution -logPath $logFilePath -action ("= " * 35) -userName $env:USERNAME
LogScriptExecution -logPath $logFilePath -action ("{0,42}" -f "Script Start") -userName $env:USERNAME
LogScriptExecution -logPath $logFilePath -action ("= " * 35) -userName $env:USERNAME

# Define vars
$global:version = "1.0.0"
$global:primaryUser = $null 
$global:primaryComputer = $null
$global:isComputer = $false
$global:isCSV = $false
$global:statusBarMessage = $null

$adUserNamePictureName = "ADUsername"
$adComputerPictureName = "ADComputer"
$csvPathPictureName = "CSVPath"
$prefixes = @($adUserNamePictureName, $adComputerPictureName, $csvPathPictureName)

# Create the main form
$global:form = CreateCanvas "QUICK-AD" -x 470 -y 380
$form.MaximizeBox = $false

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

# ================ TOOLTIPS =================== #
# ________====== Find AD User ======_________ #
# Create the buttonFindADUser tooltip
$buttonFindADUserTooltip = New-Object System.Windows.Forms.ToolTip
$buttonFindADUserTooltip.SetToolTip($buttonFindADUser, "Search for an Active Directory User account")

# Create a timer for the find AD User tooltip delay
$findADUserTooltipTimer = New-Object System.Windows.Forms.Timer
$findADUserTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event on buttonFindADUser
$buttonFindADUser.Add_MouseEnter({
    # Start the timer when the mouse enters the button
    $findADUserTooltipTimer.Start()
})

# Event handler for the MouseLeave event on buttonFindADUser
$buttonFindADUser.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $findADUserTooltipTimer.Stop()
    $buttonFindADUserTooltip.Hide($buttonFindADUser)
})

# Event handler for the Tick event on the timer
$findADUserTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $findADUserTooltipTimer.Stop()
})

# ________====== Find AD Computer ======_________ #
# Create the buttonFindComputer tooltip
$buttonFindComputerTooltip = New-Object System.Windows.Forms.ToolTip
$buttonFindComputerTooltip.SetToolTip($buttonFindComputer, "Search for an Active Directory Computer account")

# Create a timer for the find AD Computer tooltip delay
$findADComputerTooltipTimer = New-Object System.Windows.Forms.Timer
$findADComputerTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonFindComputer.Add_MouseEnter({
    # Start the timer when the mouse enters the button
    $findADComputerTooltipTimer.Start()
})

# Event handler for the MouseLeave event
$buttonFindComputer.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $findADComputerTooltipTimer.Stop()
    $buttonFindComputerTooltip.Hide($buttonFindComputer)
})

# Event handler for the Tick event on the timer
$findADComputerTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $findADComputerTooltipTimer.Stop()
})

# ________====== Browse ======_________ #
# Create the buttonBrowseCSV tooltip
$buttonBrowseTooltip = New-Object System.Windows.Forms.ToolTip
$buttonBrowseTooltip.SetToolTip($buttonBrowseForCSVFile, "Browse for a CSV file with the 'Username' or 'ComputerName' headers")

# Create a timer for the find AD Computer tooltip delay
$browseTooltipTimer = New-Object System.Windows.Forms.Timer
$browseTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonBrowseForCSVFile.Add_MouseEnter({
    # Start the timer when the mouse enters the button
    $browseTooltipTimer.Start()
})

# Event handler for the MouseLeave event
$buttonBrowseForCSVFile.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $browseTooltipTimer.Stop()
    $buttonBrowseTooltip.Hide($buttonBrowseForCSVFile)
})

# Event handler for the Tick event on the timer
$browseTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $browseTooltipTimer.Stop()
})

# ________====== Generate Password ======_________ #
# Create the buttonGeneratePassword tooltip
$buttonGeneratePasswordTooltip = New-Object System.Windows.Forms.ToolTip

# Create a timer for the find AD Computer tooltip delay
$generatePasswordTooltipTimer = New-Object System.Windows.Forms.Timer
$generatePasswordTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonGeneratePassword.Add_MouseEnter({
    if ([string]::IsNullOrEmpty($textboxCSVFilePath.Text)){
        $buttonGeneratePasswordTooltip.SetToolTip($buttonGeneratePassword, "Generate password for '$($textboxADUsername.Text)'")
    }
    else {
        $buttonGeneratePasswordTooltip.SetToolTip($buttonGeneratePassword, "Generate passwords for CSV file")
    }

    # Start the timer when the mouse enters the button
    $generatePasswordTooltipTimer.Start()
})

# Event handler for the MouseLeave event
$buttonGeneratePassword.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $generatePasswordTooltipTimer.Stop()
    $buttonGeneratePasswordTooltip.Hide($buttonGeneratePassword)
})

# Event handler for the Tick event on the timer
$generatePasswordTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $generatePasswordTooltipTimer.Stop()
})

# ________====== Reset Password ======_________ #
# Create the buttonResetPassword tooltip
$buttonResetPasswordTooltip = New-Object System.Windows.Forms.ToolTip

# Create a timer for the find AD Computer tooltip delay
$resetPasswordTooltipTimer = New-Object System.Windows.Forms.Timer
$resetPasswordTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonResetPassword.Add_MouseEnter({
    if ([string]::IsNullOrEmpty($textboxCSVFilePath.Text)){
        $buttonGeneratePasswordTooltip.SetToolTip($buttonResetPassword, "Reset password for '$($textboxADUsername.Text)'")
    }
    else {
        $buttonGeneratePasswordTooltip.SetToolTip($buttonResetPassword, "Reset passwords for CSV file")
    }

    # Start the timer when the mouse enters the button
    $resetPasswordTooltipTimer.Start()
})

# Event handler for the MouseLeave event
$buttonResetPassword.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $resetPasswordTooltipTimer.Stop()
    $buttonResetPasswordTooltip.Hide($buttonResetPassword)
})

# Event handler for the Tick event on the timer
$resetPasswordTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $resetPasswordTooltipTimer.Stop()
})

# ________====== Re-Enable ======_________ #
# Create the Re-Enable tooltip
$buttonReEnableTooltip = New-Object System.Windows.Forms.ToolTip

# Create a timer for the find AD Computer tooltip delay
$reEnableTooltipTimer = New-Object System.Windows.Forms.Timer
$reEnableTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonReEnable.Add_MouseEnter({
    if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)){
        $buttonReEnableTooltip.SetToolTip($buttonReEnable, "Re-Enable User '$($textboxADUsername.Text)' with an example AD user account")
    }
    else {
        $buttonReEnableTooltip.SetToolTip($buttonReEnable, "Re-Enable Computer '$($textboxADComputer.Text)' with an example AD computer account")
    }

    # Start the timer when the mouse enters the button
    $reEnableTooltipTimer.Start()
})

# Event handler for the MouseLeave event
$buttonReEnable.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $reEnableTooltipTimer.Stop()
    $buttonReEnableTooltip.Hide($buttonReEnable)
})

# Event handler for the Tick event on the timer
$reEnableTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $reEnableTooltipTimer.Stop()
})

# ________====== Copy Groups ======_________ #
# Create the Copy Groups tooltip
$buttonCopyGroupsTooltip = New-Object System.Windows.Forms.ToolTip
$buttonCopyGroupsTooltip.SetToolTip($buttonCopyGroups, "Copy groups from an example AD user account")

# Create a timer for the copy groups tooltip delay
$copyGroupsTooltipTimer = New-Object System.Windows.Forms.Timer
$copyGroupsTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonCopyGroups.Add_MouseEnter({
    # Start the timer when the mouse enters the button
    $copyGroupsTooltipTimer.Start()
})

# Event handler for the MouseLeave event 
$buttonCopyGroups.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $copyGroupsTooltipTimer.Stop()
    $buttonCopyGroupsTooltip.Hide($buttonCopyGroups)
})

# Event handler for the Tick event on the timer
$copyGroupsTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $copyGroupsTooltipTimer.Stop()
})

# ________====== Remove Groups ======_________ #
# Create the Remove Groups tooltip
$buttonRemoveGroupsTooltip = New-Object System.Windows.Forms.ToolTip

# Create a timer for the remove groups tooltip delay
$removeGroupsTooltipTimer = New-Object System.Windows.Forms.Timer
$removeGroupsTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonRemoveGroups.Add_MouseEnter({
    $buttonRemoveGroupsTooltip.SetToolTip($buttonRemoveGroups, "Remove '$($textboxADUsername.Text)' group memberships")

    # Start the timer when the mouse enters the button
    $removeGroupsTooltipTimer.Start()
})

# Event handler for the MouseLeave event 
$buttonRemoveGroups.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $removeGroupsTooltipTimer.Stop()
    $buttonRemoveGroupsTooltip.Hide($buttonRemoveGroups)
})

# Event handler for the Tick event on the timer
$removeGroupsTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $removeGroupsTooltipTimer.Stop()
})

# ________====== Move OU ======_________ #
# Create the Move OU tooltip
$buttonMoveOUTooltip = New-Object System.Windows.Forms.ToolTip

# Create a timer for the remove groups tooltip delay
$moveOUTooltipTimer = New-Object System.Windows.Forms.Timer
$moveOUTooltipTimer.Interval = 2000  # 2000 milliseconds (2 seconds)

# Event handler for the MouseEnter event
$buttonMoveOU.Add_MouseEnter({
    if (-not [string]::IsNullOrEmpty($textboxADUsername.Text)) {
        $buttonMoveOUTooltip.SetToolTip($buttonMoveOU, "Relocate the user '$($textboxADUsername.Text)' to OU by example user account or bulk location")
    }
    elseif (-not [string]::IsNullOrEmpty($textboxADComputer.Text)) {
        $buttonMoveOUTooltip.SetToolTip($buttonMoveOU, "Relocate the computer '$($textboxADComputer.Text)' to OU by example computer account")
    }
    elseif (-not [string]::IsNullOrEmpty($textboxCSVFilePath.Text)) {
        $buttonMoveOUTooltip.SetToolTip($buttonMoveOU, "Relocate CSV file data to OU by an AD example computer or user account")
    }
    else {
        # do nothing
    }

    # Start the timer when the mouse enters the button
    $moveOUTooltipTimer.Start()
})

# Event handler for the MouseLeave event 
$buttonMoveOU.Add_MouseLeave({
    # Stop the timer when the mouse leaves the button
    $moveOUTooltipTimer.Stop()
    $buttonMoveOUTooltip.Hide($buttonMoveOU)
})

# Event handler for the Tick event on the timer
$moveOUTooltipTimer.Add_Tick({
    # Stop the timer after displaying the tooltip
    $moveOUTooltipTimer.Stop()
})

# ============== EVENT HANDLERS =================== #
# Event handler for mouse down on AD Username textbox
$textboxADUsername.Add_MouseDown({
    $textboxADComputer.Text = ""
    $textboxCSVFilePath.Text = ""

    if ([string]::IsNullOrEmpty($textboxADUsername.Text)) {
        HideAllMarks $form $prefixes
    }
    else {
        HideMark $form $adComputerPictureName
        HideMark $form $csvPathPictureName
    }

    $buttonFindADUser.Enabled = $false
    $buttonFindComputer.Enabled = $false
    $buttonGeneratePassword.Enabled = $false
    $buttonCopyGroups.Enabled = $false
    $buttonRemoveGroups.Enabled = $false
    $buttonReEnable.Enabled = $false
})

# Event handler for mouse down on AD Computer textbox
$textboxADComputer.Add_MouseDown({
    HideMark $form $adUserNamePictureName
    HideMark $form $csvPathPictureName

    $textboxADUsername.Text = ""
    $textboxCSVFilePath.Text = ""

    $buttonFindADUser.Enabled = $false
    $buttonFindComputer.Enabled = $false
    $buttonGeneratePassword.Enabled = $false
    $buttonCopyGroups.Enabled = $false
    $buttonRemoveGroups.Enabled = $false
    $buttonReEnable.Enabled = $false
})

# = = = = = = = EVENT HANDLERS = = = = = = = #
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
    $buttonCopyGroups.Enabled = $false
    $buttonReEnable.Enabled = $false
    $buttonResetPassword.Enabled = $false
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
        GeneratePassword $global:primaryUser.SamAccountName

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
                    Write-Host "$($compName) has been re-enabled." -ForegroundColor Green
                    
                    # Display a success dialog box
                    [System.Windows.Forms.MessageBox]::Show("$($compName) has been re-enabled.", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

                    $buttonFindComputer.Enabled = $true
                    $buttonFindComputer.Focus()

                    HideMark $form $adComputerPictureName
                } 
                
                else {
                    Write-Warning "Failed to re-enable $($compName)."
                    
                    # Display a warning dialog box
                    [System.Windows.Forms.MessageBox]::Show("Failed to re-enable $($compName).", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                    
                    $global:isComputer = $false
                    $buttonReEnable.Enabled = $false
                    return $false
                }

                # Show a green V checkmark near the textbox
                HideMark $form $adComputerPictureName

                $buttonFindComputer.Enabled = $true
                $buttonReEnable.Enabled = $false
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

    # Show Copy Groups form
    ShowCopyGroupsForm
})

$buttonRemoveGroups.Add_Click({
    # Confirm action with user
    $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to remove '$($textboxADUsername.Text)' from all groups?", "Confirm Group Removal", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $isRemoved = RemoveGroups $textboxADUsername.Text
        if ($isRemoved) {
            $buttonRemoveGroups.Enabled = $false
            [System.Windows.Forms.MessageBox]::Show("User '$($global:primaryUser.SamAccountName)' has been removed from all groups.", "Remove Groups", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
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

    # Disable buttons
    $buttonFindADUser.Enabled = $false
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

$form.Add_FormClosing({
    ManageFormClose
})

# Show the form
[Windows.Forms.Application]::Run($form)
