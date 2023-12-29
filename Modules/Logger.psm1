# Function to log script execution details
function LogScriptExecution {
    param (
        [string]$action,
        [string]$userName,
        [string]$logPath
    )

    $logEntry = "{0} | User: {1} | Action: {2}" -f (Get-Date -Format "dd-MM-yyyy HH:mm:ss"), $userName, $action
    Add-Content -Path $logPath -Value $logEntry -Force
}