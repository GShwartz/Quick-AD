

# Function to create a lock file
function AcquireLock {
    # Define the lock file path
    $lockFileName = "neuralLock.lock"
    $lockFolderPath = "C:\Temp"
    $lockFilePath = Join-Path -Path $lockFolderPath -ChildPath $lockFileName

    # Create C:\Temp if it doesn't exist
    if (-not (Test-Path -Path $lockFolderPath -PathType Container)) {
        New-Item -ItemType Directory -Path $lockFolderPath | Out-Null
    }

    # Check if the lock file exists
    if (-not (Test-Path $lockFilePath)) {
        # If the lock file doesn't exist, create it
        $lockInfo = @{
            'ComputerName' = $env:COMPUTERNAME
            'Username' = $env:USERNAME
            'IPAddress' = (Test-Connection -ComputerName (hostname) -Count 1).IPV4Address.IPAddressToString
            'Timestamp' = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }

        $lockInfo | ConvertTo-Json | Set-Content -Path $lockFilePath

        return $true
    }

    # The lock file already exists, indicating that the lock is held by another process
    Write-Warning "Lock is active."
    $lockInfo = Get-Content -Raw -Path $lockFilePath | ConvertFrom-Json

    Write-Host "Lock is held by:"
    Write-Host "  ComputerName: $($lockInfo.ComputerName)"
    Write-Host "  Username: $($lockInfo.Username)"
    Write-Host "  IP Address: $($lockInfo.IPAddress)"
    Write-Host "  Timestamp: $($lockInfo.Timestamp)"

    return $false
}

# Function to release the lock file
function ReleaseLock {
    # Define the lock file path
    $lockFileName = "neuralLock.lock"
    $lockFolderPath = "C:\Temp"
    $lockFilePath = Join-Path -Path $lockFolderPath -ChildPath $lockFileName
    
    # Remove the lock file
    Remove-Item -Path $lockFilePath -ErrorAction SilentlyContinue
}
