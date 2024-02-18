# Function to check if msiexec.exe is currently running as a mutex and wait until it becomes available
function WaitForMsiExec {
    $maxAttempts = 30  # Maximum number of attempts (adjust as needed)
    $sleepInterval = 2 # Sleep interval between attempts (adjust as needed)
    $attempts = 0
    
    while ($attempts -lt $maxAttempts) {
        $msiExecProcess = Get-Process -Name msiexec -ErrorAction SilentlyContinue
        if (-not $msiExecProcess) {
            Write-Host "msiexec.exe is not currently running."
            return $true
        } elseif ($msiExecProcess.WaitForExit(1000)) {
            Write-Host "msiexec.exe is available for use."
            return $true
        } else {
            Write-Host "Waiting for msiexec.exe to become available... (Attempt $($attempts + 1) of $maxAttempts)"
            Start-Sleep -Seconds $sleepInterval
            $attempts++
        }
    }
    
    Write-Host "Timed out waiting for msiexec.exe to become available."
    return $false
}

# Function to uninstall application using product code
function Uninstall-Product {
    param (
        [string]$productCode,
        [string]$productName
    )
    try {
        # Check if msiexec.exe is available
        if (-not (WaitForMsiExec)) {
            Write-Host "Error: Timed out waiting for msiexec.exe. Exiting..."
            exit 1
        }
        
        Write-Host "Uninstalling product '$productName' with code: $productCode"
        $logFilePath = "C:\Temp\$productName.log"
        $arguments = "/x $productCode /qn /l*v `"$logFilePath`""
        $process = Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-Host "Product '$productName' with code $productCode uninstalled successfully. Log saved to: $logFilePath"
        } else {
            Write-Host "Error uninstalling product '$productName' with code: $productCode (Exit Code: $($process.ExitCode))"
        }
    } catch {
        Write-Host "Error uninstalling product '$productName' with code: $productCode - $_"
    }
}

# Function to find product code based on display name and uninstall application if found
function Find-And-Uninstall-Product {
    param (
        [string]$displayName
    )
    try {
        $uninstallKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction Stop
        foreach ($key in $uninstallKeys) {
            $key.GetValue("DisplayName") | Where-Object { $_ -like $displayName } | ForEach-Object {
                $uninstallString = $key.GetValue("UninstallString")
                $match = [regex]::Match($uninstallString, '{[A-F0-9-]+}')
                if ($match.Success) {
                    $productCode = $match.Value
                    $productName = $key.GetValue("DisplayName")
                    Uninstall-Product -productCode $productCode -productName $productName
                }
            }
        }
    } catch {
        Write-Host "Error finding and uninstalling product for $displayName: $_"
    }
}

# Main loop to find and uninstall product codes
$displayNames = @(
    "Application 1*",
    "Application 2*",
    "Application 3*"
)

# Define the name of the mutex for msiexec.exe
$mutexName = "Global\MSIExecMutex"

# Create or open the mutex
$mutex = New-Object Threading.Mutex($true, $mutexName, [ref]$mutexCreated)

# Check if the mutex was successfully created
if (-not $mutexCreated) {
    Write-Host "Mutex already exists. Waiting for release..."
}

# Attempt to acquire the mutex
$mutexAcquired = $mutex.WaitOne()

if ($mutexAcquired) {
    Write-Host "Mutex acquired. Performing uninstallation..."
    
    # Perform uninstallation for each display name
    foreach ($displayName in $displayNames) {
        Find-And-Uninstall-Product $displayName
    }
    
    # Release the mutex when the operation is complete
    $mutex.ReleaseMutex()
    Write-Host "Mutex released."
} else {
    Write-Host "Failed to acquire mutex. Another process may be using it."
}

# Dispose of the mutex object
$mutex.Dispose()