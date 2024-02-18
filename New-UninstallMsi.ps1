# Function to uninstall application using product code
function Uninstall-Product {
    param (
        [string]$applicationName,
        [string]$applicationVersion
    )
    try {
        # Check if msiexec.exe is available
        if (-not (WaitForMsiExec)) {
            Write-Host "Error: Timed out waiting for msiexec.exe. Exiting..."
            exit 1
        }
        
        # Find product code based on application name and version
        $productCode = Find-ProductCode -ApplicationName $applicationName -ApplicationVersion $applicationVersion
        if (-not $productCode) {
            Write-Host "Product code not found for application '$applicationName' version '$applicationVersion'."
            return
        }
        
        Write-Host "Uninstalling application '$applicationName' version '$applicationVersion' with code: $productCode"
        $logFilePath = "C:\Temp\$applicationName.log"
        $arguments = "/x $productCode /qn /l*v `"$logFilePath`""
        $process = Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-Host "Application '$applicationName' version '$applicationVersion' uninstalled successfully. Log saved to: $logFilePath"
        } else {
            Write-Host "Error uninstalling application '$applicationName' version '$applicationVersion' with code: $productCode (Exit Code: $($process.ExitCode))"
            throw "Uninstallation error"
        }
    } catch {
        Write-Host "Error uninstalling application '$applicationName' version '$applicationVersion': $_"
    }
}

# Function to find product code based on application name and version
function Find-ProductCode {
    param (
        [string]$applicationName,
        [string]$applicationVersion
    )
    try {
        $uninstallKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction Stop
        foreach ($key in $uninstallKeys) {
            $displayName = $key.GetValue("DisplayName")
            $displayVersion = $key.GetValue("DisplayVersion")
            if ($displayName -like "*$applicationName*" -and $displayVersion -eq $applicationVersion) {
                return $key.GetValue("PSChildName")
            }
        }
        return $null
    } catch {
        Write-Host "Error finding product code for application '$applicationName' version '$applicationVersion': $_"
        return $null
    }
}

# Uninstall application by name and version
$applicationName = "YourApplicationName"
$applicationVersion = "YourApplicationVersion"

Uninstall-Product -applicationName $applicationName -applicationVersion $applicationVersion