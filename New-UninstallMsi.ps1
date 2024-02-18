# Function to find product code based on application name and version
function Find-ProductCode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName,
        
        [Parameter(Mandatory = $true)]
        [string]$ApplicationVersion
    )
    
    begin {
        # Define regex pattern to extract product code from uninstall string
        $regexPattern = ".*?{.*?}"
    }
    
    process {
        try {
            # Get uninstall keys from registry
            $uninstallKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction Stop
            
            foreach ($key in $uninstallKeys) {
                $displayName = $key.GetValue("DisplayName")
                $displayVersion = $key.GetValue("DisplayVersion")
                
                if ($displayName -like "*$ApplicationName*" -and $displayVersion -eq $ApplicationVersion) {
                    # Get the uninstall string
                    $uninstallString = $key.GetValue("UninstallString")
                    
                    # Use regex to extract the product code from the uninstall string
                    $productCode = $uninstallString -replace $regexPattern
                    
                    if ($productCode -match "{[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}}") {
                        return $Matches[0]
                    }
                }
            }
            return $null
        } catch {
            Write-Host "Error finding product code for application '$ApplicationName' version '$ApplicationVersion': $_"
            return $null
        }
    }
}

# Function to uninstall application using product code
function Uninstall-Product {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ProductCode,
        
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName,
        
        [Parameter(Mandatory = $true)]
        [string]$ApplicationVersion
    )
    
    begin {
        # Check if msiexec.exe is available
        if (-not (WaitForMsiExec)) {
            Write-Host "Error: Timed out waiting for msiexec.exe. Exiting..."
            exit 1
        }
    }
    
    process {
        try {
            Write-Host "Uninstalling application '$ApplicationName' version '$ApplicationVersion' with code: $ProductCode"
            $logFilePath = "C:\Temp\$ApplicationName.log"
            $arguments = "/x $ProductCode /qn /l*v `"$logFilePath`""
            $process = Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
            if ($process.ExitCode -eq 0) {
                Write-Host "Application '$ApplicationName' version '$ApplicationVersion' uninstalled successfully. Log saved to: $logFilePath"
            } else {
                Write-Host "Error uninstalling application '$ApplicationName' version '$ApplicationVersion' with code: $ProductCode (Exit Code: $($process.ExitCode))"
                throw "Uninstallation error"
            }
        } catch {
            Write-Host "Error uninstalling application '$ApplicationName' version '$ApplicationVersion': $_"
        }
    }
}

# Main script
$ApplicationName = "YourApplicationName"
$ApplicationVersion = "YourApplicationVersion"

try {
    $ProductCode = Find-ProductCode -ApplicationName $ApplicationName -ApplicationVersion $ApplicationVersion
    if ($ProductCode) {
        Uninstall-Product -ProductCode $ProductCode -ApplicationName $ApplicationName -ApplicationVersion $ApplicationVersion
    } else {
        Write-Host "Product code not found for application '$ApplicationName' version '$ApplicationVersion'."
    }
} catch {
    Write-Host "An error occurred: $_"
}