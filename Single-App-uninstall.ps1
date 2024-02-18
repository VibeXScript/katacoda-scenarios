function Uninstall-Application {
    param (
        [string]$displayName,
        [string]$versionThreshold
    )

    begin {
        # Function to find product code from MSIExec uninstall string
        function Find-ProductCode {
            param (
                [string]$uninstallString
            )

            if ($uninstallString -match "/I{[A-F0-9-]+}") {
                $matches[0]
            }
        }

        # Function to search registry for uninstall information
        function Search-Registry {
            param (
                [string]$hive,
                [string]$key
            )

            if (Test-Path $key) {
                Get-ItemProperty -Path $key | Where-Object { $_.DisplayName -eq $displayName -and $_.DisplayVersion -lt $versionThreshold } | Select-Object DisplayName, DisplayVersion, UninstallString
            }
        }
    }

    process {
        try {
            # Search 32-bit registry hive
            $uninstallInfo32 = Search-Registry -hive "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" -key "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"

            # Search 64-bit registry hive
            $uninstallInfo64 = Search-Registry -hive "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -key "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

            if ($uninstallInfo32 -or $uninstallInfo64) {
                foreach ($info in $uninstallInfo32, $uninstallInfo64) {
                    foreach ($app in $info) {
                        $productCode = Find-ProductCode -uninstallString $app.UninstallString
                        if ($productCode) {
                            Write-Host "Uninstalling $($app.DisplayName) $($app.DisplayVersion)..."
                            # Start uninstall process using MSIExec with logging
                            $logPath = "C:\Temp\Uninstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
                            $logArgs = "/x $productCode /quiet /log `"$logPath`""
                            Start-Process msiexec -ArgumentList $logArgs -Wait -ErrorAction Stop
                            Write-Host "Uninstall completed for $($app.DisplayName) $($app.DisplayVersion). Log saved to $logPath."
                        } else {
                            Write-Host "Uninstall string for $($app.DisplayName) $($app.DisplayVersion) is not an MSIExec code. Skipping..."
                        }
                    }
                }
            } else {
                Write-Host "No matching applications found for uninstall."
            }
        } catch {
            Write-Host "Error occurred: $_"
        }
    }
}

# Example usage:
Uninstall-Application -displayName "YourAppDisplayName" -versionThreshold "1.0"