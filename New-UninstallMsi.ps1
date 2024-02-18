# Function to find product code based on application name and version
function Find-ProductCode {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName,
        
        [Parameter(Mandatory = $true)]
        [string]$MaxApplicationVersion
    )
    
    begin {
        # Define regex pattern to extract product code from uninstall string
        $regexPattern = "{[A-F0-9]{8}-([A-F0-9]{4}-){3}[A-F0-9]{12}}"
    }
    
    process {
        try {
            # Get uninstall keys from registry
            $uninstallKeys = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" -ErrorAction Stop
            
            foreach ($key in $uninstallKeys) {
                $displayName = $key.GetValue("DisplayName")
                $displayVersion = $key.GetValue("DisplayVersion")
                
                if ($displayName -like "*$ApplicationName*" -and [version]$displayVersion -lt [version]$MaxApplicationVersion) {
                    # Get the uninstall string
                    $uninstallString = $key.GetValue("UninstallString")
                    
                    # Use regex to extract the product code from the uninstall string
                    if ($uninstallString -match $regexPattern) {
                        return $Matches[0]
                    }
                }
            }
            return $null
        } catch {
            Write-Host "Error finding product code for application '$ApplicationName' with version less than '$MaxApplicationVersion': $_"
            return $null
        }
    }
}