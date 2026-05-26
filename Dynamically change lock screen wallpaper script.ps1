# 1. Force 64-bit execution context to avoid Wow6432Node registry redirection issues
if ($env:PROCESSOR_ARCHITECTURE -eq 'x86' -and $env:PROCESSOR_ARCHITEW6432 -eq 'AMD64') {
    & "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe" -File $MyInvocation.MyCommand.Path
    Exit
}

# Parameter definitions
$LocalFolder  = "C:\Windows\Web\Screen"
$SourceFolder = "\\domain.local\netlogon\LockScreens" # <--- Change this to your central network share path

# 2. Select a random image from the network folder
if (Test-Path $SourceFolder) {
    $Images = Get-ChildItem -Path $SourceFolder -Include *.jpg, *.jpeg, *.png -File -Recurse
    
    if ($Images) {
        # Select 1 random image from the filtered list
        $RandomImage = $Images | Get-Random
        
        # Generate a unique local filename each time to bypass the Windows OS cache engine
        $RandomID = Get-Random -Min 1000 -Max 9999
        $LocalName = "CorporateLockScreen_$RandomID.jpg"
        $LocalPath = "$LocalFolder\$LocalName"
        
        if (-not (Test-Path $LocalFolder)) {
            New-Item -ItemType Directory -Path $LocalFolder -Force | Out-Null
        }
        
        # Clean up old images from previous runs to prevent local disk space accumulation
        Remove-Item -Path "$LocalFolder\CorporateLockScreen_*.jpg" -Force -ErrorAction SilentlyContinue
        
        # Copy the newly selected image to the local directory with the unique filename
        Copy-Item -Path $RandomImage.FullName -Destination $LocalPath -Force
        
        # 3. Apply Group Policy Registry Configurations
        $PolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization"
        $CSPPath    = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

        if (-not (Test-Path $PolicyPath)) { New-Item $PolicyPath -Force | Out-Null }
        if (-not (Test-Path $CSPPath))    { New-Item $CSPPath -Force | Out-Null }

        # Enforce policies (NoChangingLockScreen = 1 is mandatory to bypass system lock defaults)
        New-ItemProperty -Path $PolicyPath -Name "LockScreenImage" -Value $LocalPath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $PolicyPath -Name "NoChangingLockScreen" -Value 1 -PropertyType DWord -Force | Out-Null

        # PersonalizationCSP configurations strictly required for Windows 11 24H2 builds
        New-ItemProperty -Path $CSPPath -Name "LockScreenImagePath" -Value $LocalPath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $CSPPath -Name "LockScreenImageUrl" -Value $LocalPath -Type String -Force | Out-Null
        New-ItemProperty -Path $CSPPath -Name "LockScreenImageStatus" -Value 1 -PropertyType DWord -Force | Out-Null

        # 4. Clear the stubborn Windows 11 System Cache directory (SystemData)
        $SystemDataPath = "C:\ProgramData\Microsoft\Windows\SystemData"
        if (Test-Path $SystemDataPath) {
            # Take ownership and grant full access to the Administrators group
            takeown /f $SystemDataPath /r /d y *>$null
            icacls $SystemDataPath /grant "Administrators:(OI)(CI)F" /t /c /q *>$null
            
            # Purge all previously cached lockscreen components from the file storage
            Get-ChildItem -Path $SystemDataPath -Recurse -Filter "LockScreen_*" | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $SystemDataPath -Recurse -Filter "img100*" | Remove-Item -Force -ErrorAction SilentlyContinue
        }

        # 5. NEW: Force LogonUI to refresh to apply the new image instantly without user interaction
        # If the computer is currently on the lock screen, this will force it to instantly redraw with the new background
        $LogonProcess = Get-Process -Name LogonUI -ErrorAction SilentlyContinue
        if ($LogonProcess) {
            Stop-Process -Name LogonUI -Force -ErrorAction SilentlyContinue
        }
        
        Write-Host "Success! Selected [$($RandomImage.Name)] and forced lock screen visual refresh." -ForegroundColor Green
    } else {
        Write-Warning "No images found in the central folder!"
    }
} else {
    Write-Error "Could not connect to the network share path!"
}
