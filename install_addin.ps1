# Sideload Word Add-in Script

$manifestFile = "C:\Users\dinhn\OneDrive\Documents\minh-add-in\manifest.xml"
$addInId = "f47ac10b-58cc-4372-a567-0e02b2c3d479"

Write-Host "1. Registering Add-in in WEF Developer registry..." -ForegroundColor Cyan
$wefDevKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $wefDevKey)) {
    New-Item -Path $wefDevKey -Force | Out-Null
}
New-ItemProperty -Path $wefDevKey -Name $addInId -Value $manifestFile -PropertyType String -Force | Out-Null
Write-Host "-> WEF Developer key registered successfully." -ForegroundColor Green

Write-Host "2. Configuring Shared Folder Catalog..." -ForegroundColor Cyan
$shareName = "minh_addin"
$folderPath = "C:\Users\dinhn\OneDrive\Documents\minh-add-in"

try {
    if (-not (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue)) {
        New-SmbShare -Name $shareName -Path $folderPath -FullAccess "$env:USERDOMAIN\$env:USERNAME" -ErrorAction SilentlyContinue | Out-Null
    }
    $uncPath = "\\$env:COMPUTERNAME\$shareName"
    
    $catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{BC7F6DDB-83AD-414E-8400-07A1FD7CA7D8}"
    if (-not (Test-Path $catalogKey)) {
        New-Item -Path $catalogKey -Force | Out-Null
    }
    New-ItemProperty -Path $catalogKey -Name "Id" -Value "{BC7F6DDB-83AD-414E-8400-07A1FD7CA7D8}" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $catalogKey -Name "Url" -Value $uncPath -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $catalogKey -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Host "-> Shared Catalog registered at: $uncPath" -ForegroundColor Green
} catch {
    Write-Warning "Could not auto-create SMB share (may require Admin). WEF Developer sideloading is already active."
}

Write-Host "3. Clearing Office WEF cache..." -ForegroundColor Cyan
$wefCache = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
if (Test-Path $wefCache) {
    Get-ChildItem -Path $wefCache -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "-> WEF cache cleared." -ForegroundColor Green
}

Write-Host "`nAll setup complete! Please open or restart Microsoft Word." -ForegroundColor Green
