param(
    [string]$SourceDirectory = $PSScriptRoot,
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA "Programs\DosyaTipCevirici"),
    [string]$DesktopDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
)

$ErrorActionPreference = "Stop"

$sourceExe = Join-Path $SourceDirectory "DosyaTipCevirici.exe"
if (-not (Test-Path -LiteralPath $sourceExe -PathType Leaf)) {
    throw "DosyaTipCevirici.exe bulunamadı. Önce ZIP paketini bütünüyle çıkarın."
}

$installedExe = Join-Path $InstallDirectory "DosyaTipCevirici.exe"
$shortcutPath = Join-Path $DesktopDirectory "Dosya Tip Dönüştürücü.lnk"

New-Item -ItemType Directory -Path $InstallDirectory -Force | Out-Null
Copy-Item -LiteralPath $sourceExe -Destination $installedExe -Force

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $installedExe
$shortcut.WorkingDirectory = $InstallDirectory
$shortcut.IconLocation = "$installedExe,0"
$shortcut.Description = "Dosya Tip Dönüştürücü"
$shortcut.Save()

Write-Host "Kurulum tamamlandı. Masaüstü kısayolu: $shortcutPath"
