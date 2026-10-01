param(
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA "Programs\DosyaTipCevirici"),
    [string]$DesktopDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
)

$ErrorActionPreference = "Stop"

$installedExe = Join-Path $InstallDirectory "DosyaTipCevirici.exe"
$shortcutPath = Join-Path $DesktopDirectory "Dosya Tip Dönüştürücü.lnk"

if (Test-Path -LiteralPath $shortcutPath -PathType Leaf) {
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    if ($shortcut.TargetPath -eq $installedExe) {
        Remove-Item -LiteralPath $shortcutPath -Force
    }
}

if (Test-Path -LiteralPath $installedExe -PathType Leaf) {
    Remove-Item -LiteralPath $installedExe -Force
}
if ((Test-Path -LiteralPath $InstallDirectory -PathType Container) -and
    (Get-ChildItem -LiteralPath $InstallDirectory -Force | Measure-Object).Count -eq 0) {
    Remove-Item -LiteralPath $InstallDirectory
}

Write-Host "Dosya Tip Dönüştürücü kaldırıldı."
