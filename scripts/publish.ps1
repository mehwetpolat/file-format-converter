param([string]$Runtime = "win-x64")

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$project = Join-Path $root "src/DosyaTipCevirici.WinForms/DosyaTipCevirici.WinForms.csproj"
$publishDirectory = Join-Path $root "artifacts/publish/$Runtime"
$archive = Join-Path $root "artifacts/DosyaTipCevirici-$Runtime.zip"

dotnet publish $project -c Release -r $Runtime --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o $publishDirectory
if ($LASTEXITCODE -ne 0) { throw "Yayın paketi oluşturulamadı." }

foreach ($file in @("install.ps1", "uninstall.ps1", "Kur.cmd", "Kaldir.cmd")) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination (Join-Path $publishDirectory $file) -Force
}
Copy-Item -LiteralPath (Join-Path $root "LICENSE") -Destination (Join-Path $publishDirectory "LICENSE.txt") -Force

if (Test-Path $archive) { Remove-Item -LiteralPath $archive }
Compress-Archive -Path (Join-Path $publishDirectory "*") -DestinationPath $archive
Write-Host "Paket hazır: $archive"
