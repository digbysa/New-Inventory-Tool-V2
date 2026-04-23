@echo off
setlocal EnableExtensions DisableDelayedExpansion

rem =========================
rem Logging default switch:
rem   Set to YES to enable logging by default (double-click)
rem   Set to NO  to disable logging by default
rem =========================
set "DEFAULT_LOG=NO"

set "ENABLE_LOG=0"
if /I "%DEFAULT_LOG%"=="YES" set "ENABLE_LOG=1"

rem Optional overrides (if you run from Command Prompt):
rem   Update-New-Inventory-Tool.bat --log
rem   Update-New-Inventory-Tool.bat --nolog
if /I "%~1"=="--nolog" set "ENABLE_LOG=0"
if /I "%~1"=="/nolog"  set "ENABLE_LOG=0"
if /I "%~1"=="--log"   set "ENABLE_LOG=1"
if /I "%~1"=="/log"    set "ENABLE_LOG=1"

set "LOG=%USERPROFILE%\Desktop\Update-New-Inventory-Tool.log"
set "PS1=%TEMP%\Update-New-Inventory-Tool.ps1"

if "%ENABLE_LOG%"=="1" (
  > "%LOG%" echo ================================
  >>"%LOG%" echo Update started: %DATE% %TIME%
  >>"%LOG%" echo ================================
)

del /q "%PS1%" 2>nul

rem ---- Build PowerShell script ----
>> "%PS1%" echo param([int]$EnableLog = 0)
>> "%PS1%" echo $ErrorActionPreference = "Stop"
>> "%PS1%" echo $ProgressPreference = "SilentlyContinue"
>> "%PS1%" echo
>> "%PS1%" echo $LogPath = Join-Path $env:USERPROFILE "Desktop\Update-New-Inventory-Tool.log"
>> "%PS1%" echo
>> "%PS1%" echo $RepoOwner        = "digbysa"
>> "%PS1%" echo $RepoName         = "New-Inventory-Tool"
>> "%PS1%" echo $Branch           = "main"
>> "%PS1%" echo $TargetFolderName = "New-Inventory-Tool"
>> "%PS1%" echo
>> "%PS1%" echo function Copy-LocationMasterCsvs([string]$srcData, [string]$dstData) ^{
>> "%PS1%" echo     if (-not (Test-Path $srcData)) { return }
>> "%PS1%" echo     New-Item -ItemType Directory -Path $dstData -Force ^| Out-Null
>> "%PS1%" echo     $files = Get-ChildItem -Path $srcData -Recurse -File -Filter "LocationMaster*.csv" -ErrorAction SilentlyContinue
>> "%PS1%" echo     foreach ($f in $files) ^{
>> "%PS1%" echo         $rel  = $f.FullName.Substring($srcData.Length).TrimStart('\')
>> "%PS1%" echo         $dest = Join-Path $dstData $rel
>> "%PS1%" echo         New-Item -ItemType Directory -Path (Split-Path $dest -Parent) -Force ^| Out-Null
>> "%PS1%" echo         Copy-Item -Path $f.FullName -Destination $dest -Force
>> "%PS1%" echo     ^}
>> "%PS1%" echo ^}
>> "%PS1%" echo
>> "%PS1%" echo function Copy-OutputFolder([string]$srcOut, [string]$dstOut) ^{
>> "%PS1%" echo     if (-not (Test-Path $srcOut)) { return }
>> "%PS1%" echo     New-Item -ItemType Directory -Path $dstOut -Force ^| Out-Null
>> "%PS1%" echo     Copy-Item -Path (Join-Path $srcOut "*") -Destination $dstOut -Recurse -Force -ErrorAction SilentlyContinue
>> "%PS1%" echo ^}
>> "%PS1%" echo
>> "%PS1%" echo try ^{
>> "%PS1%" echo     if ($EnableLog -ne 0) ^{
>> "%PS1%" echo         try { Stop-Transcript ^| Out-Null } catch {}
>> "%PS1%" echo         try { Start-Transcript -Path $LogPath -Append ^| Out-Null } catch {}
>> "%PS1%" echo     ^}
>> "%PS1%" echo
>> "%PS1%" echo     try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
>> "%PS1%" echo
>> "%PS1%" echo     $Desktop    = [Environment]::GetFolderPath("Desktop")
>> "%PS1%" echo     $TargetPath = Join-Path $Desktop $TargetFolderName
>> "%PS1%" echo     $zipUrl     = "https://github.com/$RepoOwner/$RepoName/archive/refs/heads/$Branch.zip"
>> "%PS1%" echo
>> "%PS1%" echo     $tempRoot   = Join-Path $env:TEMP ("gh_update_" + [guid]::NewGuid().ToString("N"))
>> "%PS1%" echo     $zipPath    = Join-Path $tempRoot "repo.zip"
>> "%PS1%" echo     $expandPath = Join-Path $tempRoot "expanded"
>> "%PS1%" echo
>> "%PS1%" echo     $backupPath   = $null
>> "%PS1%" echo     $preserveRoot = $null
>> "%PS1%" echo
>> "%PS1%" echo     if (Test-Path $TargetPath) ^{
>> "%PS1%" echo         Write-Host ""
>> "%PS1%" echo         Write-Host "Existing folder detected:" -ForegroundColor Yellow
>> "%PS1%" echo         Write-Host "  $TargetPath"
>> "%PS1%" echo
>> "%PS1%" echo         do ^{
>> "%PS1%" echo             $resp = Read-Host "Backup existing version before updating? (Y/N) [Y]"
>> "%PS1%" echo             if ([string]::IsNullOrWhiteSpace($resp)) { $resp = "Y" }
>> "%PS1%" echo             $resp = $resp.Trim().ToUpper()
>> "%PS1%" echo         ^} until ($resp -eq "Y" -or $resp -eq "N")
>> "%PS1%" echo
>> "%PS1%" echo         if ($resp -eq "Y") ^{
>> "%PS1%" echo             $stamp      = Get-Date -Format "yyyyMMdd_HHmmss"
>> "%PS1%" echo             $backupPath = "${TargetPath}_backup_$stamp"
>> "%PS1%" echo             Write-Host "Backing up existing folder to: $backupPath"
>> "%PS1%" echo             Rename-Item -Path $TargetPath -NewName (Split-Path $backupPath -Leaf)
>> "%PS1%" echo         ^} else ^{
>> "%PS1%" echo             $preserveRoot = Join-Path $env:TEMP ("nit_preserve_" + [guid]::NewGuid().ToString("N"))
>> "%PS1%" echo             New-Item -ItemType Directory -Path $preserveRoot -Force ^| Out-Null
>> "%PS1%" echo
>> "%PS1%" echo             Write-Host "No backup chosen. Preserving user data files temporarily..."
>> "%PS1%" echo             Copy-LocationMasterCsvs (Join-Path $TargetPath "Data")    (Join-Path $preserveRoot "Data")
>> "%PS1%" echo             Copy-OutputFolder        (Join-Path $TargetPath "Output") (Join-Path $preserveRoot "Output")
>> "%PS1%" echo
>> "%PS1%" echo             Write-Host "Removing existing folder (no backup)..."
>> "%PS1%" echo             Remove-Item -Path $TargetPath -Recurse -Force
>> "%PS1%" echo         ^}
>> "%PS1%" echo     ^}
>> "%PS1%" echo
>> "%PS1%" echo     New-Item -ItemType Directory -Path $tempRoot   ^| Out-Null
>> "%PS1%" echo     New-Item -ItemType Directory -Path $expandPath ^| Out-Null
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host ""
>> "%PS1%" echo     Write-Host "Downloading: $zipUrl"
>> "%PS1%" echo     Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host "Extracting..."
>> "%PS1%" echo     Expand-Archive -Path $zipPath -DestinationPath $expandPath -Force
>> "%PS1%" echo
>> "%PS1%" echo     $rootFolder = Get-ChildItem -Path $expandPath -Directory ^| Select-Object -First 1
>> "%PS1%" echo     if (-not $rootFolder) { throw "Extraction failed (no folder found after unzip)." }
>> "%PS1%" echo     $sourcePath = $rootFolder.FullName
>> "%PS1%" echo
>> "%PS1%" echo     New-Item -ItemType Directory -Path $TargetPath -Force ^| Out-Null
>> "%PS1%" echo     Write-Host "Copying files to: $TargetPath"
>> "%PS1%" echo     Copy-Item -Path (Join-Path $sourcePath "*") -Destination $TargetPath -Recurse -Force
>> "%PS1%" echo
>> "%PS1%" echo     $restoreFrom = $null
>> "%PS1%" echo     if ($backupPath)   { $restoreFrom = $backupPath }
>> "%PS1%" echo     if ($preserveRoot) { $restoreFrom = $preserveRoot }
>> "%PS1%" echo
>> "%PS1%" echo     if ($restoreFrom) ^{
>> "%PS1%" echo         Write-Host ""
>> "%PS1%" echo         Write-Host "Restoring user data from: $restoreFrom"
>> "%PS1%" echo         Copy-LocationMasterCsvs (Join-Path $restoreFrom "Data")    (Join-Path $TargetPath "Data")
>> "%PS1%" echo         Copy-OutputFolder        (Join-Path $restoreFrom "Output") (Join-Path $TargetPath "Output")
>> "%PS1%" echo     ^}
>> "%PS1%" echo
>> "%PS1%" echo     $toolBat = Join-Path $TargetPath "NewAssetTool.bat"
>> "%PS1%" echo     if (Test-Path $toolBat) ^{
>> "%PS1%" echo         $shortcutPath = Join-Path $Desktop "NewAssetTool.lnk"
>> "%PS1%" echo         $iconPath     = Join-Path $TargetPath "icon.ico"
>> "%PS1%" echo
>> "%PS1%" echo         if (Test-Path $shortcutPath) { Remove-Item $shortcutPath -Force -ErrorAction SilentlyContinue }
>> "%PS1%" echo
>> "%PS1%" echo         $wsh = New-Object -ComObject WScript.Shell
>> "%PS1%" echo         $sc  = $wsh.CreateShortcut($shortcutPath)
>> "%PS1%" echo         $sc.TargetPath = $toolBat
>> "%PS1%" echo         $sc.WorkingDirectory = $TargetPath
>> "%PS1%" echo         if (Test-Path $iconPath) { $sc.IconLocation = "$iconPath,0" }
>> "%PS1%" echo         $sc.Save()
>> "%PS1%" echo         Write-Host "Shortcut created: $shortcutPath"
>> "%PS1%" echo     ^} else ^{
>> "%PS1%" echo         Write-Host "Note: NewAssetTool.bat not found (shortcut not created)." -ForegroundColor Yellow
>> "%PS1%" echo     ^}
>> "%PS1%" echo
>> "%PS1%" echo     Remove-Item -Path $tempRoot -Recurse -Force
>> "%PS1%" echo     if ($preserveRoot -and (Test-Path $preserveRoot)) { Remove-Item -Path $preserveRoot -Recurse -Force }
>> "%PS1%" echo
>> "%PS1%" echo     Write-Host ""
>> "%PS1%" echo     Write-Host "Done! Updated New-Inventory-Tool on your Desktop." -ForegroundColor Green
>> "%PS1%" echo ^}
>> "%PS1%" echo catch ^{
>> "%PS1%" echo     Write-Host ""
>> "%PS1%" echo     Write-Host "UPDATE FAILED: $($_.Exception.Message)" -ForegroundColor Red
>> "%PS1%" echo     Write-Host "Location: $($_.InvocationInfo.PositionMessage)" -ForegroundColor Yellow
>> "%PS1%" echo     Write-Host "Stack: $($_.ScriptStackTrace)" -ForegroundColor DarkYellow
>> "%PS1%" echo     exit 1
>> "%PS1%" echo ^}
>> "%PS1%" echo finally ^{
>> "%PS1%" echo     if ($EnableLog -ne 0) { try { Stop-Transcript ^| Out-Null } catch {} }
>> "%PS1%" echo ^}

rem ---- Run PowerShell (logging on/off) ----
set "PS_ENABLE_LOG=0"
if "%ENABLE_LOG%"=="1" set "PS_ENABLE_LOG=1"

if "%ENABLE_LOG%"=="1" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -EnableLog %PS_ENABLE_LOG% >> "%LOG%" 2>&1
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -EnableLog %PS_ENABLE_LOG%
)

set "RC=%ERRORLEVEL%"

if "%ENABLE_LOG%"=="1" (
  echo.
  echo =======================
  echo Update log (saved here):
  echo %LOG%
  echo =======================
  echo.
  if not "%RC%"=="0" (
    echo Opening the log in Notepad...
    start notepad "%LOG%"
  )
) else (
  echo.
  echo (Logging is OFF by default.)
  echo To enable it, edit this file and change:
  echo   set "DEFAULT_LOG=NO"
  echo to:
  echo   set "DEFAULT_LOG=YES"
)

del /q "%PS1%" 2>nul
echo.
pause
exit /b %RC%
