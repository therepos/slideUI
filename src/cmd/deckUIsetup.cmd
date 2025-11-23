@echo off
setlocal EnableExtensions

:: Sanity: required files must be beside this cmd
if not exist "%~dp0deckUI.ppam" (
  echo Missing deckUI.ppam next to deckUIsetup.cmd
  pause
  exit /b 1
)

:: Extract embedded PowerShell payload
set "PAYTAG=::PAYLOAD"
for /f "delims=:" %%A in ('findstr /n /c:"%PAYTAG%" "%~f0"') do set /a LN=%%A+1
set "TMPPS=%TEMP%\deckUI_setup_%RANDOM%.ps1"
more +%LN% "%~f0" > "%TMPPS%"

:: Pass our folder via ENV and PS param; run STA for Office COM
set "SETUP_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%TMPPS%" -SourceDirOverride "%~dp0"
set "rc=%ERRORLEVEL%"
del "%TMPPS%" >nul 2>&1
if not "%rc%"=="0" (
  echo.
  echo Installer reported an error (code %rc%). See messages above.
  pause
)
endlocal
exit /b %rc%

::PAYLOAD
param([string]$SourceDirOverride)

# ================= SETUP =================
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$Host.UI.RawUI.WindowTitle = 'deckUI setup'
# ========================================

<# 
PowerPoint Add-in Installer/Uninstaller (single file)
- Copies to C:\Apps (creates if missing)
- Adds PowerPoint Trusted Location (HKCU for 16.0/15.0/14.0 if present)
- Loads/Unloads .ppam via PowerPoint COM AddIns
- Clean status messages
#>

# ===== CONFIG =====
$TargetDir      = 'C:\Apps'
$FilesToInstall = @('deckUI.ppam')
$AddInFile      = 'deckUI.ppam'
$TrustedDesc    = 'Installer-Managed: Apps PowerPoint Add-in'
# ==================

# -------- Helpers --------
function Get-ScriptDir {
  if ($env:SETUP_DIR)     { return $env:SETUP_DIR }
  if ($SourceDirOverride) { return $SourceDirOverride }
  if ($PSScriptRoot)      { return $PSScriptRoot }
  if ($PSCommandPath)     { return (Split-Path -Parent $PSCommandPath) }
  if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
  return (Get-Location).Path
}
$SourceDir = (Get-ScriptDir).TrimEnd('\') + '\'

function Ensure-Dir($p){ if(-not(Test-Path $p)){ New-Item -ItemType Directory -Path $p -Force | Out-Null } }
function Get-OfficeVersions { @('16.0','15.0','14.0') | Where-Object { Test-Path "HKCU:\Software\Microsoft\Office\$_" } }

function Status($msg,[scriptblock]$act,[switch]$Fatal){
  $w=34; Write-Host ($msg.PadRight($w)) -NoNewline
  try{ & $act | Out-Null; Write-Host "Done" -ForegroundColor Green }
  catch{
    Write-Host "Failed" -ForegroundColor Red
    Write-Host ("  " + $_.Exception.Message) -ForegroundColor DarkRed
    if($Fatal){ throw }
  }
}

# PowerPoint Trusted Locations
function Add-TrustedLocation($path,$desc){
  $ok=$false
  foreach($ver in Get-OfficeVersions){
    $base="HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Security\Trusted Locations"
    if(-not(Test-Path $base)){ continue }
    $n=1; while(Test-Path "$base\Location$n"){ $n++ }
    $k="$base\Location$n"
    New-Item -Path $k -Force | Out-Null
    New-ItemProperty -Path $k -Name Path -Value ($path.TrimEnd('\')+'\') -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $k -Name AllowSubFolders -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $k -Name Description -Value $desc -PropertyType String -Force | Out-Null
    $ok=$true
  }
  return $ok
}
function Remove-TrustedLocation($path,$desc){
  $removed=$false
  foreach($ver in Get-OfficeVersions){
    $base="HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Security\Trusted Locations"
    if(-not(Test-Path $base)){ continue }
    Get-ChildItem $base | ForEach-Object{
      $p=(Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue)
      if($p.Path -and ($p.Path.TrimEnd('\')+'\') -ieq ($path.TrimEnd('\')+'\') -and ($p.Description -like "$desc*")){
        Remove-Item $_.PsPath -Recurse -Force -ErrorAction SilentlyContinue; $removed=$true
      }
    }
  }
  return $removed
}

# PowerPoint COM
function With-PowerPointCOM([scriptblock]$action){
  $pp=$null
  try{
    $pp=New-Object -ComObject PowerPoint.Application
    $pp.Visible=$false
    & $action $pp | Out-Null
  } finally {
    if($pp){ try{ $pp.Quit() | Out-Null }catch{} }
    if($pp){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Try-LoadAddIn-COM($full){
  try{
    With-PowerPointCOM {
      param($pp)
      $name=[IO.Path]::GetFileName($full)
      $a=$pp.AddIns | Where-Object { $_.FullName -ieq $full -or $_.Name -ieq $name }
      if(-not $a){ $a=$pp.AddIns.Add($full,$true) }
      $a.Loaded=$true
    }; $true
  } catch { $false }
}
function Try-UnloadAddIn-COM($full){
  try{
    With-PowerPointCOM {
      param($pp)
      $name=[IO.Path]::GetFileName($full)
      $pp.AddIns | Where-Object { $_.FullName -ieq $full -or $_.Name -ieq $name } | ForEach-Object { $_.Loaded=$false }
    }; $true
  } catch { $false }
}

# Detection (PowerPoint only)
function Detect-Installed{
  $present=@()

  $local=Join-Path $TargetDir $AddInFile
  if(Test-Path $local){ $present+=$local }

  try{
    With-PowerPointCOM {
      param($pp)
      $pp.AddIns | ForEach-Object {
        if($_.FullName -and (Split-Path $_.FullName -Leaf) -ieq $AddInFile -and (Test-Path $_.FullName)){
          $script:__paths=($script:__paths + $_.FullName)
        }
      }
    }
  } catch {}
  
  ($present + $script:__paths) | Where-Object { $_ } | Select-Object -Unique
}

# Actions
function Install-Addin{
  Status "Installing files" -Fatal {
    Ensure-Dir $TargetDir
    foreach($f in $FilesToInstall){
      $src=Join-Path $SourceDir $f; $dst=Join-Path $TargetDir $f
      if(-not(Test-Path $src)){ throw "Missing file '$f' in $SourceDir" }
      Copy-Item $src $dst -Force
      if($f -like '*.ppam'){ try{ Unblock-File -Path $dst -ErrorAction SilentlyContinue }catch{} }
    }
  }
  Status "Adding to Trusted Location" { if(-not(Add-TrustedLocation -path $TargetDir -desc $TrustedDesc)){ throw "Could not add location" } }
  $addinPath=Join-Path $TargetDir $AddInFile
  Status "Loading PowerPoint add-in" {
    if(-not(Try-LoadAddIn-COM $addinPath)){ throw "Could not load PowerPoint add-in" }
  }
}

function Uninstall-Addin($paths){
  foreach($p in $paths){
    Status "Unloading add-in"          { Try-UnloadAddIn-COM $p | Out-Null }
  }
  Status "Removing files" {
    foreach($f in $FilesToInstall){
      $dst=Join-Path $TargetDir $f
      if(Test-Path $dst){ Remove-Item $dst -Force -ErrorAction SilentlyContinue }
    }
  }
  Status "Removing Trusted Location" {
    Remove-TrustedLocation -path $TargetDir -desc $TrustedDesc | Out-Null
  }
}

# Main
$paths=Detect-Installed
if(-not $paths -or $paths.Count -eq 0){
  $ans=Read-Host "deckUI is NOT installed. Install now? (Y/N)"
  if($ans -match '^[Yy]'){ Install-Addin }
} else {
  $ans=Read-Host ("deckUI is installed at " + ($paths -join ', ') + ". Uninstall it? (Y/N)")
  if($ans -match '^[Yy]'){ Uninstall-Addin $paths }
}

Write-Host ""
Write-Host "All done. You can close this window now." -ForegroundColor Yellow

exit 0
