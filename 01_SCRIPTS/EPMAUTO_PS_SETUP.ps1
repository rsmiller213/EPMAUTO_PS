# ===================================================================================
#   Author : Randy Miller (SolveX Consulting, LLC)
#   Created On : 04-02-2020
#   Purpose : EPM Automate Setup Script
# ===================================================================================

Unblock-File -Path "$PSScriptRoot\*"

Write-Host -ForegroundColor Yellow @"
Please ensure that you have set the variables under "Login Information" in the _EPM_Config.ps1 file before continuing.
    EPM_USER
    EPM_PASSFILE (do not need to run encryption, this setup will prompt you and will then encrypt)
    EPM_DOMAIN
    EPM_DATACENTER
"@

$Ans = Read-Host -Prompt "Have You Set these Variables Up? [Y/N]:"

if ($Ans -notlike 'y*') {
    Write-Host -ForegroundColor Yellow "Please Update and Save Config File, then re-run setup"
    break
}

# Grab Config Variables
Write-Host -ForegroundColor Yellow "Grabbing configuration, you may be prompted to enter a password for the user defined which will then be encrypted"
. "$PSScriptRoot\_EPM_Config.ps1" -Process "Setup" -UseAPI -ExecEnvironment "TEST"

Get-Variable EPM_ENV,EPM_PROCESS,EPM_USER,EPM_PASSFILE,EPM_DOMAIN,EPM_DATACENTER,EPM_URL,EPM_LOG_ERROR,EPM_LOG_FULL,EPM_PATH_CURRENT_ARCHIVE,EPMAPI_PASSFILE,EPMAPI_PLN_BASE_URI,EPMAPI_MIG_BASE_URI,EPMAPI_DMG_BASE_URI


Write-Host -ForegroundColor Yellow "`nTesting for Pass Files"
if (Test-Path $EPM_PASSFILE.Replace('"','')) { 
    Write-Host -ForegroundColor Green "   $($EPM_PASSFILE.Replace('"','')) Found!" 
} else { 
    Write-Host -ForegroundColor Red "   $($EPM_PASSFILE.Replace('"','')) Not Found!"
    break
}
if (Test-Path $EPMAPI_PASSFILE) { 
    Write-Host -ForegroundColor Green "   $EPMAPI_PASSFILE Found!" 
} else { 
    Write-Host -ForegroundColor Red "   $EPMAPI_PASSFILE Not Found!"
    break
}


Write-Host -ForegroundColor Yellow "`nSetting Up Folder Structure & Testing Login"
EPM_Start-Process
EPM_End-Process

Write-Host -ForegroundColor Yellow "`nTesting Folder Structure"
$PathVars = Get-Variable EPM_PATH* -Exclude EPM_PATH_CURRENT_ARCHIVE,EPM_PATH_AUTO -ValueOnly
ForEach ($item in $PathVars) {
    if (Test-Path $item) { 
        Write-Host -ForegroundColor Green "   $item Found!" 
    } else {
        Write-Host -ForegroundColor Red "   $item Missing!"
        break
    }
}
if (Test-Path "$EPM_PATH_CURRENT_ARCHIVE.zip") {
    Write-Host -ForegroundColor Green "   $EPM_PATH_CURRENT_ARCHIVE.zip Found!"
} else {
    Write-Host -ForegroundColor Red "   $EPM_PATH_CURRENT_ARCHIVE.zip Missing!"
    break
}