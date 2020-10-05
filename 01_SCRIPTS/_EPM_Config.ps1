# ===================================================================================
#   Purpose : House common variables used throughout the automation framework
# ===================================================================================
Param(
        [String]$Process = "ADHOC",
        [ValidateSet("TEST","PROD","SETUP")][String]$ExecEnvironment = "TEST",
        [switch]$UseAPI
    )


$DIR_WORKING = Split-Path -Path $PSScriptRoot -Parent
$DIR_MODULES = "$DIR_WORKING\00_Modules"
Set-Location -Path "$DIR_WORKING\01_SCRIPTS"

# -----------------------------------------------------------------------------------
#   UTILITY IMPORTS
# -----------------------------------------------------------------------------------
#Import-Module "$DIR_MODULES\EPM_Utils\EPM_Utils.psm1" -Force -WarningAction SilentlyContinue -DisableNameChecking
ForEach ($module in (Get-ChildItem -Path "$DIR_MODULES\*.psm1" -Recurse -Force)){
    Unblock-File -Path $module.fullname
    Import-Module $module.fullname -Force -WarningAction SilentlyContinue -DisableNameChecking
}

# Ensure Clean Variables
Remove-Variable EPM_*
Remove-Variable EPMAPI_*
Remove-Variable SV_*

# -----------------------------------------------------------------------------------
#   GLOBAL INFORMATION
# -----------------------------------------------------------------------------------
$EPM_PROCESS_START = Get-Date
$EPM_FILE_STAMP = EPM_Get-TimeStamp -StampType File
$EPM_ENV = $ExecEnvironment
$EPM_PROCESS = $Process
$EPM_AUTO_CALL = "epmautomate"

# -----------------------------------------------------------------------------------
#   FILE & PATH INFORMATION
# -----------------------------------------------------------------------------------
$EPM_PATH_AUTO = $DIR_WORKING
$EPM_PATH_SCRIPTS = "$EPM_PATH_AUTO\01_SCRIPTS"
$EPM_PATH_LOGS = "$EPM_PATH_AUTO\02_LOGS"
$EPM_PATH_FILES_IN = "$EPM_PATH_AUTO\03_FILES\INBOUND"
$EPM_PATH_FILES_OUT = "$EPM_PATH_AUTO\03_FILES\OUTBOUND"
$EPM_PATH_BACKUPS = "$EPM_PATH_AUTO\04_BACKUPS"
$EPM_PATH_ARCHIVES = "$EPM_PATH_AUTO\05_ARCHIVE"
$EPM_PATH_CURRENT_ARCHIVE = $EPM_PATH_ARCHIVES + "\" + $EPM_FILE_STAMP + "_" + $EPM_PROCESS
# -- Default Log / Files
$EPM_LOG_LEVEL = "VERBOSE"
$EPM_LOG_FULL = "$EPM_PATH_LOGS\LOG_FULL.log"
$EPM_LOG_ERROR = "$EPM_PATH_LOGS\LOG_ERRORS.log"
$EPM_LOG_KICKOUTS = "$EPM_PATH_LOGS\LOG_KICKOUTS.log"
$EPM_LOG_SECURITY = "$EPM_PATH_LOGS\SecurityErrors.csv"
$EPM_FILE_LISTFILES = "$EPM_PATH_LOGS\listfiles.txt"
$EPM_FILE_SUBVARS = "$EPM_PATH_LOGS\Subvars.txt"
$EPM_FILE_SECURITY = "$EPM_PATH_FILES_IN\CurrentAppSecurity.csv"
# Archive/Backup Retention Policy can be either "DAYS" or "NUM"
#   DAYS = Will keep archives for the specified number of days in $EPM_ARCHIVES_RETAIN_NUM or $EPM_BACKUPS_RETAIN_NUM
#   NUM = Will keep the last number of archives specified in $EPM_ARCHIVES_RETAIN_NUM or $EPM_BACKUPS_RETAIN_NUM
$EPM_ARCHIVES_RETAIN_POLICY="NUM"
$EPM_ARCHIVES_RETAIN_NUM = 10
$EPM_BACKUPS_RETAIN_POLICY="DAYS"
$EPM_BACKUPS_RETAIN_NUM = 30


# -----------------------------------------------------------------------------------
#   LOGIN INFORMATION
# -----------------------------------------------------------------------------------
$EPM_USER = ""
$EPM_PASSFILE = "`"$EPM_PATH_SCRIPTS\pw.epw`""
$EPM_DOMAIN = ""
$EPM_DATACENTER = ""
# Setup URLS
$EPM_URL_PROD = "https://epm-$EPM_DOMAIN.epm.$EPM_DATACENTER.oraclecloud.com"
$EPM_URL_TEST = "https://epm-test-$EPM_DOMAIN.epm.$EPM_DATACENTER.oraclecloud.com"
#Setup Common URL
$EPM_URL = "NA"
if ($EPM_ENV -eq "PROD") {
    $EPM_URL = $EPM_URL_PROD
} elseif ($EPM_ENV -eq "TEST") {
    $EPM_URL = $EPM_URL_TEST
}

# -----------------------------------------------------------------------------------
#   USE EPM API
# -----------------------------------------------------------------------------------
#if ($useAPI) {
#    . "$EPM_PATH_SCRIPTS\_EPMAPI_Config.ps1"
#}
#$EPMAPI_USED = $UseAPI

# -----------------------------------------------------------------------------------
#   OTHER
# -----------------------------------------------------------------------------------
$EPM_APP = ""
$EPM_TASK_SEPARATOR = "=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-="
$EPM_PROCESS_RUNNING_FLAG = "$EPM_PATH_SCRIPTS\$EPM_PROCESS-Running.flag"
$EPM_TASKLIST = New-EPMTaskList
# Remove Error Log if Exists (would be from previous run)
if (Test-Path $EPM_LOG_KICKOUTS) { Remove-Item $EPM_LOG_KICKOUTS }
if (Test-Path $EPM_LOG_SECURITY) { Remove-Item $EPM_LOG_SECURITY }

# -----------------------------------------------------------------------------------
#   NOTIFICATION SETUP
# -----------------------------------------------------------------------------------
$EPM_EMAIL_TO = @()
$EPM_EMAIL_CC = @()
$EPM_EMAIL_FROM = ""
$EPM_EMAIL_SERVER = ""
$EPM_EMAIL_PORT = 587
$EPM_EMAIL_CREDENTIALS = ""
# Global Notify Level, these can be overridden in at the individual script level
# Options : 
#   NONE = Will not notify
#   SUCCESS = Will send on Success, Warning and Error
#   WARNING = Will send on Warning and Error
#   ERROR = Will send on Error
$EPM_EMAIL_NOTIFYLEVEL = "NONE"
