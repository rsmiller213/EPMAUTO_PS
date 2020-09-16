# ===================================================================================
#   Purpose : House common variables used throughout the automation framework for the REST API
# ===================================================================================

# -----------------------------------------------------------------------------------
#   UTILITY IMPORTS
# -----------------------------------------------------------------------------------
Import-Module "$DIR_MODULES\EPMAPI_Utils\EPMAPI_Utils.psm1" -Force -WarningAction SilentlyContinue -DisableNameChecking

Remove-Variable EPMAPI_*

# -----------------------------------------------------------------------------------
#   EPM API
# -----------------------------------------------------------------------------------
$EPMAPI_PASSFILE = "$EPM_PATH_SCRIPTS\API_pw.epw"
If (-not (Test-Path -Path $EPMAPI_PASSFILE)) {
    $Pass = Read-Host -Prompt "Please Enter Password for $EPM_USER"
    EPMAPI_Create-AuthFile -Pass $Pass
}


If (Test-Path -Path $EPMAPI_PASSFILE) {
    $EPMAPI_HEADER = @{"Authorization"="Basic " + (Get-Content -Path "$EPMAPI_PASSFILE")}
    # Get Version / Base URI
    EPMAPI_Create-APIVerURIFile
    $EPMAPI_INFO = ConvertFrom-StringData(Get-Content "$EPM_PATH_SCRIPTS\EPMAPI_Info.txt" -raw)
    $EPMAPI_PLN_VERSION = "$($EPMAPI_INFO.PLAN_VERSION)"
    $EPMAPI_PLN_BASE_URI = "$EPM_URL/HyperionPlanning/rest/$EPMAPI_PLN_VERSION"
    $EPMAPI_MIG_VERSION = "$($EPMAPI_INFO.MIG_VERSION)"
    $EPMAPI_MIG_BASE_URI = "$EPM_URL/interop/rest/$EPMAPI_MIG_VERSION"
    $EPMAPI_DMG_VERSION = "$($EPMAPI_INFO.DMG_VERSION)"
    $EPMAPI_DMG_BASE_URI = "$EPM_URL/aif/rest/$EPMAPI_DMG_VERSION"
    Remove-Item "$EPM_PATH_SCRIPTS\EPMAPI_Info.txt"
    Remove-Variable EPMAPI_INFO
} else {
    Write-Host "Auth File Doesn't Exist @ $EPMAPI_PASSFILE" -ForegroundColor Red
}