
$startTime = Get-Date
Start-Sleep 8
$Elapsed = New-TimeSpan -Start $StartTime -End (Get-Date)
if ($Elapsed.TotalSeconds -le 60) {
	$wait = 60 - $Elapsed.TotalSeconds
	Write-Host "Wait : $wait"
} else {
	Write-Host "Ok to run $($Elapsed.TotalSeconds)"
}