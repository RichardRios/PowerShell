$shell = New-Object -ComObject "Wscript.Shell"
$minutes = 60

for($i = 0; $i -lt $minutes; $i++)
{
	$remaining = $($minutes - $i)
    $shell.SendKeys(".")
	if($remaining -gt ($minutes / 3))
	{
		Write-Host -ForegroundColor Green "$remaining minutes remaining"
	}
	elseif(($remaining -le ($minutes / 3)) -and ($remaining -gt ($minutes / 6)))
	{
		Write-Host -ForegroundColor Yellow "$remaining minutes remaining"
	}
	else
	{
		Write-Host -ForegroundColor Red "$remaining minutes remaining"
	}
	Start-Sleep -Seconds 60
}