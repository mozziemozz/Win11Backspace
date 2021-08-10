$wshell = New-Object -ComObject wscript.shell

$wshell.SendKeys('%({TAB}')

Start-Sleep -Seconds 0.2
$wshell.SendKeys('{BACKSPACE}')