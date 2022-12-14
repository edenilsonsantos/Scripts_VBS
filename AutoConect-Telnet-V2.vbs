' By Edenilson Santos
' Faz login automatico via Telnet
' Faz a captura dos dados para o arquivo log.txt
' O IP foi alterado para um ficticio por seguranca

' Captura data e coloca em uma variavel, para usar no nome do arquivo
dt = right("0" & day(now),2) & right("0" & month(now),2) &  year(now)
'wscript.echo dt

WScript.Sleep 2000
set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "cmd" 
WScript.Sleep 500
WshShell.AppActivate "cmd"
WScript.Sleep 500 
WshShell.SendKeys "telnet" & "~"
WScript.Sleep 2000 
WshShell.SendKeys "set logfile c:\temp\log_"& dt &".txt" & "~" 
WScript.Sleep 2000 
WshShell.SendKeys "open 10.10.10.1 2300" & "~" 
WScript.Sleep 2000 
WshShell.SendKeys "smdr" & "~" 
WScript.Sleep 2000 
WshShell.SendKeys "pccsmdr" & "~" 
WScript.Sleep 2000 
