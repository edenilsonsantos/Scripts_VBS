'Desenvolvido por Edenilson Santos - PackTi Serviços


set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "firefox" 
WScript.Sleep 3000
WshShell.AppActivate "firefox"
WScript.Sleep 100
WshShell.SendKeys "^l"
WScript.Sleep 100

'Digite na linha abaixo entre as aspas, o site que vai abrir

WshShell.SendKeys "http://191.101.235.224/triagem"
WScript.Sleep 2000
WshShell.SendKeys "{ENTER}"
WScript.Sleep 500
WshShell.SendKeys "{F11}"
WScript.Sleep 500