'Simula que foi apertado CTRL+a (seleciona), e F8 (executa)

WScript.Sleep 3000
set oScrShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep 300
oScrShell.SendKeys "%e"
WScript.Sleep 200
oScrShell.SendKeys "{DOWN}"
WScript.Sleep 200
oScrShell.SendKeys "{DOWN}"
WScript.Sleep 200
oScrShell.SendKeys "{DOWN}"
WScript.Sleep 200
oScrShell.SendKeys "{DOWN}"
WScript.Sleep 200
oScrShell.SendKeys "{DOWN}"
WScript.Sleep 1000
oScrShell.SendKeys "1"
WScript.Sleep 500
oScrShell.SendKeys "{ENTER}"
WScript.Sleep 1000
oScrShell.SendKeys "{ENTER}"
WScript.Sleep 1000
oScrShell.SendKeys "^a{F8}"
WScript.Sleep 500






