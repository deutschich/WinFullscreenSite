Option Explicit

Dim fsShell, i

Set  fsShell = WScript.CreateObject("WScript.Shell")
' 10 Sekunden warten
WScript.Sleep 10000


  fsShell.SendKeys "{F11}"


