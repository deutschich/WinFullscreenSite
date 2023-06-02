Option Explicit

Dim fsShell, i

Set  fsShell = WScript.CreateObject("WScript.Shell")
' Waiting 10 secounds
WScript.Sleep 10000


  fsShell.SendKeys "{F11}"


