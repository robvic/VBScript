Option Explicit
Dim counter
Dim WshShell
counter = 5
While counter > 0
set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep(3000)
WshShell.SendKeys "%{ESC}" '%{TAB} não funciona!

counter = counter - 1
wend