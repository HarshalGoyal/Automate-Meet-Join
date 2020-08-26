Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c terminal.bat"
oShell.Run strArgs, 0, false