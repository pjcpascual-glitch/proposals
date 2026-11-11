' ATI Offer Tracker — Silent 1-Click Desktop Launcher
' Double-clicking this script launches the server silently in the background (no black terminal window)
' and opens the dashboard in your default web browser.

Set WshShell = CreateObject("WScript.Shell")
strCurrentDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
WshShell.CurrentDirectory = strCurrentDir

' Launch Python server silently (WindowStyle 0 = Hidden)
WshShell.Run "py server.py", 0, False

' Wait 1.5 seconds for the server to bind, then launch browser
WScript.Sleep 1500
WshShell.Run "http://localhost:8000/", 1, False

