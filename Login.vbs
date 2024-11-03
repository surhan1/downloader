Set WshShell = CreateObject("WScript.Shell")

Dim userFolder
userFolder = CreateObject("WScript.Network").UserName
startFile = "C:\Users\" & userFolder & "\Processes\Start.vbs"

WshShell.Run "wscript.exe """ & startFile & """", 0, False
