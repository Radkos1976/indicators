   '***************
   ' *** 64bit check
   ' ***************
   ' check to see if we are on 64bit OS -> re-run this script with 32bit cscript
   Function RestartWithCScript32(extraargs)
   Dim strCMD, iCount
   strCMD = r32wShell.ExpandEnvironmentStrings("%SYSTEMROOT%") & "\SysWOW64\cscript.exe"
   If NOT r32fso.FileExists(strCMD) Then strCMD = "cscript.exe" ' This may not work if we can't find the SysWOW64 Version
   strCMD = strCMD & Chr(32) & Wscript.ScriptFullName & Chr(32)
   If Wscript.Arguments.Count > 0 Then
    For iCount = 0 To WScript.Arguments.Count - 1
     if Instr(Wscript.Arguments(iCount), " ") = 0 Then ' add unspaced args
      strCMD = strCMD & " " & Wscript.Arguments(iCount) & " "
     Else
      If Instr("/-\", Left(Wscript.Arguments(iCount), 1)) > 0 Then ' quote spaced args
       If InStr(WScript.Arguments(iCount),"=") > 0 Then
        strCMD = strCMD & " " & Left(Wscript.Arguments(iCount), Instr(Wscript.Arguments(iCount), "=") ) & """" & Mid(Wscript.Arguments(iCount), Instr(Wscript.Arguments(iCount), "=") + 1) & """ "
       ElseIf Instr(WScript.Arguments(iCount),":") > 0 Then
        strCMD = strCMD & " " & Left(Wscript.Arguments(iCount), Instr(Wscript.Arguments(iCount), ":") ) & """" & Mid(Wscript.Arguments(iCount), Instr(Wscript.Arguments(iCount), ":") + 1) & """ "
       Else
        strCMD = strCMD & " """ & Wscript.Arguments(iCount) & """ "
       End If
      Else
       strCMD = strCMD & " """ & Wscript.Arguments(iCount) & """ "
      End If
     End If
    Next
   End If
   r32wShell.Run strCMD & " " & extraargs, 0, False
   End Function

   Dim r32wShell, r32env1, r32env2, r32iCount
   Dim r32fso
   SET r32fso = CreateObject("Scripting.FileSystemObject")
   Set r32wShell = WScript.CreateObject("WScript.Shell")
   r32env1 = r32wShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
   If r32env1 <> "x86" Then ' not running in x86 mode
    For r32iCount = 0 To WScript.Arguments.Count - 1
     r32env2 = r32env2 & WScript.Arguments(r32iCount) & VbCrLf
    Next
    If InStr(r32env2,"restart32") = 0 Then RestartWithCScript32 "restart32" Else MsgBox "Cannot find 32bit version of cscript.exe or unknown OS type " & r32env1
    Set r32wShell = Nothing
    WScript.Quit
   End If
   Set r32wShell = Nothing
   Set r32fso = Nothing
   ' *******************
   ' *** END 64bit check
   ' *******************

 'SprawdŸ czy process ju¿ nie chodzi

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
chk="false"
chk1="false"
Set colProcesses = objWMIService.ExecQuery _
("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
Wscript.Echo "Liczba skryptów " & colProcesses.Count
If colProcesses.Count > 0 Then
For Each objitem In colProcesses
    if instr(1,objitem.CommandLine,"ARC.vbs")>0 then chk="true"
    if instr(1,objitem.CommandLine,"SUM.vbs")>0 then chk1="true"	
Next
end if
Const HIDDEN_WINDOW = 0
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set objStartup = objWMIService.Get("Win32_ProcessStartup")
Set objConfig = objStartup.SpawnInstance_
objConfig.ShowWindow = HIDDEN_WINDOW
If chk<>"true" Then
Set objProcess = objWMIService.Get("Win32_Process")
intReturn = objProcess.Create _
    ("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\ARC.vbs", Null, objConfig, intProcessID)
end if
'if chk1<>"true" then
'intReturn = objProcess.Create _
'    ("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\SUM.vbs", Null, objConfig, intProcessID)
'end if
Wscript.Echo "Complete"
