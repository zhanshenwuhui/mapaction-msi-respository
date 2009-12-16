'<?xml version="1.0" encoding="ISO-8859-1"?>

'*************************************************
' Script    UnRegister.vbs
'Andy Smith 19/11/09
'*************************************************
'<job id="test">
'<script language="VBScript">
'<![CDATA[

'%0\..\Register.bat /u

Option Explicit
On Error Resume Next
'On Error GoTo 0
Const myDEBUG = false

Dim wshShell
Set wshShell = CreateObject("WScript.Shell")

Dim myStdErr
Set myStdErr = WScript.StdErr

Dim pathStr 
pathStr = left(Wscript.ScriptFullName, (len(Wscript.ScriptFullName) - len(Wscript.ScriptName)))

Dim cmdStr
cmdStr = pathStr & "Register.bat /u"
'MsgBox  cmdStr

'Use RegAsm.exe to unregister AddETspatialTB.dll
retCode = wshShell.Run(cmdStr, 0, true)

if myDEBUG then
     myStdErr.WriteLine "retCode = " & retCode 
end if

']]>
'</script>
'</job>