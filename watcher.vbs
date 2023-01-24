Set Args = WScript.Arguments

' *******************************************
' Script, that allow to limit quantity of the
' started applications on a terminal server.
' The names of applications and the maximum
' quantity of simultaneously started spears
' are stored in config file.
'
' @authors: SkyF, hasherfrog
' @site: oszone.net
'

CONST FIELDS = 3
CONST PSNAME = 0
CONST PSNUMB = 1
CONST PSMAXN = 2

Dim thePsArray ( )
Dim nToWatch
nToWatch = 0

Function AddPsToWath( name, number )
ReDim Preserve thePsArray ( FIELDS, nToWatch )
thePsArray(PSNAME, nToWatch) = name
thePsArray(PSNUMB, nToWatch) = number
thePsArray(PSMAXN, nToWatch) = number
addPsToWatch = nToWatch
nToWatch = nToWatch + 1
End Function

' *******************************************

Sub ReadPsWatchFile(filename)
Dim fso, f
' On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(filename, 1, False)
if Err.Number<>0 then
call Error_(0)
else
while not f.atEndOfStream
s = f.ReadLine
ms = Split(s, " ", -1, vbBinaryComapre)
s1 = ms(0)
n2 = CInt(ms(1))
n = AddPsToWath(s1, n2)
Err.Clear
wend
f.Close
end if
End Sub

'Протоколирование в одноимённый log файл
Sub AppLog( aMessage )
Dim fso, f, flnm, renflnm, dtmstr
flnm = Replace(Wscript.ScriptFullName,".vbs",".log")
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(flnm)) Then
Set f = fso.GetFile(flnm)
If f.Size > 100000 then
dtmstr = Replace(Replace(Replace(FormatDateTime(Now),"/","-"),":",".")," ","-")
renflnm = Replace(flnm,".log","-" & dtmstr & ".log")
f.Move renflnm
End If
Set f = Nothing
End If
Set f = fso.OpenTextFile(flnm, 8, True)
f.WriteLine FormatDateTime(Now) & " - " & aMessage
f.Close
End Sub

'********************************************
Function CountUserProcess( ProcessName, ProcessId )
Dim strNameOfUser,strOwnerProcess,strDomain
CountUserProcess = 0
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where ProcessId=" & CStr(ProcessId) )
For Each objProcess in colProcessList
colProps = objProcess.GetOwner(strOwnerProcess, strDomain)
Next
AppLog "-- ProcessName-" & ProcessName & ", ProcessId-" & CStr(ProcessId) & ", strOwnerProcess-" & strOwnerProcess & ", strDomain-" & strDomain
Set colProcessList = Nothing
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & ProcessName & "'")
For Each objProcess in colProcessList
colProps = objProcess.GetOwner(strNameOfUser, strDomain)
AppLog "strNameOfUser-" & strNameOfUser
if strNameOfUser = strOwnerProcess then
CountUserProcess = CountUserProcess + 1
end if
Next
End Function

' *******************************************

Sub StartUpWatcher()
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE Targetinstance ISA 'Win32_Process'")
Do While 1
Set objLatestProcess = colMonitoredEvents.NextEvent
For psn=0 to nToWatch-1
If objLatestProcess.TargetInstance.Name = thePsArray(PSNAME, psn) Then
If objLatestProcess.Path_.Class = "__InstanceCreationEvent" Then
psCount = CountUserProcess(objLatestProcess.TargetInstance.Name, objLatestProcess.TargetInstance.ProcessId)
'Wscript.Echo thePsArray(PSNAME, psn), psCount
iF psCount > thePsArray(PSNUMB, psn) Then
Applog "Count of process " & objLatestProcess.TargetInstance.Name & " is " & Cstr(psCount) & "... Terminate last process."
objLatestProcess.TargetInstance.Terminate
End If
End If
End If
Next
Loop
End Sub

' *******************************************

for i=0 to Args.count-1
If Args(i) = "-f" Then
iaFilename = Args(i+1)
End If
Next

If iaFilename<>"" Then
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(iaFilename)) Then
Call ReadPsWatchFile(iaFilename)
If nToWatch=0 Then
Wscript.Echo "File is empty."
Else
Call StartUpWatcher()
End If
Else
Wscript.Echo "File exists not."
End If
Else
Wscript.Echo "watcher.vbs -f config.txt"
End If