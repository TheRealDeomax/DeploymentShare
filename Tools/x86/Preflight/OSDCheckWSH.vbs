'---------------------------------------------------------
'Validate Windows Scripting Host
'---------------------------------------------------------

Option Explicit
On Error Resume Next

Dim oShell
Dim oNetwork
Dim oFSO
Dim oEnv

' Create general-purpose WSH objects.  These should always succeed; if not, 
' WSH is seriously broken.
Set oShell = CreateObject("WScript.Shell")
If Err then
	WScript.Echo "Create WScript.Shell object failed"
	WScript.Quit 1
End if
WScript.Echo "Create WScript.Shell object OK"

Set oNetwork = CreateObject("WScript.Network")
If Err then
	WScript.Echo "Create WScript.Network object failed"
	WScript.Quit 1
End if
WScript.Echo "Create WScript.Network object OK"

Set oFSO = CreateObject("Scripting.FileSystemObject")
If Err then
	WScript.Echo "Create WScript.FileSystemObject object failed"
	WScript.Quit 1
End if
WScript.Echo "Create WScript.FileSystemObject object OK"

Set oEnv = oShell.Environment("PROCESS")
If Err then
	WScript.Echo "Create Shell Process Environment failed"
	WScript.Quit 1
End if
WScript.Echo "Create Shell Process Environment OK"

WScript.Echo ""
WScript.Quit 0