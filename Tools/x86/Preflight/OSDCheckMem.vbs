'==================================================================
' Memory Preflight Check
'==================================================================

Option Explicit
On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const MinimumMemoryMB = 1000

wscript.echo
wscript.echo "--------------------- Script Begin"
wscript.echo


wscript.echo
wscript.echo "--------------------- Query WMI for RAM"
wscript.echo

Dim objWMIService
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

Dim colItems
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

Dim objItem, strMem
For Each objItem In colItems
   strMem = round((objItem.TotalPhysicalMemory /1024)/1024)
   Exit For
Next

 If Err.Number <> 0 Then
		Wscript.Echo "Error: " & Err.Number 
		Wscript.Echo "Description: " & Err.Description
		Wscript.Echo "Error occurred while calculating Computers Memory. Exiting with [" & Err.Number & "]."
		Wscript.Quit (Err.Number)
  End If

Wscript.Echo "Required physical memory is   : " & MinimumMemoryMB & " MB"
Wscript.Echo "Available physycal memory is  : " & strMem & " MB"


wscript.echo
wscript.echo "--------------------- Running Logic"
wscript.echo

If Int(strMem) < MinimumMemoryMB Then 
   wscript.echo "FAIL: Not enough memory!"
   wscript.Quit -1
Else
   wscript.echo "PASS: Sufficient Memory"
   wscript.Quit 0 
End If