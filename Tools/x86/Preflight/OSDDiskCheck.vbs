'---------------------------------------------------------
' Combined Disk Check
'---------------------------------------------------------
' Usage: cscript.exe OSDDiskCheck.vbs C:
' - Windows Only (not PE)
' - Return code indicates which check failed
'---------------------------------------------------------
Option Explicit

	'---------------------------------------------------------
	' Configurable Checks
	'---------------------------------------------------------
	Const RequiredDiskSize 	= 20000		' Size in MB
	Const RequiredDiskFree 	= 1000		' Size in MB
	Const RequiredDiskType 	= "NTFS"
	
	'---------------------------------------------------------
	' Configurable Error Codes
	'---------------------------------------------------------
	Const DiskSizeError		= 10
	Const DiskFreeError		= 20
	Const DiskTypeError		= 30

	
	'---------------------------------------------------------
	' Script Start
	'---------------------------------------------------------
	Dim ScriptReturn
		ScriptReturn = -1

	If (Wscript.Arguments.Count=1) Then
		ScriptReturn = RunDiskChecks( Wscript.Arguments(0) )
	Else
		Wscript.Echo "Please supply one argument (drive letter)."
	End If

	QuitScript( ScriptReturn )
	'---------------------------------------------------------
	' Script End
	'---------------------------------------------------------
	
	
	
	
	
'---------------------------------------------------------
' Disk Checks
'---------------------------------------------------------	
Function RunDiskChecks( theDrive )
	Dim ReturnCode
	
	Dim oWMI
	Dim oItems
	Dim oItem
	
	ReturnCode = 0
	
	' --------------------------------------
	' WMI Connect and Query
	' --------------------------------------
	SET oWMI = GetObject("winmgmts:\\.\root\CIMV2")
	SET oItems = oWMI.ExecQuery("SELECT * FROM Win32_LogicalDisk where caption = '" & theDrive & "'", "WQL")

	If ( oItems.Count <> 1 ) Then
		Wscript.Echo "WMI Query returned more/less than one object: [" & oItems.Count & "]"
		RunDiskChecks = -1
		Exit Function
	End If
	
	' --------------------------------------
	' Read WMI Items
	' --------------------------------------
	Dim diskType
	Dim diskSize
	Dim diskFree

	On Error Resume Next
	Err.Clear
	
	For Each oItem In oItems	
		diskSize = Round( oItem.Size / 1024 / 1024 )
		diskFree = Round( oItem.FreeSpace / 1024 / 1024)
		diskType = oItem.FileSystem
	Next
	
	If (Err.Number <>0) Then
		Wscript.Echo "WMI Property Query Error: [" & Err.Number & "]"
		RunDiskChecks = -1
		Exit Function
	End iF
	
	Wscript.Echo "Disk Size   : [" & diskSize & "] MB"
	Wscript.Echo "Disk Free   : [" & diskFree & "] MB"
	Wscript.Echo "Disk Type   : [" & diskType & "]"
	
	If (ReturnCode = 0) Then ReturnCode = RunSizeCheck( diskSize )
	If (ReturnCode = 0) Then ReturnCode = RunFreeCheck( diskFree )
	If (ReturnCode = 0) Then ReturnCode = RunTypeCheck( diskType )
	
	RunDiskChecks = ReturnCode
	
End Function

'---------------------------------------------------------
' Size Check
'---------------------------------------------------------
Function RunSizeCheck( theSize )

	If (theSize < RequiredDiskSize) Then
		Wscript.Echo "Insufficient Disk Size! Requirement: [" & RequiredDiskSize & "] MB"
		RunSizeCheck = DiskSizeError
	Else
		RunSizeCheck = 0	
	End If

End Function


'---------------------------------------------------------
' Type Check
'---------------------------------------------------------
Function RunTypeCheck( theType )

	If (UCase(theType) <> UCase(RequiredDiskType)) Then
		Wscript.Echo "Insufficient Disk Type! Requirement: [" & RequiredDiskType & "]"
		RunTypeCheck = DiskTypeError
	Else
		RunTypeCheck = 0
	End If

End Function


'---------------------------------------------------------
' Free Space Check
'---------------------------------------------------------
Function RunFreeCheck( theSize )

	If (theSize < RequiredDiskFree) Then
		Wscript.Echo "Insufficient Disk Free! Requirement: [" & RequiredDiskSize & "] MB"
		RunFreeCheck = DiskFreeError
	Else
		RunFreeCheck = 0
	End If
	
End Function


'---------------------------------------------------------
' Quit Script
'---------------------------------------------------------
Sub QuitScript( theCode )

	Wscript.Echo "--------------------------------------"
	Wscript.Echo " Script Complete | Exit Code : [" & theCode & "]"
	Wscript.Echo "--------------------------------------"

	Wscript.Quit( theCode )
	
End Sub