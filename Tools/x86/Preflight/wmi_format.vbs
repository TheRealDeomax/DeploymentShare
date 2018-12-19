' ////////////////////////////////////////////////////////////////////
' // Format Volume
' ////////////////////////////////////////////////////////////////////
' // Created: MICHS 4/7/2009
' ////////////////////////////////////////////////////////////////////
' // Use: cscript.exe wmi_format.vbs c:
' ////////////////////////////////////////////////////////////////////
Option Explicit


	'||||||||||||||||||||||||||||||||||||||
	' Ensure at least one parameter passed.
	'||||||||||||||||||||||||||||||||||||||
	if (Wscript.Arguments.Count <> 1) Then
		Wscript.echo "Error: Drive argument not supplied. Please provide at least one drive letter. Example: C:"
		Wscript.Quit (-1)
	Elseif ( Len ( Wscript.Arguments(0)) <2 ) Then
		Wscript.echo "Error: Drive argument invalid. Please provide at least one drive letter. Example: C:"
		Wscript.Quit (-2)	
	End If

	'||||||||||||||||||||||||||||||||||||||
	' Argument Format
	'||||||||||||||||||||||||||||||||||||||
	Dim sDrive
	sDrive = Left(Wscript.Arguments(0), 2)
	Wscript.Echo "Drive Argument: " & sDrive

	
	'||||||||||||||||||||||||||||||||||||||
	' Search for drive/volume
	'||||||||||||||||||||||||||||||||||||||
	Dim oWMI
	Dim oVOL
	
	SET oWMI = GetObject("winmgmts:\\.\root\cimv2")
	SET oVOL = oWMI.ExecQuery("select * from win32_volume where driveletter='" & sDrive & "'")
	
	If ( oVOL.Count < 1 ) Then
		WScript.Echo "Error: Volume not found."
		Wscript.Quit (3)
	End If

	If ( oVOL.Count > 1 ) Then
		WScript.Echo "Error: Too many volumes found."
		Wscript.Quit (4)
	End If
	
	'||||||||||||||||||||||||||||||||||||||
	' Perform Format
	'||||||||||||||||||||||||||||||||||||||
	Dim singleVolume
	Dim retVal
	
	For Each singleVolume in oVOL
      retVal = singleVolume.Format( "NTFS", true, 4096, "", false)
	Next

		
	
	'||||||||||||||||||||||||||||||||||||||
	' Quit
	'||||||||||||||||||||||||||||||||||||||
	Wscript.Echo "Format Result: [" & retVal & "]"
	Wscript.Quit (retVal)

	
