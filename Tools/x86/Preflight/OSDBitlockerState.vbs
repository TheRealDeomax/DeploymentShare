'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|| BDEstate.vbs
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|| Usage: cscript.exe BDE-state.vbs C:
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|| Checks the drive letter passed for it's encrypted state.
'|| http://msdn.microsoft.com/en-us/library/aa376433(VS.85).aspx
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Option Explicit

'//////////////////////////////////////////////////////////////////////////
'   CONSTANTS
'//////////////////////////////////////////////////////////////////////////
    CONST CMD_BDESTATUS         = "manage-bde.exe -status"
    CONST STATE_TEXT            = "Conversion Status"
    CONST STATE_DECRYPTED       = "Fully Decrypted"

'//////////////////////////////////////////////////////////////////////////
'   GLOBAL VARIABLES
'//////////////////////////////////////////////////////////////////////////
    Dim oTSE
	Dim oWSH
    Set oWSH = CreateObject("Wscript.Shell")

    Dim driveToCheck
    Dim ScriptPath
    ScriptPath = Left( Wscript.ScriptFullName, Len(Wscript.ScriptFullName) - Len(Wscript.ScriptName))

'//////////////////////////////////////////////////////////////////////////
'   MAIN
'//////////////////////////////////////////////////////////////////////////
    Dim exitCode
        exitCode = 0


	SET oTSE = CreateObject("Microsoft.SMS.TSEnvironment")   		
	oTSE("OSDBitlockerStatus") = "Unprotected"

	
    '||||||||||||||||||||||||||||||||||||||||||
    ' Arguments
    '||||||||||||||||||||||||||||||||||||||||||
    If ( Wscript.Arguments.Count <> 1 ) Then
        Wscript.echo "Please provide a drive letter."
        QuitScript(-1)
    End if

    '||||||||||||||||||||||||||||||||||||||||||
    ' Format Drive Argument
    '||||||||||||||||||||||||||||||||||||||||||
    driveToCheck = Wscript.Arguments(0)
    
    if ( Len(driveToCheck) > 2 ) Then driveToCheck = Left(driveToCheck, 2)
    if ( Len(driveToCheck) = 1 ) Then driveToCheck = driveToCheck & ":"
    
    Wscript.Echo 
    Wscript.Echo "[---------------------------------------------]"
    wscript.echo " Drive Argument: " & drivetoCheck
    Wscript.Echo "[---------------------------------------------]"
    Wscript.Echo

    '||||||||||||||||||||||||||||||||||||||||||
    ' Drive Query
    '||||||||||||||||||||||||||||||||||||||||||
    Dim aEncVolumes,oTemp
    
    Err.Clear
    Set aEncVolumes = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy}!//./root/cimv2/Security/MicrosoftVolumeEncryption:Win32_EncryptableVolume").Instances_

	'||||||||||||||||||||||||||||||||||||||||||
	'wbemErrInvalidNamespace 
	'||||||||||||||||||||||||||||||||||||||||||
	If (cstr(hex(err.number)) = "8004100E") Then 
		wscript.echo "Namespace not found. Exiting with success."
		QuitScript ( 0 )
	End If
	
    If ( err.number <> 0 ) Then
        wscript.echo "ERROR connecting to WMI: " & err.number
        QuitScript ( -1 ) 
    End If

    Dim wmiDrive
    Dim wmiDriveEncStatus
    Dim wmiDriveEncPercent

    

    '||||||||||||||||||||||||||||||||||||||||||
    ' Loop through all drives
    '||||||||||||||||||||||||||||||||||||||||||
    For Each oTemp In aEncVolumes
    
        wmiDrive = oTemp.DriveLetter
        Wscript.echo "..Found WMI drive: [" & wmiDrive & "]"
        
        '||||||||||||||||||||||||||||||||||||||||||
        ' Drive Match
        '||||||||||||||||||||||||||||||||||||||||||
        if ( InStr(1, wmiDrive, driveToCheck, 1) ) Then
            wmiDriveEncStatus = -1
            
            wscript.echo "Checking BDE Status: " & wmiDrive
			' ===============================
			'  0 = Fully Decrypted
			'  1 = Fully Encrypted
			'  2 = Encryption in Progress
			'  3 = Decrypt in Progress
			'  4 = Encryption Paused
			'  5 = Decryption Paused
			' ===============================
            If ( oTemp.GetConversionStatus(wmiDriveEncStatus, wmiDriveEncPercent) = 0 ) Then 
                exitCode = wmiDriveEncStatus
                wscript.echo
                wscript.echo "-----------------------"
                wscript.echo "Drive Match Found:"
                wscript.echo "..Conversion Status: " & wmiDriveEncStatus
                wscript.echo "..Exit code equals conversion status: " & exitCode
                wscript.echo "-----------------------"
                wscript.echo
            Else
                wscript.echo ("..ERROR, call to GetConversionStatus failed.")
            End If
            
        Else
            wscript.echo "..No match, skipping: " & wmiDrive
        End If
    Next

    '||||||||||||||||||||||||||||||||||||||||||
    ' If no drives are found, will exit successfully.
    '||||||||||||||||||||||||||||||||||||||||||
	if ( exitCode > 0) Then oTSE("OSDBitlockerStatus") = "Protected"
	
    '||||||||||||||||||||||||||||||||||||||||||
    ' All done.
    '||||||||||||||||||||||||||||||||||||||||||
    QuitScript ( exitCode )

'//////////////////////////////////////////////////////////////////////////
'   FUNCTIONS
'//////////////////////////////////////////////////////////////////////////

Sub QuitScript ( QuitCode )
    Wscript.Echo 
    Wscript.Echo "[---------------------------------------------]"
    wscript.echo " Script Complete | Exit: " & QuitCode
    Wscript.Echo "[---------------------------------------------]"
    Wscript.Echo

    wscript.quit( QuitCode )
End Sub