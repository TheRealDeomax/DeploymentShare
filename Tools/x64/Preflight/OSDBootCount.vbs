'---------------------------------------------------------
' OSD Bootable Operating System Count
'---------------------------------------------------------
' Usage: cscript osdbootcount.vbs
' - Windows Vista and Windows 7 only
' - Return code indicates count of displayed operating systems
'   during boot
' - Return codes > 900 indicate an error
'---------------------------------------------------------
'
' Modified MSDN sample from
' "Boot Configuration Data in Windows Vista"
' (http://download.microsoft.com/download/a/f/7/af7777e5-7dcd-4800-8a0a-b18336565f5b/BCD.docx)
'
Option Explicit


	'---------------------------------------------------------
	' These hardcoded values will be replaced with official constants when
	' available.
	'---------------------------------------------------------
	const C_BOOT_MGR_ID = "{9dea862c-5cdd-4e70-acc1-f32b344d4795}"
	const C_DISPLAY_ORDER_TYPE = &h24000001


	'---------------------------------------------------------
	' Globals
	'---------------------------------------------------------
	Dim oBCDStore
	Dim oBCDManager
	Dim iCount

	'---------------------------------------------------------
	' Script Start
	'---------------------------------------------------------
	
	If ( ConnectBCD()        = False )  Then QuitScript( 910 )
	If ( OpenBootManager()   = False )  Then QuitScript( 920 )
	If ( CountDisplayOrder() = False)   Then QuitScript( 930 )

	Wscript.Echo
	Wscript.Echo " Final Count: [" & iCount & "]"
	Wscript.Echo
	
	QuitScript( iCount )
	
	'---------------------------------------------------------
	' Script End
	'---------------------------------------------------------
	
	
'#########################################################
'---------------------------------------------------------
' Functions
'---------------------------------------------------------
'#########################################################

	
	'---------------------------------------------------------
	' Connect to WMI BCD
	'---------------------------------------------------------
	Function ConnectBCD
	
		Dim oWMI
		Dim oSVC
		Dim BcdStoreClass
	
		SET oWMI = CreateObject("WbemScripting.SWbemLocator")
		SET oSVC = oWMI.ConnectServer(".", "root\wmi")
		
		oSVC.Security_.ImpersonationLevel = 3
	
		'---------------------------------------------------------
		'Open up a connection to WMI BcdStore class, allowing for
		'impersonation. We need to request that Backup and Restore
		'privileges be granted as well.
		'---------------------------------------------------------
		
		On Error Resume Next

		SET BcdStoreClass = GetObject("winmgmts:{impersonationlevel=impersonate,(Backup,Restore)}!" & "root/wmi:bcdStore")
		
		If ( ErrorPresent() = True) Then 
			WScript.Echo "Couldn't connect to root/wmi:BcdStore!"
			ConnectBCD = False
			Exit Function
		End If

		On Error Goto 0
		
		If Not BcdStoreClass.OpenStore( "",oBCDStore ) Then
			WScript.Echo "Couldn't open the system store!"
			ConnectBCD = False
			Exit Function
		End If
		
		ConnectBCD = True
	
	End Function

	
	'---------------------------------------------------------
	' Open Boot Manager
	'---------------------------------------------------------
	Function OpenBootManager
	
		'---------------------------------------------------------
		' Open the "boot manager" object.
		'---------------------------------------------------------
		If Not oBCDStore.OpenObject( C_BOOT_MGR_ID, oBCDManager ) then
			WScript.Echo "Couldn't open the boot manager object!"
			OpenBootManager = False
			Exit Function
		End If
		
		OpenBootManager = True
	
	End Function

	
	'---------------------------------------------------------
	' Count bootable "Display Order" items
	'---------------------------------------------------------
	Function CountDisplayOrder
	
		Dim bootOrderList
		Dim OSIdentifier

		'---------------------------------------------------------
		' Get the boot manager's display order list.
		'---------------------------------------------------------
		If Not oBCDManager.GetElement( C_DISPLAY_ORDER_TYPE, bootOrderList ) Then
			WScript.Echo "Couldn't get the display order list!"
			CountDisplayOrder = False
			Exit Function
		Else
			'---------------------------------------------------------
			' Remove the target os from the boot order list.
			'---------------------------------------------------------
			iCount = 0
			For Each OSIdentifier In bootOrderList.Ids
			   iCount = iCount + 1
			   wscript.echo " Boot Display Item: " + OSIdentifier
			Next

		End If
		
		Wscript.Echo
		Wscript.Echo "--------------------------------------"
		WScript.Echo "Successfully enumerated display order entries."
		Wscript.Echo "--------------------------------------"
		Wscript.Echo
		
		CountDisplayOrder = True
	
	End Function


	'---------------------------------------------------------
	' Quit Script
	'---------------------------------------------------------
	Sub QuitScript( theCode )

		Wscript.Echo
		Wscript.Echo "--------------------------------------"
		Wscript.Echo " Script Complete | Exit Code : [" & theCode & "]"
		Wscript.Echo "--------------------------------------"
		Wscript.Echo

		Wscript.Quit( theCode )
		
	End Sub
	
	
	'---------------------------------------------------------
	' Display Error Details (if present)
	'---------------------------------------------------------
	Function ErrorPresent()
	
		If Err.Number <> 0 Then 
			Wscript.Echo "--------------------------------------"
			Wscript.Echo " Error Number      : [" & Err.Number & "]"
			Wscript.Echo " Error Description : [" & Err.Description & "]"
			Wscript.Echo "--------------------------------------"
			ErrorPresent = True
		Else
			ErrorPresent = False
		End If
	
	End Function