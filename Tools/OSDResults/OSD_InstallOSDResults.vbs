'///////////////////////////////////////////////////////
' OSD_InstallOSDResults
'///////////////////////////////////////////////////////
'08.20.09 | v1.0.0.0 | MICHS & BENSHY
'09.27.11 | v1.0.0.1 | Cameronk - Updated all usage of modena to UDI (MDT bug 11749)
'10.17.11 | v1.0.0.2 | Cameronk - Added Scheduled task to cleanup %windir%\UDI
'10.17.11 | v1.0.0.3 | Cameronk - Added logic to only set AppInstall.exe reg key if CM 2012 or above client is installed
'03.08.12 | v1.0.0.4 | v-cbarr - fixed scheduled task creation bug where it wasnt creating as SYSTEM. also added a space between /c ScheduledTaskName in executed cmdline
'04.13.12 | v1.0.0.5 | jerod - wait for TS before setting registry keys, wait for WinUpdate before rebooting, enable mouse pointer for win 8
'///////////////////////////////////////////////////////
Option Explicit

	' ############################################################ CONST

    Const CmdRegAddResults      = "reg add HKLM\SYSTEM\Setup /v CmdLine /d ""System32\wscript.exe //B %windir%\UDI\Bootstrapper.vbs"" /f"
    Const CmdRegAddSetupType    = "reg add HKLM\SYSTEM\Setup /v SetupType /d 2 /t REG_DWORD /f"

    Const CmdPeekTasks          = "cmd /c tasklist /v > %windir%\UDI\TaskList.log"
    Const CmdPeekRegBefore      = "cmd /c reg query HKLM\SYSTEM\Setup > %windir%\UDI\regbefore.reg"
    Const CmdPeekRegAfter       = "cmd /c reg query HKLM\SYSTEM\Setup > %windir%\UDI\regafter.reg"
    Const CmdRegAddAppInstall   = "reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v RunAppInstall /d %windir%\UDI\AppInstall.exe /f"
    Const CmdRegAddOSDResults   = "reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce /v RunOSDResults /d %windir%\UDI\OSDResults.exe /f"
    Const SCHTask               = "schtasks /create /RU ""NT AUTHORITY\SYSTEM"" /tn UDI_cleanup /rl HIGHEST /SC ONCE /ST 12:00 /f /tr ""cmd /c rd /s /q %windir%\UDI"""
    Const SCHTaskRegClean       = "schtasks /create /RU ""NT AUTHORITY\SYSTEM"" /tn UDI_Regcleanup /rl HIGHEST /SC ONCE /ST 06:00 /f /tr ""reg delete HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run /v RunAppInstall /f"""
    Const SCHTaskRestoreMouse   = "schtasks /create /RU ""NT AUTHORITY\SYSTEM"" /tn UDI_RegRestoreMousecleanup /rl HIGHEST /SC ONLOGON /f /tr ""reg add HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\System /v EnableCursorSuppression /t REG_DWORD /d 1 /f"""
    Const CursorKey             = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    Const CursorKeyValueName    = "EnableCursorSuppression"
    Const CmdIsTouchEnabled     = "%windir%\UDI\IsTouchEnabled.exe"

	Const HKEY_LOCAL_MACHINE = &H80000002

    ' ############################################################ GLOBAL

	Dim oWSH
	Dim oWMI
	Dim CreateSCHTask
	Dim CreateSCHTaskRemReg
    Dim CreateSCHTaskMouseRemReg
    Dim SmsClientVersion
    Dim SmsClientVerCheck
    Dim HasCursorKey : HasCursorKey = false

	' ############################################################ MAIN BEGIN

	PrintTitle("Initializing Objects")
	If (SetObjects = false) Then QuitScript( 100 )

	PrintTitle("Installing OSDResults")
	If (InstallOSDResults = false) Then QuitScript (200)

    PrintTitle("Checking CM Client Version")
	If (CheckClientVersion = false) Then QuitScript (300)

    	SmsClientVerCheck = (left(SmsClientVersion,1))

    	If (SmsClientVerCheck) >= 5 then
            PrintTitle("Installing AppInstall Component")
    	    If (InstallAppInstall = false) Then QuitScript (400)
    	End IF

    PrintTitle("Creating Cleanup Scheduled Tasks")
	If (UDICleanup = false) Then QuitScript (500)

	PrintTitle("Script Completed Successfully!")
	QuitScript ( 0 )

	' ############################################################ MAIN END

	' /////////////////////////////////////////////////////////
	' Install OSDResults
	' /////////////////////////////////////////////////////////
	Function InstallOSDResults()
		InstallOSDResults = true

		On Error Resume Next
		Err.Number = 0

		Call oWSH.run( CmdPeekRegBefore, 0, True )

        If IsTouchEnabled Then
            Call oWSH.run( CmdRegAddOSDResults, 0,True )
        else
            Call oWSH.run( CmdRegAddResults, 0,True )
        End If

		If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CmdRegAddResults & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			InstallOSDResults = false
		End If

		Err.Number = 0

		Call oWSH.run( CmdRegAddSetupType, 0,True )
		If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CmdRegAddSetupType & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			InstallOSDResults = false
		End If

		Call oWSH.run( CmdPeekRegAfter, 0, True )
		Call oWSH.run( CmdPeekTasks, 0, True )

        EnableMousePointer()

		On Error Goto 0

	End Function

    Function Reboot()
		Reboot = true

		On Error Resume Next
		Err.Number = 0
		wscript.echo "Executing: " & CmdReboot
		Call oWSH.run( CmdReboot, 0, True )
		If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CmdReboot & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			Reboot = false
		End If

		On Error Goto 0

	End Function

    ' /////////////////////////////////////////////////////////
	' Checking Configuration Manager Client Version
	' /////////////////////////////////////////////////////////
	Function CheckClientVersion()
		CheckClientVersion = true
        Dim SmsClient

		On Error Resume Next
		Err.Number = 0

        Set SmsClient = GetObject("winmgmts:ROOT/CCM:SMS_Client=@")
        SmsClientVersion = SmsClient.ClientVersion
        If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CmdRegAddAppInstall & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			CheckClientVersion = false
		End If
        SmsClientVersion = CInt(SmsClientVersion)
        wscript.echo "Client Version: " & SmsClientVersion

		On Error Goto 0

	End Function

	' /////////////////////////////////////////////////////////
	' Install App Install
	' /////////////////////////////////////////////////////////
	Function InstallAppInstall()
		InstallAppInstall = true

		On Error Resume Next
		Err.Number = 0

		Call oWSH.run( CmdRegAddAppInstall, 0,True )
		If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CmdRegAddAppInstall & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			InstallAppInstall = false
		End If

		On Error Goto 0

	End Function

    	' /////////////////////////////////////////////////////////
	' Scheduled task for UDI cleanup
	' /////////////////////////////////////////////////////////
	Function UDICleanup()
        Dim UDIRemovalDate
		UDICleanup = true

		On Error Resume Next
		Err.Number = 0
        UDIRemovalDate = CStr(DateAdd("m",1,date))

        '
        ' ensure the date is in mm/dd/yyyy format.
        ' this is req'd for schedtasks cmd.
        '
        dim mm, dd
        mm = Month(UDIRemovalDate)
        dd = Day(UDIRemovalDate)

        If len( mm ) < 2 Then
            mm = "0" & mm
        End If
        If len( dd ) < 2 Then
            dd = "0" & dd
        End If
        UDIRemovalDate = (mm & "/" & dd & "/" & Year(UDIRemovalDate))

		CreateSCHTask = "cmd /c " & SCHTask & " /SD " & UDIRemovalDate
        CreateSCHTaskRemReg = "cmd /c " & SCHTaskRegClean & " /SD " & UDIRemovalDate
        CreateSCHTaskMouseRemReg = "cmd /c " & SCHTaskRestoreMouse

		Call oWSH.run( CreateSCHTask, 0,True )
		If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CreateSCHTask & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			CreateSCHTask = false
            Err.Number = 0
		End If

       Call oWSH.run( CreateSCHTaskRemReg, 0,True )
		If ( Err.Number <> 0 ) Then
			wscript.echo " --| Command: [" & CreateSCHTaskRemReg & "]"
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
			CreateSCHTaskRemReg = false
            Err.Number = 0
		End If

        if (HasCursorKey = true) Then

            wscript.echo " Creating Task :" & CreateSCHTaskMouseRemReg
            Call oWSH.run( CreateSCHTaskMouseRemReg, 0,True )

		    If ( Err.Number <> 0 ) Then
			    wscript.echo " --| Command: [" & CreateSCHTaskMouseRemReg & "]"
			    wscript.echo " --| Error: [" & Err.Number & "]"
			    wscript.echo " --| Description: [" & Err.Description & "]"
			    CreateSCHTaskRemReg = false
                Err.Number = 0
		    End If

        End if

		On Error Goto 0

	End Function

	' /////////////////////////////////////////////////////////
	' Enables Mouse Pointer only if the registry exists, meaning we are in win 8
	' /////////////////////////////////////////////////////////
    Function EnableMousePointer()
        Dim strValue
        Dim oReg
        Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
        Wscript.Echo "Looking for key :" & CursorKey

        oReg.GetDWORDValue HKEY_LOCAL_MACHINE, CursorKey , CursorKeyValueName , strValue

        If strValue > 0 Then
            Wscript.Echo "Current WSH Trust Policy Value: " & strValue
            HasCursorKey = true
            oReg.SetDWORDValue HKEY_LOCAL_MACHINE, CursorKey, CursorKeyValueName, 0
        End If

    End Function

	' /////////////////////////////////////////////////////////
	' Scan Processes and Wait
	' /////////////////////////////////////////////////////////
	Function ScanProcessWait(whereClause)

		Dim WMICollection
		Dim WMIItem
		Dim processName
		Dim queryString
		Do
			queryString = "Select * from Win32_Process where " & whereClause
            Wscript.Echo queryString
			Set WMICollection = oWMI.ExecQuery(queryString)
			If (WMICollection.Count > 0) Then
				For Each WMIItem in WMICollection
					processName = WMIItem.Name
					Wscript.Echo " --| Waiting for the following process to exit [" & processName & "]"
				Next

				'Wscript.Echo " --| Process Found [" & ProcessWatchName & "], sleeping 0.5 seconds..."
				WScript.Sleep 500
			End If

		Loop While (WMICollection.Count > 0)

	End Function

    Function IsTouchEnabled
        IsTouchEnabled = false

        Dim queryString
        Dim PointingDevices
        Dim isTouchResult

        queryString = "Select * from Win32_PointingDevice"

        Wscript.Echo queryString
        Set PointingDevices = oWMI.ExecQuery(queryString)

        If(PointingDevices.Count > 1) Then
         WScript.Echo "Has more than one pointing device"
         Exit Function
        End IF

        isTouchResult =  oWSH.run( CmdIsTouchEnabled, 0, True )

        if(isTouchResult > 0) Then
            IsTouchEnabled = true
             WScript.Echo "Has Touch Screen Only"
        else
             WScript.Echo "Has Mouse Only"
        end if

    End Function

	' /////////////////////////////////////////////////////////
	' Initialize/Set Objects
	' /////////////////////////////////////////////////////////
	Function SetObjects
		SetObjects = true

		On Error Resume Next
		Err.Number = 0

		Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set oWSH = CreateObject("WScript.Shell")

		if (Err.Number <> 0) Then
			SetObjects = false
			wscript.echo " --| Error: [" & Err.Number & "]"
			wscript.echo " --| Description: [" & Err.Description & "]"
		End If

		On Error Goto 0

	End Function

	' /////////////////////////////////////////////////////////
	' Print Title
	' /////////////////////////////////////////////////////////
	Sub PrintTitle( theTitle )

		wscript.echo "--------------------------------"
		wscript.echo theTitle
		wscript.echo "--------------------------------"
		wscript.echo ""

	End Sub

	' /////////////////////////////////////////////////////////
	' Quit Script
	' /////////////////////////////////////////////////////////
	Sub QuitScript( theExitCode )

		wscript.echo "--------------------------------"
		wscript.echo " Exiting with [" & theExitCode & "]"
		wscript.echo "--------------------------------"

		wscript.Quit( theExitCode )

	End Sub
'' SIG '' Begin signature block
'' SIG '' MIIaWgYJKoZIhvcNAQcCoIIaSzCCGkcCAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFMAkEWaybjrK
'' SIG '' WLi2GtzBfc59k6V2oIIVNjCCBKkwggORoAMCAQICEzMA
'' SIG '' AACIWQ48UR/iamcAAQAAAIgwDQYJKoZIhvcNAQEFBQAw
'' SIG '' eTELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEjMCEGA1UEAxMaTWlj
'' SIG '' cm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EwHhcNMTIwNzI2
'' SIG '' MjA1MDQxWhcNMTMxMDI2MjA1MDQxWjCBgzELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjENMAsGA1UECxMETU9QUjEeMBwGA1UE
'' SIG '' AxMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMIIBIjANBgkq
'' SIG '' hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAs3R00II8h6ea
'' SIG '' 1I6yBEKAlyUu5EHOk2M2XxPytHiYgMYofsyKE+89N4w7
'' SIG '' CaDYFMVcXtipHX8BwbOYG1B37P7qfEXPf+EhDsWEyp8P
'' SIG '' a7MJOLd0xFcevvBIqHla3w6bHJqovMhStQxpj4TOcVV7
'' SIG '' /wkgv0B3NyEwdFuV33fLoOXBchIGPfLIVWyvwftqFifI
'' SIG '' 9bNh49nOGw8e9OTNTDRsPkcR5wIrXxR6BAf11z2L22d9
'' SIG '' Vz41622NAUCNGoeW4g93TIm6OJz7jgKR2yIP5dA2qbg3
'' SIG '' RdAq/JaNwWBxM6WIsfbCBDCHW8PXL7J5EdiLZWKiihFm
'' SIG '' XX5/BXpzih96heXNKBDRPQIDAQABo4IBHTCCARkwEwYD
'' SIG '' VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFCZbPltd
'' SIG '' ll/i93eIf15FU1ioLlu4MA4GA1UdDwEB/wQEAwIHgDAf
'' SIG '' BgNVHSMEGDAWgBTLEejK0rQWWAHJNy4zFha5TJoKHzBW
'' SIG '' BgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNDb2RT
'' SIG '' aWdQQ0FfMDgtMzEtMjAxMC5jcmwwWgYIKwYBBQUHAQEE
'' SIG '' TjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NlcnRzL01pY0NvZFNpZ1BDQV8w
'' SIG '' OC0zMS0yMDEwLmNydDANBgkqhkiG9w0BAQUFAAOCAQEA
'' SIG '' D95ASYiR0TE3o0Q4abJqK9SR+2iFrli7HgyPVvqZ18qX
'' SIG '' J0zohY55aSzkvZY/5XBml5UwZSmtxsqs9Q95qGe/afQP
'' SIG '' l+MKD7/ulnYpsiLQM8b/i0mtrrL9vyXq7ydQwOsZ+Bpk
'' SIG '' aqDhF1mv8c/sgaiJ6LHSFAbjam10UmTalpQqXGlrH+0F
'' SIG '' mRrc6GWqiBsVlRrTpFGW/VWV+GONnxQMsZ5/SgT/w2at
'' SIG '' Cq+upN5j+vDqw7Oy64fbxTittnPSeGTq7CFbazvWRCL0
'' SIG '' gVKlK0MpiwyhKnGCQsurG37Upaet9973RprOQznoKlPt
'' SIG '' z0Dkd4hCv0cW4KU2au+nGo06PTME9iUgIzCCBLowggOi
'' SIG '' oAMCAQICCmECjkIAAAAAAB8wDQYJKoZIhvcNAQEFBQAw
'' SIG '' dzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBMB4XDTEyMDEwOTIy
'' SIG '' MjU1OFoXDTEzMDQwOTIyMjU1OFowgbMxCzAJBgNVBAYT
'' SIG '' AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
'' SIG '' EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
'' SIG '' cG9yYXRpb24xDTALBgNVBAsTBE1PUFIxJzAlBgNVBAsT
'' SIG '' Hm5DaXBoZXIgRFNFIEVTTjpGNTI4LTM3NzctOEE3NjEl
'' SIG '' MCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
'' SIG '' dmljZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoC
'' SIG '' ggEBAJbsjkdNVMJclYDXTgs9v5dDw0vjYGcRLwFNDNjR
'' SIG '' Ri8QQN4LpFBSEogLQ3otP+5IbmbHkeYDym7sealqI5vN
'' SIG '' Yp7NaqQ/56ND/2JHobS6RPrfQMGFVH7ooKcsQyObUh8y
'' SIG '' NfT+mlafjWN3ezCeCjOFchvKSsjMJc3bXREux7CM8Y9D
'' SIG '' SEcFtXogC+Xz78G69LPYzTiP+yGqPQpthRfQyueGA8Az
'' SIG '' g7UlxMxanMTD2mIlTVMlFGGP+xvg7PdHxoBF5jVTIzZ3
'' SIG '' yrDdmCs5wHU1D92BTCE9djDFsrBlcylIJ9jC0rCER7t4
'' SIG '' utV0A97XSxn3U9542ob3YYgmM7RHxqBUiBUrLHUCAwEA
'' SIG '' AaOCAQkwggEFMB0GA1UdDgQWBBQv6EbIaNNuT7Ig0N6J
'' SIG '' TvFH7kjB8jAfBgNVHSMEGDAWgBQjNPjZUkZwCu1A+3b7
'' SIG '' syuwwzWzDzBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8v
'' SIG '' Y3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0
'' SIG '' cy9NaWNyb3NvZnRUaW1lU3RhbXBQQ0EuY3JsMFgGCCsG
'' SIG '' AQUFBwEBBEwwSjBIBggrBgEFBQcwAoY8aHR0cDovL3d3
'' SIG '' dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3Nv
'' SIG '' ZnRUaW1lU3RhbXBQQ0EuY3J0MBMGA1UdJQQMMAoGCCsG
'' SIG '' AQUFBwMIMA0GCSqGSIb3DQEBBQUAA4IBAQBz/30unc2N
'' SIG '' iCt8feNeFXHpaGLwCLZDVsRcSi1o2PlIEZHzEZyF7BLU
'' SIG '' VKB1qTihWX917sb1NNhUpOLQzHyXq5N1MJcHHQRTLDZ/
'' SIG '' f/FAHgybgOISCiA6McAHdWfg+jSc7Ij7VxzlWGIgkEUv
'' SIG '' XUWpyI6zfHJtECfFS9hvoqgSs201I2f6LNslLbldsR4F
'' SIG '' 50MoPpwFdnfxJd4FRxlt3kmFodpKSwhGITWodTZMt7MI
'' SIG '' qt+3K9m+Kmr93zUXzD8Mx90Gz06UJGMgCy4krl9DRBJ6
'' SIG '' XN0326RFs5E6Eld940fGZtPPnEZW9EwHseAMqtX21Tyi
'' SIG '' 4LXU+Bx+BFUQaxj0kc1Rp5VlMIIFvDCCA6SgAwIBAgIK
'' SIG '' YTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYK
'' SIG '' CZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJ
'' SIG '' bWljcm9zb2Z0MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9v
'' SIG '' dCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
'' SIG '' MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29k
'' SIG '' ZSBTaWduaW5nIFBDQTCCASIwDQYJKoZIhvcNAQEBBQAD
'' SIG '' ggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBCmXZTbD4b
'' SIG '' 1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1
'' SIG '' VwqJyq4gSfTwaKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNE
'' SIG '' f9eB2/O98jakyVxF3K+tPeAoaJcap6Vyc1bxF5Tk/TWU
'' SIG '' cqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2
'' SIG '' L1TdFDBZ+NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049
'' SIG '' oDI9kM2hOAaFXE5WgigqBTK3S9dPY+fSLWLxRT3nrAgA
'' SIG '' 9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzB
'' SIG '' rAlfA9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMB
'' SIG '' Af8wHQYDVR0OBBYEFMsR6MrStBZYAck3LjMWFrlMmgof
'' SIG '' MAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
'' SIG '' MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+m
'' SIG '' PLzYLTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTAf
'' SIG '' BgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnkpDBQ
'' SIG '' BgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jv
'' SIG '' c29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9taWNyb3Nv
'' SIG '' ZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEESDBGMEQG
'' SIG '' CCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
'' SIG '' b20vcGtpL2NlcnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNy
'' SIG '' dDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+fyZGr+tvQLEy
'' SIG '' tWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAh
'' SIG '' Xaw9L0y6oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6Wc
'' SIG '' IC36C1DEVs0t40rSvHDnqA2iA6VW4LiKS1fylUKc8fPv
'' SIG '' 7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WR
'' SIG '' Tsgb0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZE
'' SIG '' olHN+emjWFbdmwJFRC9f9Nqu1IIybvyklRPk62nnqaIs
'' SIG '' vsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
'' SIG '' NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxP
'' SIG '' bQ4TTj18KUicctHzbMrB7HCjV5JXfZSNoBtIA1r3z6Nn
'' SIG '' CnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDordEN5k
'' SIG '' 9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7
'' SIG '' B2BHZWBATrBC7E7ts3Z52Ao0CW0cgDEf4g5U3eWh++VH
'' SIG '' EK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jshrg1cnPCi
'' SIG '' roZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONx
'' SIG '' PdcAfmJH0c6IybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3
'' SIG '' UAjWwz0wggYHMIID76ADAgECAgphFmg0AAAAAAAcMA0G
'' SIG '' CSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNj
'' SIG '' b20xGTAXBgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTAr
'' SIG '' BgNVBAMTJE1pY3Jvc29mdCBSb290IENlcnRpZmljYXRl
'' SIG '' IEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0
'' SIG '' MDMxMzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
'' SIG '' EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
'' SIG '' HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
'' SIG '' BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCC
'' SIG '' ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJ+h
'' SIG '' bLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn0Uyt
'' SIG '' dDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4E
'' SIG '' mPCJzB/LMySHnfL0Zxws/HvniB3q506jocEjU8qN+kXP
'' SIG '' CdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4nrIZPVVIM
'' SIG '' 5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0
'' SIG '' RZCfSABKR2YRJylmqJfk0waBSqL5hKcRRxQJgp+E7VV4
'' SIG '' /gGaHVAIhQAQMEbtt94jRrvELVSfrx54QTF3zJvfO4OT
'' SIG '' oWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAasw
'' SIG '' ggGnMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0
'' SIG '' +NlSRnAK7UD7dvuzK7DDNbMPMAsGA1UdDwQEAwIBhjAQ
'' SIG '' BgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQO
'' SIG '' rIJgQFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmS
'' SIG '' JomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1p
'' SIG '' Y3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
'' SIG '' Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxz
'' SIG '' WPQHEy5lMFAGA1UdHwRJMEcwRaBDoEGGP2h0dHA6Ly9j
'' SIG '' cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
'' SIG '' L21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcB
'' SIG '' AQRIMEYwRAYIKwYBBQUHMAKGOGh0dHA6Ly93d3cubWlj
'' SIG '' cm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0Um9v
'' SIG '' dENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0G
'' SIG '' CSqGSIb3DQEBBQUAA4ICAQAQl4rDXANENt3ptK132855
'' SIG '' UU0BsS50cVttDBOrzr57j7gu1BKijG1iuFcCy04gE1CZ
'' SIG '' 3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji
'' SIG '' 8FMV3U+rkuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZ
'' SIG '' Lg33B+JwvBhOnY5rCnKVuKE5nGctxVEO6mJcPxaYiyA/
'' SIG '' 4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tO
'' SIG '' i3/FNSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLa
'' SIG '' FJj1PLlmWLMtL+f5hYbMUVbonXCUbKw5TNT2eb+qGHpi
'' SIG '' Ke+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
'' SIG '' NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCe
'' SIG '' FTBm6EISXhrIniIh0EPpK+m79EjMLNTYMoBMJipIJF9a
'' SIG '' 6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2JoXZh
'' SIG '' tG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0t
'' SIG '' r1mPuOQh5bWwymO0eFQF1EEuUKyUsKV4q7OglnUa2ZKH
'' SIG '' E3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng9wFlb4kL
'' SIG '' fchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj
'' SIG '' /TGCBJAwggSMAgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25p
'' SIG '' bmcgUENBAhMzAAAAiFkOPFEf4mpnAAEAAACIMAkGBSsO
'' SIG '' AwIaBQCggbIwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
'' SIG '' AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUw
'' SIG '' IwYJKoZIhvcNAQkEMRYEFAR4annhAXHQqlOLkb2xZ+/L
'' SIG '' yg8+MFIGCisGAQQBgjcCAQwxRDBCoCSAIgBNAEQAVAAg
'' SIG '' AFUARABJAHYAMwAgAFQAbwBvAGwAawBpAHShGoAYaHR0
'' SIG '' cDovL3d3dy5taWNyb3NvZnQuY29tMA0GCSqGSIb3DQEB
'' SIG '' AQUABIIBAISaWb56BlgbbPJKLcUqYLwYUQvHsHru2lld
'' SIG '' NSwA8k9t1Vw7C9mGkf62B3lCCfKvWg7t55DUZlUWQyqW
'' SIG '' tLeUva6PN+Rsr0fC0438M/515M3OWAySJj/GhPy9BsnB
'' SIG '' HncUfCa4aNxPmTOUIpBta+GR8tMaQ+Vkv3ZbxkaXLj5w
'' SIG '' krnGpTMoS4qCxWWMw5sHS93ibZjl6wwdyeXM5ILwW6fa
'' SIG '' FNM4OMoP26asKbItt3kiWmFa63uvWR1ATKcEUOyEgNv/
'' SIG '' gNwvdMVTzCdtSbdttupN+I8jmpCvJ8oN90/6aEt8haUe
'' SIG '' bfTskktFjo8qo2YFZn9fBX187h3bpq6loXuMQgNq7f6h
'' SIG '' ggIfMIICGwYJKoZIhvcNAQkGMYICDDCCAggCAQEwgYUw
'' SIG '' dzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBAgphAo5CAAAAAAAf
'' SIG '' MAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZI
'' SIG '' hvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xMjA4MzAxOTU3
'' SIG '' MDJaMCMGCSqGSIb3DQEJBDEWBBRMaCT50foAWISp7b/n
'' SIG '' LanG5RsBozANBgkqhkiG9w0BAQUFAASCAQA3hi0IxQh+
'' SIG '' wxHvml0RDxjjwV7BpveE8sgIFD8000jH4SNRnlvVumAc
'' SIG '' Luu01vHJcAmxJY6FkwUWWbq3r72FM4p8ftXpvLi0xDuT
'' SIG '' BAWa+ZJsfLW+ewN/NNKuLfL2EygpQLTGR6kcy5sDhXTW
'' SIG '' puAIhfRfekIQZMSKW1EEtKWwaitBKflWZ1R6C7YEkTEo
'' SIG '' yMbzzMrQsodOPU8SBLeuIz4IrkyEY4mzc2hYCJ6Tgcml
'' SIG '' 53R9+XLVKmRuodLGri6YuUP8wGR4ouqPkHP1mMOELfbN
'' SIG '' kxMcbC36j/GkBK0Oa5zhq2jJx7fINPotcli+5T1iyNYi
'' SIG '' JGYV8b3SbRvuN6zJvQChTCEt
'' SIG '' End signature block
