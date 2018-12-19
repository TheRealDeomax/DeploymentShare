    '///////////////////////////////////////////////////////
    '// OSD_GetPrinters
    '///////////////////////////////////////////////////////
    '// Scan all profiles and capture printers
    '// to XML file.
    '///////////////////////////////////////////////////////
    Option Explicit

    '///////////////////////////////////////////////////////
    '// Constants
    '///////////////////////////////////////////////////////
    Const HKEY_USERS = &H80000003

    '///////////////////////////////////////////////////////
    '// Globals
    '///////////////////////////////////////////////////////
    Dim XmlPrinterEntry
    Dim PrinterStart
    Dim PrinterEnd
    Dim XmlNetworkEntry
    Dim NetworkStart
    Dim NetworkEnd
    Dim XmlUserStart
    Dim XmlUserEnd
    Dim XmlDocRoot
    Dim XmlFileName
    Dim XmlFileName2

        XmlPrinterEntry    = "<Printer Path=""[--PATH--]"" Name=""[--NAME--]"" />"
        PrinterStart       = "<Printers>"
        PrinterEnd         = "</Printers>"

        XmlNetworkEntry    = "<NetworkShare Path=""[--PATH--]"" Drive=""[--NAME--]"" />"
        NetworkStart       = "<NetwrokShares>"
        NetworkEnd         = "</NetwrokShares>"

        XmlUserStart       = "<User SID=""[--SID--]"" >"
        XmlUserEnd         = "</User>"
        XmlDocRoot         = "<Root>[--USERPRINTERCONTENT--]</Root>"
        XmlFileName        = ""
        XmlFileName2       = ""

    Dim oREG
    Set oREG = GetObject("winmgmts:\\.\root\default:StdRegProv")

    Dim oFSO
    Dim oFSOFileOut
    Set oFSO = CreateObject("Scripting.FileSystemObject")

'///////////////////////////////////////////////////////
'// Start / Main
'///////////////////////////////////////////////////////

    '||||||||||||||||||||||||||||||
    '| Check Arguments
    '||||||||||||||||||||||||||||||
    If ( GoCheckArguments() = FALSE ) Then QuitScript( 10 )
    If ( GoCreateFile() = FALSE ) Then QuitScript ( 20 )

    '||||||||||||||||||||||||||||||
    '| Parse Printers
    '||||||||||||||||||||||||||||||
    GoParsePrintersNetworks()

    '||||||||||||||||||||||||||||||
    '| Write Data to File
    '||||||||||||||||||||||||||||||
    oFSOFileOut.WriteLine( XmlDocRoot )
    oFSOFileOut.Close

    if ( len(XmlFileName2) > 1 ) then

           On Error Resume Next

            oFSO.CopyFile XmlFileName, XmlFileName2, true

            On Error Goto 0
    end if

    '||||||||||||||||||||||||||||||
    '| Exit
    '||||||||||||||||||||||||||||||
    QuitScript( 0 )

'///////////////////////////////////////////////////////
'// End / Main
'///////////////////////////////////////////////////////

    '///////////////////////////////////////////////////////
    '// GoCreateFile
    '-------------------------------------------------------
    '// Creates file (leaves file open!)
    '///////////////////////////////////////////////////////
    Function GoCreateFile()
        Dim returnValue

        returnValue = TRUE

        '||||||||||||||||||||||||||||||
        '| Test File Open
        '||||||||||||||||||||||||||||||
        Err.Clear()
        On Error Resume Next

        Set oFSOFileOut = oFSO.CreateTextFile( XmlFileName, TRUE )

        If (Err.number <> 0) Then

            Wscript.Echo "Error while creating file [" & XmlFileName & "]!"
            GoWriteLastError()
            returnValue = FALSE

        Else

            oFSOFileOut.WriteLine( "" )

        End If

        On Error Goto 0

        '||||||||||||||||||||||||||||||
        '| Write a line and return
        '||||||||||||||||||||||||||||||

        GoCreateFile = returnValue

    End Function

    '///////////////////////////////////////////////////////
    '// GoCheckArguments
    '-------------------------------------------------------
    '// Ensure a valid filename is given
    '///////////////////////////////////////////////////////
    Function GoCheckArguments()
        Dim returnValue
        Dim fileArg2
        Dim fileArg

        returnValue = TRUE

        '||||||||||||||||||||||||||||||
        '| Test File Name Argument
        '||||||||||||||||||||||||||||||
        If (returnValue = TRUE) Then If ( Wscript.Arguments.Count <> 1 ) Then returnValue = FALSE
        If (returnValue = TRUE) Then If ( Len( Wscript.Arguments(0) ) < 13 ) Then returnValue = FALSE
        If (returnValue = TRUE) Then If ( InStr( 1, Wscript.Arguments(0), "/xmlout:", 1 ) <> 1 ) Then returnValue = FALSE

        '||||||||||||||||||||||||||||||
        '| Assign Filename
        '||||||||||||||||||||||||||||||

        If (returnValue = TRUE) Then

            fileArg = Wscript.Arguments(0)

            Dim fileArray : fileArray = Split( fileArg, ";" )

            if( UBound(fileArray) > 0 ) Then
                fileArg = fileArray(0)
                fileArg2 = fileArray(1)

                XmlFileName2 = fileArg2

                Wscript.Echo XmlFileName2

            end if

            XmlFileName = Right( fileArg, (Len(fileArg)-8))

            Wscript.Echo XmlFileName

        End IF

        GoCheckArguments = returnValue

    End Function

    '///////////////////////////////////////////////////////
    '// GoWriteLastError
    '-------------------------------------------------------
    '// Outputs error details
    '///////////////////////////////////////////////////////
    Sub GoWriteLastError()

        Wscript.Echo ""
        Wscript.Echo " - Error Number: [" & Err.number & "]"
        Wscript.Echo " - Error Description: [" & Err.Description & "]"
        Wscript.Echo ""

    End Sub

    '///////////////////////////////////////////////////////
    '// QuitScript
    '-------------------------------------------------------
    '// Exit Script
    '///////////////////////////////////////////////////////
    Sub QuitScript( exitCode )

        Select Case ( exitCode )

        Case 0  Wscript.Echo "Exiting successfully"
        Case 10 Wscript.Echo "Please provide an output filename ( /xmlout:File.xml )"
        Case 20 Wscript.Echo "Unable to create output file [" & XmlFileName & "]"

        End Select

        Wscript.Quit( exitCode )

    End Sub

    '///////////////////////////////////////////////////////
    '// GoParsePrinters
    '-------------------------------------------------------
    '// Parse printers for all users
    '///////////////////////////////////////////////////////
    Sub GoParsePrintersNetworks()

        Dim enumSubKeys
        Dim itemSubKey

        Dim RootNodeText
            RootNodeText = ""

        oREG.EnumKey HKEY_USERS, "", enumSubKeys

        '||||||||||||||||||||||||||||||
        '| Loop Users
        '||||||||||||||||||||||||||||||
        For Each itemSubKey In enumSubKeys

            Wscript.Echo "Detected user: [" & itemSubKey & "]"

            Dim UserNodeText

            Dim enumPrintKeys
            Dim itemPrintKey

            Dim regPrintPath

            Dim EnumSuccess
            Dim printersFound

            UserNodeText = ""
            printersFound = False
            UserNodeText = UserNodeText & XmlUserStart

            regPrintPath = itemSubKey & "\Printers\Connections"

            EnumSuccess = oREG.EnumKey (HKEY_USERS, regPrintPath, enumPrintKeys)

           '||||||||||||||||||||||||||||||
           '| If any items enumerated...
           '||||||||||||||||||||||||||||||
            If ( (EnumSuccess = 0) And (IsArray(enumPrintKeys)) ) Then

                For Each itemPrintKey In enumPrintKeys
                    Wscript.Echo " .. Detected possible printer entry: [" & itemPrintKey & "]"

                   '||||||||||||||||||||||||||||||
                   '| Read Printer Properties
                   '||||||||||||||||||||||||||||||
                   Dim printerProperties
                   Dim printerName
                   Dim printerPath

                   printerProperties = Split( itemPrintKey, "," )

                   '||||||||||||||||||||||||||||||
                   '| If Printer entry has properties...
                   '||||||||||||||||||||||||||||||
                   if ( UBound( printerProperties ) > 2 ) Then

                       printerPath = printerProperties(2)
                       printerName = printerProperties(3)

                       Dim newXmlPrinter

                       newXmlPrinter = ""
                       newXmlPrinter = newXmlPrinter & XmlPrinterEntry

                       newXmlPrinter = Replace( newXmlPrinter, "[--PATH--]", printerPath )
                       newXmlPrinter = Replace( newXmlPrinter, "[--NAME--]", printerName )

                       '||||||||||||||||||||||||||||||
                       '| Add Printer to User Node
                       '||||||||||||||||||||||||||||||
                       Wscript.Echo " .. Adding printer entry: [" & newXmlPrinter & "] [" & printerName & "]"
                       If (printersFound = False) Then
                        UserNodeText = UserNodeText & PrinterStart
                       End IF
                       UserNodeText = UserNodeText & newXmlPrinter
                       printersFound = true

                   End If

                Next

            End If

            If (printersFound = true) Then
                UserNodeText = UserNodeText & PrinterEnd
            End If

            '||||||||||||||||||||||||||||||||||||||||||||||||
            '| Scan Networkshares
            '||||||||||||||||||||||||||||||||||||||||||||||||

            Dim enumNetworkKeys
            Dim itemNetwrokKey

            Dim regNetworkPath

            Dim EnumNWSuccess
            Dim networksFound

            networksFound = False
            regNetworkPath = itemSubKey & "\Network"

            EnumNWSuccess = oREG.EnumKey (HKEY_USERS, regNetworkPath, enumNetworkKeys)

           '||||||||||||||||||||||||||||||
           '| If any items enumerated...
           '||||||||||||||||||||||||||||||
            If ( (EnumNWSuccess = 0) And (IsArray(enumNetworkKeys)) ) Then

                For Each itemNetwrokKey In enumNetworkKeys
                    Wscript.Echo " .. Detected possible Network entry: [" & itemNetwrokKey & "]"

                   '||||||||||||||||||||||||||||||
                   '| Read Networkshare Properties
                   '||||||||||||||||||||||||||||||
                   Dim networkDrivePath
                   Dim networkDrive
                   Dim networkPath
                   Dim strValueName
		           Dim remotePathFound

                   networkPath = ""
                   networkDrive = itemNetwrokKey
                   networkDrivePath = regNetworkPath & "\" & networkDrive
                   strValueName = "RemotePath"

		           Dim newXmlNetwork

                   newXmlNetwork = ""
                   newXmlNetwork = newXmlNetwork & XmlNetworkEntry
                   newXmlNetwork = Replace( newXmlNetwork, "[--NAME--]", networkDrive )

                   remotePathFound = oReg.GetStringValue (HKEY_USERS,networkDrivePath,strValueName,networkPath )

		           If (remotePathFound = 0) Then
			            newXmlNetwork = Replace( newXmlNetwork, "[--PATH--]", networkPath )
		           Else
			            newXmlNetwork = Replace( newXmlNetwork, "[--PATH--]", "")
                   End IF

                    '||||||||||||||||||||||||||||||
                    '| Add Networkshare to User Node
                    '||||||||||||||||||||||||||||||
                    Wscript.Echo " .. Adding printer entry: [" & networkPath & "] [" & networkDrive & "]"
                    If (networksFound = False) Then
                        UserNodeText = UserNodeText & NetworkStart
                    End IF

                    UserNodeText = UserNodeText & newXmlNetwork
                    networksFound = true
                Next
            End If

            If (networksFound = true) Then
                UserNodeText = UserNodeText & NetworkEnd
            End If

           '||||||||||||||||||||||||||||||
           '| Add User Node to Document
           '||||||||||||||||||||||||||||||
            UserNodeText = Replace(UserNodeText,"[--SID--]", itemSubKey)
            UserNodeText = UserNodeText & XmlUserEnd

            RootNodeText = RootNodeText & UserNodeText

        Next

        XmlDocRoot = Replace( XmlDocRoot, "[--USERPRINTERCONTENT--]", RootNodeText )

    End Sub