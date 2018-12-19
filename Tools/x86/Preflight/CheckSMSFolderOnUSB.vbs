    '///////////////////////////////////////////////////////
    '// CheckSMSFolderOnUSB
    '///////////////////////////////////////////////////////
    '// Scan for _SMSTaskSequence folder on all USB connected drives
    '///////////////////////////////////////////////////////
    Option Explicit

    Dim FSO, Path,query 
    Dim WMIServices, wmiDiskDrives, wmiDiskDrive, wmiDiskPartitions, wmiDiskPartition , wmiLogicalDisks, wmiLogicalDisk
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Set WMIServices = GetObject("winmgmts:\\.\root\cimv2")
    Set wmiDiskDrives = wmiServices.ExecQuery ("SELECT DeviceID FROM Win32_DiskDrive WHERE InterfaceType='USB'")

    For Each wmiDiskDrive In wmiDiskDrives
    '    'Use the disk drive device id to find associated partition
      query = "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & wmiDiskDrive.DeviceID & "'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"    
      Set wmiDiskPartitions = wmiServices.ExecQuery(query)

        For Each wmiDiskPartition In wmiDiskPartitions
            'Use partition device id to find logical disk
            Set wmiLogicalDisks = wmiServices.ExecQuery _
                ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" _
                & wmiDiskPartition.DeviceID & "'} WHERE AssocClass = Win32_LogicalDiskToPartition") 

            For Each wmiLogicalDisk In wmiLogicalDisks
               'WScript.Echo  "Drive Letter: " & wmiLogicalDisk.DeviceID
                Path = wmiLogicalDisk.DeviceID & "\_SMSTaskSequence"
                If FSO.FolderExists(Path) Then 
                    'Problem: Folder exists. Quit with non zero
                    WScript.Quit 1
                End IF
            Next      
        Next
    Next
    'We are Good
    WScript.Quit 0