Attribute VB_Name = "modDeadCode"
'---------------------------------------------------------------------------------------
' Module    : DeadCode
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : GetExePathFromPID
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : getting the full path of a running process is not as easy as you'd expect
'---------------------------------------------------------------------------------------
'
'Public Function GetExePathFromPID(ByVal idProc As Long) As String
'    Dim sBuf As String
'    Dim sChar As Long
'    Dim useloop As Integer
'    Dim hProcess As Long
'
'    On Error GoTo GetExePathFromPID_Error
'
'    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, idProc)
'    If hProcess Then
'        sBuf = String$(260, vbNullChar)
'        sChar = GetProcessImageFileName(hProcess, sBuf, 260)
'        If sChar Then
'            sBuf = NoNulls(sBuf)
'            ' this loop replaces the internal windows volume name with the legacy naming convention, ie. C:\, D:\ &c
'            For useloop = 1 To lstDevicesListCount
'                If InStr(1, sBuf, lstDevices(1, useloop)) > 0 Then
'                    sBuf = Replace$(sBuf, lstDevices(1, useloop), Chr$(lstDevices(0, useloop)) & ":")
'                    Exit For
'                End If
'            Next useloop
'            GetExePathFromPID = sBuf
'        End If
'        CloseHandle hProcess
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'GetExePathFromPID_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetExePathFromPID of Module common"
'End Function

'---------------------------------------------------------------------------------------
' Procedure : NoNulls
' Author    : beededea
' Date      : 25/08/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Public Function NoNulls(ByVal Strng As String) As String
'    Dim I As Integer
'   On Error GoTo NoNulls_Error
'
'    If Len(Strng) > 0 Then
'        I = InStr(Strng, vbNullChar)
'        Select Case I
'            Case 0
'                NoNulls = Strng
'            Case 1
'                NoNulls = vbNullString
'            Case Else
'                NoNulls = Left$(Strng, I - 1)
'        End Select
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'NoNulls_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NoNulls of Module common"
'End Function




'---------------------------------------------------------------------------------------
' Procedure : ShowDevices
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : put the device names in an accessible list so that they can be mapped later
'             used especially to obtain the unexpectedly hard-to-extract default folder name of a process in the function fIsRunning
'---------------------------------------------------------------------------------------
'
'Public Sub ShowDevices(sDriveStrings As String)
'    Dim vDrive As Variant ' probably handling this already in .NET
'    Dim sDeviceName As String
'    Dim thiskey As String
'    Dim driveCount As Integer
'
'    driveCount = 0
'
'    On Error GoTo ShowDevices_Error
'
'    For Each vDrive In GetDrives(sDriveStrings) ' getdrives is a collection of drive name strings C:\, D:\ &c
'        sDeviceName = GetNtDeviceNameForDrive(vDrive) ' \Device\HarddiskVolume1 are the default naming conventions for Windows drives
'        driveCount = driveCount + 1
'
'        lstDevices(0, driveCount) = Asc(Mid$(vDrive, 1, 1))
'        lstDevices(1, driveCount) = sDeviceName
'
'    Next
'
'    lstDevicesListCount = driveCount ' global variable
'
'   On Error GoTo 0
'   Exit Sub
'
'ShowDevices_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShowDevices of Module common"
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDrives
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : getdrives returns a collection of drive name strings C:\, D:\ &c
'---------------------------------------------------------------------------------------
'
'Public Function GetDrives(ByRef sDriveStrings As String) As Collection
'
'    Dim colDrives As New Collection
'    Dim lSize As Long
'    Dim lR As Long
'    Dim iLastPos As Long
'    Dim iPos As Long
'    Dim sDrive As String
'
'   On Error GoTo GetDrives_Error
'
'   lSize = GetLogicalDriveStringsA(0, ByVal 0&)
'   sDriveStrings = String(lSize + 1, 0)
'   lR = GetLogicalDriveStringsA(lSize, ByVal sDriveStrings)
'   iLastPos = 1
'   Do
'      iPos = InStr(iLastPos, sDriveStrings, vbNullChar)
'      If Not (iPos = 0) Then
'         sDrive = Mid$(sDriveStrings, iLastPos, iPos - iLastPos)
'         iLastPos = iPos + 1
'      Else
'         sDrive = Mid$(sDriveStrings, iLastPos)
'      End If
'      If Len(sDrive) > 0 Then
'         colDrives.Add sDrive
'      End If
'   Loop While Not (iPos = 0)
'   Set GetDrives = colDrives
'
'   On Error GoTo 0
'   Exit Function
'
'GetDrives_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDrives of Module common"
'
'End Function
    
'---------------------------------------------------------------------------------------
' Procedure : GetNtDeviceNameForDrive
' Author    : beededea
' Date      : 25/08/2020
' Purpose   : \Device\HarddiskVolume1 are the default naming conventions for Windows drives
'---------------------------------------------------------------------------------------
'
'Public Function GetNtDeviceNameForDrive(ByVal sDrive As String) As String
'
'    Dim bDrive() As Byte
'    Dim bResult() As Byte
'    Dim lR As Long
'    Dim sDeviceName As String
'
'   On Error GoTo GetNtDeviceNameForDrive_Error
'
'   If Right$(sDrive, 1) = "\" Then
'      If Len(sDrive) > 1 Then
'         sDrive = Left$(sDrive, Len(sDrive) - 1)
'      End If
'   End If
'   bDrive = sDrive
'
'   ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
'   ReDim bResult(0 To 260 * 2 + 1) As Byte
'
'   lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bResult(0)), 260)
'   If (lR > 2) Then
'      sDeviceName = bResult
'      sDeviceName = Left$(sDeviceName, lR - 2)
'      GetNtDeviceNameForDrive = sDeviceName
'   End If
'
'   On Error GoTo 0
'   Exit Function
'
'GetNtDeviceNameForDrive_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetNtDeviceNameForDrive of Module common"
'
'End Function

'Public Function GetUDTDateTime() As String
'    Const TIME_ZONE_ID_DAYLIGHT As Long = 2
'    Dim tzi As TIME_ZONE_INFORMATION
'    Dim dwBias As Long
'    Dim sZone As String
'    Dim tmp As String
'    Select Case GetTimeZoneInformation(tzi)
'        Case TIME_ZONE_ID_DAYLIGHT
'            dwBias = tzi.Bias + tzi.DaylightBias
'            sZone = " " & Left$(tzi.DaylightName, 1) & "DT"
'        Case Else
'            dwBias = tzi.Bias + tzi.StandardBias
'            sZone = " " & Left$(tzi.StandardName, 1) & "ST"
'    End Select
'    tmp = "  " & Right$("00" & CStr(dwBias \ 60), 2) & Right$("00" & CStr(dwBias Mod 60), 2) & sZone
'    If dwBias > 0 Then
'        Mid$(tmp, 2, 1) = "-"
'    Else
'        Mid$(tmp, 2, 2) = "+0"
'    End If
'    GetUDTDateTime = Format$(Now, "ddd, dd mmm yyyy Hh:Mm:Ss") & tmp
'End Function



'    timer code STARTS
'    Dim lngReturn As Long
'    Dim curFreq As Currency
'    Dim curStart As Currency
'    Dim curEnd As Currency
'    Dim sngTime As Single
'
'    lngReturn = QueryPerformanceFrequency(curFreq)
'    If lngReturn = 0 Then MsgBox "Your Hardware does not support a high-resolution timer"
'
'    lngReturn = QueryPerformanceCounter(curStart)
'    timer code ENDS
    
'    Dim aFile As Object
'    Dim bFile As Object
'    Dim cFile As Object
'    Dim lineToRead As String
'    Dim nextLine As String
'    Dim test As Boolean
'    Dim findCount As Integer
'    Dim duplicateCount As Integer
'
'
'        Const ForWriting As Integer = 2
'
'    'test = False
'    test = True
'
'
'    If test = True Then
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        Set cFile = fso.OpenTextFile("E:\dropbox sync\Dropbox\yahoo widgets\harry read this pleaseNoDup.txt", ForWriting, 0)
'
'        Set aFile = fso.OpenTextFile(sFName, forReading, False, 0)
'        Do While Not aFile.AtEndOfStream
'            lineToRead = aFile.readLine
'
'            Set bFile = fso.OpenTextFile("E:\dropbox sync\Dropbox\yahoo widgets\harry read this pleaseDup.txt", forReading, False, 0)
'            Do While Not bFile.AtEndOfStream
'                nextLine = bFile.readLine
'                If lineToRead = nextLine Then
'                    findCount = findCount + 1
'                    If findCount >= 2 Then
''                        Dim a As Integer
''                        a = 1 ' something to breakpoint
''                        duplicateCount = duplicateCount + 1
''                        Exit Do
'                    End If
'                End If
'            Loop
'            bFile.Close
'            If findCount >= 2 Then
'                cFile.Write "Duplicate " & findCount & " " & lineToRead & vbCrLf
'            Else
'                cFile.Write lineToRead & vbCrLf
'            End If
'            findCount = 0
'        Loop
'        aFile.Close
'        cFile.Close
'        MsgBox "Duplicates Found " & duplicateCount
'
'     End If
