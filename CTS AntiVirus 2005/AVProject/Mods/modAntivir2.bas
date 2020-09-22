Attribute VB_Name = "modAntivir2"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Function CalcCRC(ByVal strFileName As String) As String
Dim cCRC32  As New cCRC32
Dim lCRC32  As Long
Dim cStream As New cBinaryFileStream
    On Error GoTo err
    cStream.file = strFileName
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    CalcCRC = Hex$(lCRC32)
Exit Function
err:
    ErrorFunc err.Number, err.Description, "modAntivir.CalcCRC", strFileName
End Function
Public Sub CheckExe()
'    On Error GoTo Ignore
'    If GetSetting(AV.AVname, "Settings", "CRC", CalcCRC(App.path & "\" & App.EXEName & ".exe")) <> CalcCRC(App.path & "\" & App.EXEName & ".exe") Then
'        MsgBox LoadResString(139), vbCritical + vbOKOnly, LoadResString(140)
'        End
'    End If
    SaveSetting AV.AVname, "Settings", "CRC", CalcCRC(App.path & "\" & App.EXEName & ".exe")
Ignore:
End Sub
Public Sub Checkfolder(Optional ByVal StrFolder As String)
Dim Result As Variant
'Dim c      As Collection
    Debug.Print Time
    On Error Resume Next
    If StrFolder = vbNullString Then
        Set Result = SH.BrowseForFolder(frmMain.hwnd, LoadResString(141), 1)
        With Result.Items.Item
            AVE.ScanFolder .path, True
        End With
    Else
        AVE.ScanFolder StrFolder, True
    End If
    On Error GoTo 0
End Sub
Private Function CreateKey(lhKey As Long, _
                           SubKey As String, _
                           NewSubKey As String) As Boolean
Dim lhKeyOpen    As Long
Dim lhKeyNew     As Long
Dim lDisposition As Long
Dim lResult      As Long
Dim Security     As SECURITY_ATTRIBUTES
    lhKeyOpen = OpenKey(lhKey, SubKey, KEY_CREATE_SUB_KEY)
    lResult = RegCreateKeyEx(lhKeyOpen, NewSubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, Security, lhKeyNew, lDisposition)
    If lResult = ERROR_SUCCESS Then
        CreateKey = True
        RegCloseKey (lhKeyNew)
    Else
        CreateKey = False
    End If
    RegCloseKey (lhKeyOpen)
End Function
Private Function OpenKey(lhKey As Long, _
                         SubKey As String, _
                         ulOptions As Long) As Long
Dim lhKeyOpen As Long
Dim lResult   As Long
    lhKeyOpen = 0
    lResult = RegOpenKeyEx(lhKey, SubKey, 0, ulOptions, lhKeyOpen)
    If lResult <> ERROR_SUCCESS Then
        OpenKey = 0
    Else
        OpenKey = lhKeyOpen
    End If
End Function
Public Function RegisterFile(sFileExt As String, _
                             sFileDescr As String, _
                             sAppID As String, _
                             sOpenCmd As String, _
                             sIconFile As String) As Boolean
Dim hKey      As Long
Dim bSuccess  As Boolean
Dim bSuccess2 As Boolean
    bSuccess = False
    hKey = HKEY_LOCAL_MACHINE
    If CreateKey(hKey, REG_PRIMARY_KEY, sFileExt) Then
        If SetValue(hKey, REG_PRIMARY_KEY & sFileExt, sAppID) Then
            If CreateKey(hKey, REG_PRIMARY_KEY, sAppID) Then
                If SetValue(hKey, REG_PRIMARY_KEY & sAppID, sFileDescr) Then
                    If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY) Then
                        bSuccess = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_SHELL_KEY & REG_SHELL_OPEN_KEY & REG_SHELL_OPEN_COMMAND_KEY, sOpenCmd)
                        If CreateKey(hKey, REG_PRIMARY_KEY & sAppID, REG_ICON_KEY) Then
                            bSuccess2 = SetValue(hKey, REG_PRIMARY_KEY & sAppID & "\" & REG_ICON_KEY, sIconFile)
                        End If
                    End If
                End If
            End If
        End If
    End If
    RegisterFile = (bSuccess = bSuccess2)
End Function
Private Function SetValue(lhKey As Long, _
                          SubKey As String, _
                          sValue As String) As Boolean
Dim lhKeyOpen As Long
Dim lResult   As Long
Dim lTyp      As Long
Dim lByte     As Long
    lByte = Len(sValue)
    lTyp = REG_SZ
    lhKeyOpen = OpenKey(lhKey, SubKey, KEY_SET_VALUE)
    lResult = RegSetValue(lhKey, SubKey, lTyp, sValue, lByte)
    If lResult <> ERROR_SUCCESS Then
        SetValue = False
    Else
        SetValue = True
        RegCloseKey (lhKeyOpen)
    End If
End Function


