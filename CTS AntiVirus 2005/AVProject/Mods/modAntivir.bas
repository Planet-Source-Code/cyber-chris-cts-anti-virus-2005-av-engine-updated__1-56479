Attribute VB_Name = "modAntivir"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public AVE  As New AntiVirus.Engine
Public Function CheckFile(ByVal strFileName As String, _
                          Optional ByVal pReport As Boolean = False) As Boolean
Dim strResult As String
Dim temp()    As String
    On Error GoTo err
    CheckFile = False
    If LCase$(Left$(strFileName, Len("signatures.db"))) = "signautres.db" Then
        Exit Function
    End If
    If UCase$(Mid$(strFileName, Len(strFileName) - 4, 4)) = ".ZIP" Or UCase$(Mid$(strFileName, Len(strFileName) - 4, 4)) = ".ARC" Then
        ScanZipFile strFileName
        Exit Function
    Else
        If GetFileOI(strFileName) Then
            strResult = AVE.ScanFile(strFileName)
            If strResult <> "NOTHING" Then
                With Virus
                    .FileName = strFileName
                    .Reason = strResult
                    temp = Split(.FileName, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                Log "Virus found: " & Virus.Reason & " in " & Virus.FileName, 1, True
                SaveSetting AV.AVname, "Settings", "countVirus", GetSetting(AV.AVname, "Settings", "countVirus", 0) + 1
                frmAlert.Show
                CheckFile = True
            End If
        End If
        If AVE.IsFileaScript(strFileName) Then
            If AVE.ScanScript(strFileName) Then
                With Virus
                    .FileName = strFileName
                    .Reason = LoadResString(151)
                    temp = Split(.FileName, "\")
                    .FileNameShort = temp(UBound(temp))
                End With 'Virus
                SaveSetting AV.AVname, "Settings", "countVirus", GetSetting(AV.AVname, "Settings", "countVirus", 0) + 1
                frmAlert.Show
                CheckFile = True
            End If
        End If
    End If
    If pReport Then
        Report = "<p><span class=" & Chr$(34) & "Stil15" & Chr$(34) & ">File Scanned: " & strFileName & "</span></p>" & _
            "<p><span class=" & Chr$(34) & "Stil15" & Chr$(34) & ">File is clear!</span></p>"
        LoadPage (App.path & "\Gui\Report.htm")
        Currentpage = App.path & "\Gui\Report.htm"
    End If
    SaveSetting AV.AVname, "Settings", "countFiles", GetSetting(AV.AVname, "Settings", "countFiles", 0) + 1
    DoEvents
Exit Function
err:
    ErrorFunc err.Number, err.Description, "modAntivir.Checkfile", strFileName
End Function
Public Function FileExist(ByVal strFileName As String) As Boolean
    On Error Resume Next
    FileExist = True
    If FileLen(strFileName) = 0 Then
        FileExist = False
    End If
    On Error GoTo 0
End Function
Public Function FileText(ByVal strFileName As String) As String
Dim handle As Long
    On Error GoTo err
    handle = FreeFile
    Open strFileName For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
Exit Function
err:
    ErrorFunc err.Number, err.Description, "modAntivir.Filetext", strFileName
End Function
Private Function IsWinNT() As Boolean
Dim myOS As OSVERSIONINFO
    On Error GoTo err
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
Exit Function
err:
    ErrorFunc err.Number, err.Description, "modAntivir.IsWinNT"
End Function
Public Sub KeepOnTop(F As Form)
Const SWP_NOMOVE   As Long = 2
Const SWP_NOSIZE   As Long = 1
Const HWND_TOPMOST As Long = -1
    SetWindowPos F.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
Public Function LoadIcon(Size As Long, _
                         ByVal strFileName As String) As IPictureDisp
'Dim Result    As Long
Dim file      As String
Dim Unkown    As IUnknown
Dim Icon      As IconType
Dim CLSID     As CLSIdType
Dim ShellInfo As ShellFileInfoType
    On Error GoTo err
    file = strFileName
    SHGetFileInfo file, 0, ShellInfo, Len(ShellInfo), Size
    With Icon
        .cbSize = Len(Icon)
        .picType = vbPicTypeIcon
        .hIcon = ShellInfo.hIcon
    End With 'Icon
    CLSID.Id(8) = &HC0
    CLSID.Id(15) = &H46
    OleCreatePictureIndirect Icon, CLSID, 1, Unkown
    Set LoadIcon = Unkown
Exit Function
err:
    ErrorFunc err.Number, err.Description, "modAntivir.LoadIcon", Size & ":" & strFileName
End Function
Public Sub Main()
Dim Result
Dim AA
Dim Scriptfile As String
    AVE.Initialise App.path & "\signatures.db", "http://members.aon.at/csss/signatures.db"
    On Error GoTo err
    If App.PrevInstance Then
        MsgBox "Only one instance allowed!", vbOKOnly, "Error"
        End
    End If
    With AV
        .AVname = "CTS Antivirus 2005"
    End With
    SetAttr App.path & "\secure\", vbHidden + vbSystem     ' Set the attributes,..
    CheckExe
    RegisterFile ".secure", LoadResString(135) & AV.AVname, "Anti Virus", App.path & "\" & App.EXEName & ".exe /R %l", App.path & "\secicon.ico"  '"This file is secured by "
'Associate App.path & "\" & App.EXEName & ".exe /SCRIPT %l", "vbs", "CTS AV Protected Visual Basic Script", App.path & "\Script.ico"
    If InStr(1, UCase$(Command), "/SCRIPT") <> 0 Then
        MsgBox "script!!!!!!!!!!"
        SaveSetting AV.AVname, "Settings", "countFiles", GetSetting(AV.AVname, "Settings", "countFiles", 0)
        Scriptfile = Mid$(Command, 9, Len(Command) - 7)
        If AVE.ScanScript(Scriptfile) = False Then
            AA = Space$(255)
            Result = GetShortPathName(Scriptfile, AA, Len(AA))
            Shell "c:\windows\System32\WScript.exe """"" & Mid$(AA, 1, Result) & """"""
        Else
            frmScriptAlert.Show
            frmScriptAlert.lblText(1).Caption = Scriptfile
            SaveSetting AV.AVname, "Settings", "countVirus", GetSetting(AV.AVname, "Settings", "countVirus", 0)
        End If
        Exit Sub
    End If
    Select Case UCase$(Left$(Command, 2))
    Case "/S"
        CheckFile (Mid$(Command, 3, Len(Command) - 3))
    Case vbNullString
        frmMain.Show
    Case "/G"
        frmMain.Show
    Case "/U"
        frmUpdate.Show
    Case "/T"
        frmMain.Visible = False
    Case "/C"
        frmMain.Show
        AV.Runmode = Normal
        frmMain.ShowFileSearch
    Case "/F"
        AV.Runmode = ScanFile
        Checkfolder (Mid$(Command, 3, Len(Command) - 3))
    Case "/R"
        MsgBox "This file is secured! If you want to desecure it, goto: Main/Extras/Secured Files"
        End
    Case Else
        MsgBox "Invalid Parameter!"
    End Select
Exit Sub
err:
    ErrorFunc err.Number, err.Description, "modAntivir.Main"
End Sub
Public Function ShowOpenDlg(ByVal Owner As Form, _
                            Optional ByVal InitialDir As String, _
                            Optional ByVal strFilter As String, _
                            Optional ByVal DefaultExtension As String, _
                            Optional ByVal DlgTitle As String) As String
Dim sBuf As String
    On Error GoTo err
    InitialDir = IIf(IsMissing(InitialDir), vbNullString, InitialDir)
    strFilter = IIf(IsMissing(strFilter), LoadResString(129) & "|*.*", Replace(strFilter, "|", vbNullChar)) & vbNullChar
    DefaultExtension = IIf(IsMissing(DefaultExtension), vbNullString, DefaultExtension)
    DlgTitle = IIf(IsMissing(DlgTitle), LoadResString(138), DlgTitle)
    sBuf = Space$(256)
    If IsWinNT Then
        GetFileNameFromBrowseW Owner.hwnd, StrPtr(sBuf), Len(sBuf), StrPtr(InitialDir), StrPtr(DefaultExtension), StrPtr(strFilter), StrPtr(DlgTitle)
    Else
        GetFileNameFromBrowseA Owner.hwnd, sBuf, Len(sBuf), InitialDir, DefaultExtension, strFilter, DlgTitle
    End If
    ShowOpenDlg = Trim$(sBuf)
Exit Function
err:
    ErrorFunc err.Number, err.Description, "modAntivir.ShowOpenDlg", Owner.Name & ":" & InitialDir & ":" & strFilter & ":" & DefaultExtension & ":" & DlgTitle
End Function


