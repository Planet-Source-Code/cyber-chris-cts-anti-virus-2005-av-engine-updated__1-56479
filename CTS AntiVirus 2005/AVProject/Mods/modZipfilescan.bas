Attribute VB_Name = "modZipfilescan"
Option Explicit
'Option Explicit
Public Sub ScanZipFile(ByVal strZipFilename As String)
'winzip32.exe -min -e m:\example.zip m:\example
'On Error Resume Next
'Dim c          As Collection
'Dim Candidates As Collection
'Dim file       As Variant
'    strZipFilename = Left$(strZipFilename, Len(strZipFilename) - 1)
'    Shell "c:\programme\winzip\winzip32.exe -min -e " & strZipFilename & " c:\ziptemp"
'    AVE.ScanFolder "c:\ziptemp", True
'    Set Files = New Collection
'    FindFiles "c:\ziptemp\", Candidates, "*.*", vbNormal, True
'    For Each strFilename In Candidates
'        strFilename = Replace(strFilename, "\\", "\")
'        If GetFileOI(strFilename) Then
'            strResult = AVE.ScanFile(strFilename)
'            If strResult <> "NOTHING" Then
'                On Error Resume Next
'                Unload frmAlert
'                With Virus
'                    .FileName = strZipFilename & "@" & Replace(strFilename, "c:\ziptemp\", vbNullString)
'                    .Reason = strResult
'                    temp = Split(.FileName, "\")
'                    .FileNameShort = temp(UBound(temp))
'                End With 'Virus
'                Log "Virus found: " & Virus.Reason & " in " & Virus.FileName & " stored in " & strZipFilename, 1, True
'                SaveSetting AV.AVname, "Settings", "countVirus", GetSetting(AV.AVname, "Settings", "countVirus", 0) + 1
'                frmAlert.Show
'' CheckFile = True
'            End If
'        End If
'    Next strFilename
'    DelTree "c:\ziptemp"
'    On Error GoTo 0
End Sub
