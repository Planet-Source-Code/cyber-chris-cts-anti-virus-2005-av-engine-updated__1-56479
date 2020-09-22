Attribute VB_Name = "modList"
Option Explicit
' Benötigte API-Deklarationen
' INI lesen+schreiben
Private Const nBUFSIZEINI       As Long = 1024
Private Const nBUFSIZEINIALL    As Long = 4096
Private Declare Function OSGetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                    ByVal lpKeyName As Any, _
                                                                                                    ByVal lpDefault As String, _
                                                                                                    ByVal lpReturnedString As String, _
                                                                                                    ByVal nSize As Long, _
                                                                                                    ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                                        ByVal lpKeyName As Any, _
                                                                                                        ByVal lpString As Any, _
                                                                                                        ByVal lpFileName As String) As Long
' Prüfen, ob Datei existiert
Public Function FileExists(ByVal sFile As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(sFile) <> "")
    On Error GoTo 0
End Function
' INI-Eintrag lesen
Private Function GetINIString(ByVal szSection As String, _
                              ByVal szEntry As Variant, _
                              ByVal szDefault As String, _
                              ByVal szFileName As String) As String
Dim szTmp As String
Dim nRet  As Long
    If IsNull(szEntry) Then
        szTmp = String$(nBUFSIZEINIALL, 0)
        nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
    Else
        szTmp = String$(nBUFSIZEINI, 0)
        nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
    End If
    GetINIString = Left$(szTmp, nRet)
End Function
''
''' INI-Eintrag speichern
''Private Sub WriteINIString(ByVal szSection As String, ByVal szEntry As Variant, ByVal vValue As Variant, ByVal szFileName As String)
''
''
'''Dim nRet As Long
''
''
''
''If IsNull(szEntry) Then
''OSWritePrivateProfileString szSection, 0&, 0&, szFileName
''
''ElseIf (IsNull(vValue)) Then
''OSWritePrivateProfileString szSection, CStr(szEntry), 0&, szFileName
''
''Else
''OSWritePrivateProfileString szSection, CStr(szEntry), CStr(vValue), szFileName
''
''End If
''End Sub
''
''' ListBox-Inhalt aus INI-Datei auslesen
''Public Sub LoadList(oListBox As ListBox, ByVal strFileName As String)
''
''
''
''Dim nCount As Long
''Dim sText  As String
''nCount = 0
''' Inhalt der Liste löschen
''With oListBox
''.Clear
''Do
''' Eintrag aus INI-Liste lesen
''nCount = nCount + 1
''sText = GetINIString("List", CStr(nCount), vbNullString, strFileName)
''' Fügt den Eintrag zur Liste hinzu
''If LenB(sText) Then
''If InStr(sText, "|") > 0 Then
''sText = Replace(sText, "|", String$(10, vbTab))
''End If
''.AddItem sText
''End If
''Loop Until LenB(sText) = 0
''End With
''End Sub
''
''
''' ListBox-Inhalt in INI-Datei speichern
''Public Sub SaveList(oListBox As ListBox, ByVal strFileName As String)
''
''
''
''Dim nCount As Long
''On Error Resume Next
''' Löscht die alte Datei um sicher zu sein,
''' dass jede extra Information die nicht
''' benötigt wird gelöscht wird
''Kill strFileName
''
''With oListBox
''nCount = 0
''' alle ListBox-Einträge speichern
''Do
''nCount = nCount + 1
''WriteINIString "List", CStr(nCount), .List(nCount - 1), strFileName
''Loop Until nCount = .ListCount
''End With
''On Error GoTo 0
''End Sub
''
''


