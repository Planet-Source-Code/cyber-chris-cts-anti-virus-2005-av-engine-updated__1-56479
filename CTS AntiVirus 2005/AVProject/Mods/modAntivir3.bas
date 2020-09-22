Attribute VB_Name = "modAntivir3"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Sub DelTree(sFolder As String)
Dim sCurrFile As String
    sCurrFile = Dir(sFolder & "\*.*", vbDirectory)
    Do While Len(sCurrFile) > 0
        If sCurrFile <> "." And sCurrFile <> ".." Then
            If (GetAttr(sFolder & "\" & sCurrFile) And vbDirectory) = vbDirectory Then
                DelTree sFolder & "\" & sCurrFile
                sCurrFile = Dir(sFolder & "\*.*", vbDirectory)
            Else
                Kill sFolder & "\" & sCurrFile
                sCurrFile = Dir
            End If
        Else
            sCurrFile = Dir
        End If
    Loop
    RmDir sFolder
End Sub
Public Function GetFileOI(ByVal strFileName As String) As Boolean
Dim Counter As Long
    If GetSetting(AV.AVname, "Settings", "Scan for Exec", True) Then
        For Counter = 1 To Len(FileTypesofInterrest) Step 3
            If InStr(1, strFileName, Mid$(FileTypesofInterrest, Counter, 3), vbTextCompare) <> 0 Then
                GetFileOI = True
                Exit Function
            End If
        Next Counter
    End If
    GetFileOI = False
End Function
Public Sub Log(ByVal strLog As String, _
               ByVal typ As Long, _
               Optional ByVal ForceLog As Boolean)
    If GetSetting(AV.AVname, "Settings", "LogFile", "OFF") = "OFF" Then
        If ForceLog <> True Then
            Exit Sub
        End If
    End If
    frmLog.AddLog typ, 0, strLog
End Sub
''
''Public Sub Associate(Program As String, Extension As String, ByVal Description As String, Optional ByVal strIcon As String)
''
''
''
''
''
''
'''RGCreateKey HKEY_CLASSES_ROOT, "." & Extension
'''RGSetKeyValue HKEY_CLASSES_ROOT, "." & Extension, "", UCase(Extension) & "File"
''' RGCreateKey HKEY_CLASSES_ROOT, UCase(Extension) & "File"
''' RGCreateKey HKEY_CLASSES_ROOT, UCase(Extension) & "File\shell"
'''    RGCreateKey HKEY_CLASSES_ROOT, UCase(Extension) & "File\shell\open"
'''   RGCreateKey HKEY_CLASSES_ROOT, UCase(Extension) & "File\shell\open\command"
''RGSetKeyValue HKEY_CLASSES_ROOT, UCase$(Extension) & "File\shell\open\command", "", Program
'''  RGCreateKey HKEY_CLASSES_ROOT, UCase(Extension) & "File\DefaultIcon"
''' RGSetKeyValue HKEY_CLASSES_ROOT, UCase(Extension) & "File", "", Description 'Set file description
'''RGSetKeyValue HKEY_CLASSES_ROOT, UCase(Extension) & "File\DefaultIcon", "", Icon 'Set file icon
''End Function
''


