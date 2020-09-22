Attribute VB_Name = "modFileOperations"
Option Explicit
Public Function CalcCRC(ByVal strFilename As String) As String
Dim cCRC32  As New cCRC32
Dim lCRC32  As Long
Dim cStream As New cBinaryFileStream
    On Error GoTo err
    cStream.file = strFilename
    lCRC32 = cCRC32.GetFileCrc32(cStream)
    CalcCRC = Hex$(lCRC32)
Exit Function
err:
'  ErrorFunc err.Number, err.Description, "modAntivir.CalcCRC", strFilename
End Function
Public Function FileText(ByVal strFilename As String) As String
Dim handle As Long
    On Error GoTo err
    handle = FreeFile
    Open strFilename For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
Exit Function
err:
    MsgBox err.Number & " " & err.Description, vbCritical, "Error"
End Function
Public Function FindTermInFile(ByVal strFilename As String, _
                               ByVal strString As String, _
                               ByVal strFiletext As String) As Boolean
    FindTermInFile = False
    If InStr(1, strFiletext, strString, vbTextCompare) <> 0 Then
        FindTermInFile = True
    End If
End Function


