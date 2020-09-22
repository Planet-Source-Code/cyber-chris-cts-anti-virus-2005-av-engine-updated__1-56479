VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBinaryFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_sFile     As String
Private m_iFile     As Long
Private m_iLen      As Long
Private m_iOffset   As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)
Private Sub Class_Terminate()
    Dispose
End Sub
Public Sub Dispose()
    If m_iFile Then
        Close #m_iFile
        m_iFile = 0
    End If
End Sub
Public Property Get file() As String
    file = m_sFile
End Property
Public Property Let file(ByVal sFile As String)
Dim lErr As Long
    On Error Resume Next
    Dispose
    m_sFile = sFile
    If (FileExists(m_sFile, lErr)) Then
        m_iFile = FreeFile
        Open m_sFile For Binary Access Read Lock Write As #m_iFile
        m_iLen = LOF(m_iFile)
    Else
'err.Raise lErr, App.EXEName & ".File"
    End If
    On Error GoTo 0
End Property
Private Function FileExists(ByVal sFile As String, _
                            ByRef lErr As Long) As Boolean
Dim sDir As String
    lErr = 0
    On Error Resume Next
    sDir = Dir(sFile)
    lErr = err.Number
    On Error GoTo 0
    If lErr = 0 Then
        If Len(sDir) > 0 Then
            FileExists = True
        Else 'NOT (LEN(SDIR)...
            lErr = 53
        End If
    End If
End Function
Public Property Get Length() As Long
    Length = m_iLen
End Property
Public Function Read(buffer() As Byte, _
                     ByVal readSize As Long) As Long
Dim lReadSize As Long
    lReadSize = readSize
    If (m_iOffset + lReadSize >= m_iLen) Then
        readSize = m_iLen - m_iOffset
        If readSize > 0 Then
            ReDim newBuffer(0 To readSize - 1) As Byte
            Get #m_iFile, , newBuffer
            CopyMemory buffer(0), newBuffer(0), readSize
        Else 'NOT (READSIZE...
            Dispose
        End If
        m_iOffset = m_iOffset + readSize
    Else
' Can read
        Get #m_iFile, , buffer
        m_iOffset = m_iOffset + readSize
    End If
    Read = readSize
End Function


