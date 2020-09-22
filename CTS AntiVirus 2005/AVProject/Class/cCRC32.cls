VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCRC32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private crc32Table()         As Long
Private Const BUFFER_SIZE    As Long = 8192
Private Sub Class_Initialize()
Dim i            As Long
Dim j            As Long
Dim dwCrc        As Long
Dim dwPolynomial As Long
    dwPolynomial = &HEDB88320
    ReDim crc32Table(256)
    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                dwCrc = dwCrc Xor dwPolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
            End If
        Next j
        crc32Table(i) = dwCrc
    Next i
End Sub
Public Function GetFileCrc32(ByRef stream As cBinaryFileStream) As Long
Dim buffer(0 To BUFFER_SIZE - 1) As Byte
Dim readSize                     As Long
Dim count                        As Long
Dim i                            As Long
Dim iLookup                      As Long
Dim crc32Result                  As Long
    crc32Result = &HFFFFFFFF
    readSize = BUFFER_SIZE
    count = stream.Read(buffer, readSize)
    Do While (count > 0)
        For i = 0 To count - 1
            iLookup = (crc32Result And &HFF) Xor buffer(i)
            crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And 16777215
' nasty shr 8 with vb :/
            crc32Result = crc32Result Xor crc32Table(iLookup)
        Next i
        count = stream.Read(buffer, readSize)
    Loop
    GetFileCrc32 = Not (crc32Result)
End Function
