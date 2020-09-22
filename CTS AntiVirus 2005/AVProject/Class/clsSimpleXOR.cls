VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSimpleXOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Key()                  As Byte
Private m_KeyLen                 As Long
Private m_KeyValue               As String
Public Event Progress(Percent As Long)
Public Sub DecryptByte(ByteArray() As Byte, _
                       Optional key As String)
    EncryptByte ByteArray(), key
End Sub
Public Sub DecryptFile(SourceFile As String, _
                       DestFile As String, _
                       Optional key As String)
Dim Filenr      As Long
Dim ByteArray() As Byte
    If Not FileExist(SourceFile) Then
        Exit Sub
    End If
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1) As Byte
    Get #Filenr, , ByteArray()
    Close #Filenr
    DecryptByte ByteArray(), key
    If FileExist(DestFile) Then
        Kill DestFile
    End If
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr
End Sub
Public Sub EncryptByte(ByteArray() As Byte, _
                       Optional ByVal key As String)
Dim Offset      As Long
Dim ByteLen     As Long
Dim ResultLen   As Long
Dim CurrPercent As Long
Dim NextPercent As Long
    If Len(key) > 0 Then
        Me.key = key
    End If
    ByteLen = UBound(ByteArray) + 1
    ResultLen = ByteLen
    For Offset = 0 To (ByteLen - 1)
        ByteArray(Offset) = ByteArray(Offset) Xor m_Key(Offset Mod m_KeyLen)
        If Offset >= NextPercent Then
            CurrPercent = Int((Offset / ResultLen) * 100)
            NextPercent = (ResultLen * ((CurrPercent + 1) / 100)) + 1
            RaiseEvent Progress(CurrPercent)
        End If
    Next '  OFFSET '  OFFSET '  OFFSET OFFSET OFFSET OFFSET
    If CurrPercent <> 100 Then
        RaiseEvent Progress(100)
    End If
End Sub
Public Sub EncryptFile(SourceFile As String, _
                       DestFile As String, _
                       Optional key As String)
Dim Filenr      As Long
Dim ByteArray() As Byte
    If Not FileExist(SourceFile) Then
        MsgBox "Source file doesn't exist", "SimpleXOR Error Procedure"
        Exit Sub
    End If
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1) As Byte
    Get #Filenr, , ByteArray()
    Close #Filenr
    EncryptByte ByteArray(), key
    If FileExist(DestFile) Then
        Kill DestFile
    End If
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr
End Sub
Public Property Let key(ByVal New_Value As String)
    If m_KeyValue = New_Value Then
        Exit Property
    End If
    m_KeyValue = New_Value
    m_KeyLen = Len(New_Value)
    m_Key() = StrConv(m_KeyValue, vbFromUnicode)
End Property


