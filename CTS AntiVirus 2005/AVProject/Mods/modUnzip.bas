Attribute VB_Name = "modUnzip"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Type UNZIPnames
    s(0 To 1023)        As String
End Type
Private Type CBChar
    ch(0 To 32800)      As Byte
End Type
Private Type CBCh
    ch(0 To 255)        As Byte
End Type
Public Type DCLIST
    ExtractOnlyNewer    As Long
    SpaceToUnderScore   As Long
    PromptToOverwrite   As Long
    fQuiet              As Long
    ncflag              As Long
    ntflag              As Long
    nvflag              As Long
    nUflag              As Long
    nzflag              As Long
    ndflag              As Long
    noflag              As Long
    naflag              As Long
    nZIflag             As Long
    C_flag              As Long
    fPrivilege          As Long
    lpszZipFN           As String
    lpszExtractDir      As String
End Type
Private Type USERFUNCTION
    lptrPrnt            As Long
    lptrSound           As Long
    lptrReplace         As Long
    lptrPassword        As Long
    lptrMessage         As Long
    lptrService         As Long
    lTotalSizeComp      As Long
    lTotalSize          As Long
    lCompFactor         As Long
    lNumMembers         As Long
    cchComment          As Integer
End Type
Public Type ZIPVERSIONTYPE
    major               As Byte
    minor               As Byte
    patchlevel          As Byte
    not_used            As Byte
End Type
Public Type UZPVER
    structlen           As Long
    flag                As Long
    betalevel           As String * 10
    date                As String * 20
    zlib                As String * 10
    Unzip               As ZIPVERSIONTYPE
    zipinfo             As ZIPVERSIONTYPE
    os2dll              As ZIPVERSIONTYPE
    windll              As ZIPVERSIONTYPE
End Type
Private m_cUnzip      As cUnzip
Private m_bCancel     As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)
Private Declare Function Wiz_SingleEntryUnzip Lib "unzip.dll" (ByVal ifnc As Long, _
                                                               ByRef ifnv As UNZIPnames, _
                                                               ByVal xfnc As Long, _
                                                               ByRef xfnv As UNZIPnames, _
                                                               dcll As DCLIST, _
                                                               Userf As USERFUNCTION) As Long
Public Declare Sub UzpVersion2 Lib "unzip.dll" (uzpv As UZPVER)

Private Sub ParseFileFolder(ByRef sFileName As String, _
                            ByRef sFolder As String)

  Dim iPos     As Long
  Dim iLastPos As Long

    iPos = InStr(sFileName, vbNullChar)
    If (iPos <> 0) Then
        sFileName = Left$(sFileName, iPos - 1)
    End If
    iLastPos = ReplaceSection(sFileName, "/", "\")
    If (iLastPos > 1) Then
        sFolder = Left$(sFileName, iLastPos - 2)
        sFileName = Mid$(sFileName, iLastPos)
    End If

End Sub

Private Function plAddressOf(ByVal lPtr As Long) As Long

    plAddressOf = lPtr

End Function

Private Function ReplaceSection(ByRef sString As String, _
                                ByVal sToReplace As String, _
                                ByVal sReplaceWith As String) As Long

  Dim iPos     As Long
  Dim iLastPos As Long

    iLastPos = 1
    Do
        iPos = InStr(iLastPos, sString, "/")
        If (iPos > 1) Then
            Mid$(sString, iPos, 1) = "\"
            iLastPos = iPos + 1
        End If
    Loop While Not (iPos = 0)
    ReplaceSection = iLastPos

End Function

Private Sub UnzipMessageCallBack(ByVal ucsize As Long, _
                                 ByVal csiz As Long, _
                                 ByVal cfactor As Integer, _
                                 ByVal mo As Integer, _
                                 ByVal dy As Integer, _
                                 ByVal yr As Integer, _
                                 ByVal hh As Integer, _
                                 ByVal mm As Integer, _
                                 ByVal c As Byte, _
                                 ByRef fname As CBCh, _
                                 ByRef meth As CBCh, _
                                 ByVal crc As Long, _
                                 ByVal fCrypt As Byte)

  Dim sFileName As String
  Dim sFolder   As String
  Dim dDate     As Date
  Dim sMethod   As String
  Dim iPos      As Long

    On Error Resume Next
    With m_cUnzip
        sFileName = StrConv(fname.ch, vbUnicode)
        ParseFileFolder sFileName, sFolder
        dDate = DateSerial(yr, mo, hh)
        dDate = dDate + TimeSerial(hh, mm, 0)
        sMethod = StrConv(meth.ch, vbUnicode)
        iPos = InStr(sMethod, vbNullChar)
        If (iPos > 1) Then
            sMethod = Left$(sMethod, iPos - 1)
        End If
        'debug.Print fCrypt
        .DirectoryListAddFile sFileName, sFolder, dDate, csiz, crc, ((fCrypt And 64) = 64), cfactor, sMethod
    End With 'M_CUNZIP
    On Error GoTo 0

End Sub

Private Function UnzipPasswordCallBack(ByRef pwd As CBCh, _
                                       ByVal lngX As Long, _
                                       ByRef s2 As CBCh, _
                                       ByRef cbcName As CBCh) As Long

  Dim bCancel   As Boolean
  Dim sPassword As String
  Dim b()       As Byte
  Dim lSize     As Long

    On Error Resume Next
    UnzipPasswordCallBack = 1
    If m_bCancel Then
        Exit Function
    End If
    m_cUnzip.PasswordRequest sPassword, bCancel
    sPassword = Trim$(sPassword)
    If bCancel Or Len(sPassword) = 0 Then
        m_bCancel = True
        Exit Function
    End If
    lSize = Len(sPassword)
    If lSize > 254 Then
        lSize = 254
    End If
    b = StrConv(sPassword, vbFromUnicode)
    CopyMemory pwd.ch(0), b(0), lSize
    UnzipPasswordCallBack = 0
    On Error GoTo 0

End Function

Private Function UnzipPrintCallback(ByRef fname As CBChar, _
                                    ByVal lngX As Long) As Long

  Dim sFile As String

    On Error Resume Next
    If lngX > 1 Then
        If lngX < 32000 Then
            ReDim b(0 To lngX) As Byte
            CopyMemory b(0), fname, lngX
            sFile = StrConv(b, vbUnicode)
            ReplaceSection sFile, "/", "\"
            m_cUnzip.ProgressReport sFile
        End If
    End If
    UnzipPrintCallback = 0
    On Error GoTo 0

End Function

Private Function UnzipReplaceCallback(ByRef fname As CBChar) As Long

  Dim eResponse As EUZOverWriteResponse
  Dim iPos      As Long
  Dim sFile     As String

    On Error Resume Next
    eResponse = euzDoNotOverwrite
    sFile = StrConv(fname.ch, vbUnicode)
    iPos = InStr(sFile, vbNullChar)
    If (iPos > 1) Then
        sFile = Left$(sFile, iPos - 1)
    End If
    ReplaceSection sFile, "/", "\"
    m_cUnzip.OverwriteRequest sFile, eResponse
    UnzipReplaceCallback = eResponse
    On Error GoTo 0

End Function

Private Function UnZipServiceCallback(ByRef mname As CBChar, _
                                      ByVal lngX As Long) As Long

  Dim iPos    As Long
  Dim sInfo   As String
  Dim bCancel As Boolean

    On Error Resume Next
    If lngX > 1 Then
        If lngX < 32000 Then
            ReDim b(0 To lngX) As Byte
            CopyMemory b(0), mname, lngX
            sInfo = StrConv(b, vbUnicode)
            iPos = InStr(sInfo, vbNullChar)
            If iPos > 0 Then
                sInfo = Left$(sInfo, iPos - 1)
            End If
            ReplaceSection sInfo, "\", "/"
            m_cUnzip.Service sInfo, bCancel
            If bCancel Then
                UnZipServiceCallback = 1
             Else 'BCANCEL = FALSE/0
                UnZipServiceCallback = 0
            End If
        End If
    End If
    On Error GoTo 0

End Function

Public Function VBUnzip(cUnzipObject As cUnzip, _
                        tDCL As DCLIST, _
                        iIncCount As Long, _
                        sInc() As String, _
                        iExCount As Long, _
                        sExc() As String) As Long

  Dim lErr  As Long
  Dim tUser As USERFUNCTION
  Dim tInc  As UNZIPnames
  Dim tExc  As UNZIPnames
  Dim i     As Long

    On Error GoTo ErrorHandler
    Set m_cUnzip = cUnzipObject
    With tUser
        .lptrPrnt = plAddressOf(AddressOf UnzipPrintCallback)
        .lptrSound = 0&
        .lptrReplace = plAddressOf(AddressOf UnzipReplaceCallback)
        .lptrPassword = plAddressOf(AddressOf UnzipPasswordCallBack)
        .lptrMessage = plAddressOf(AddressOf UnzipMessageCallBack)
        .lptrService = plAddressOf(AddressOf UnZipServiceCallback)
    End With 'tUser
    If (iIncCount > 0) Then
        For i = 1 To iIncCount
            tInc.s(i - 1) = sInc(i)
        Next i
        tInc.s(iIncCount) = vbNullChar
     Else 'NOT (IINCCOUNT...
        tInc.s(0) = vbNullChar
    End If
    If (iExCount > 0) Then
        For i = 1 To iExCount
            tExc.s(i - 1) = sExc(i)
        Next i
        tExc.s(iExCount) = vbNullChar
     Else 'NOT (IEXCOUNT...
        tExc.s(0) = vbNullChar
    End If
    m_bCancel = False
    VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)

Exit Function

ErrorHandler:
    lErr = Err.Number
    If lErr = 53 Then
        CheckDLL
        Exit Function
    End If
    VBUnzip = -1
    Set m_cUnzip = Nothing
    'Err.Raise lErr, App.EXEName & ".VBUnzip", Err.Description

End Function


