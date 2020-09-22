VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Engine"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Sign()                      As String    'The Signatures will be loaded into this array
''Private SignStr()                   As String
Private SignVirusStringType()       As String * 1
Private SignVirusName()             As String
Private ScriptHeuristic             As String
Private Type SHFILEOPSTRUCT
    Hwnd                                As Long
    wFunc                               As Long
    pFrom                               As String
    pTo                                 As String
    fFlags                              As Long
    fAnyOperationsAborted               As Boolean
    hNameMappings                       As Long
    lpszProgressTitle                   As String
End Type
Private Const FO_DELETE             As Long = &H3
'//------------------------ Informationblock
Private Type FILETIME
    dwLowDateTime                       As Long
    dwHighDateTime                      As Long
End Type
Private Type SgnData
    SignatureDate                       As String
    SignatureCount                      As Long
    SignatureFilename                   As String
    SignatureUpdatepath                 As String
End Type
Private Type AntiVir
    Signature                           As SgnData
End Type
Private AV                          As AntiVir
Private Running                     As Boolean
Private Type WIN32_FIND_DATA
    dwFileAttributes                    As Long
    ftCreationTime                      As FILETIME
    ftLastAccessTime                    As FILETIME
    ftLastWriteTime                     As FILETIME
    nFileSizeHigh                       As Long
    nFileSizeLow                        As Long
    dwReserved0                         As Long
    dwReserved1                         As Long
    cFileName                           As String * 260
    cAlternate                          As String * 14
End Type
Public Event VirusFound(Filename As String, InfectedWith As String)
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function FindFirstFileA Lib "KERNEL32" (ByVal lpFileName As String, _
                                                        lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFileA Lib "KERNEL32" (ByVal hFindFile As Long, _
                                                       lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributesA Lib "KERNEL32" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "KERNEL32" (ByVal hFindFile As Long) As Long
Public Sub About()
    frmAbout.Show
End Sub
Private Function AnalyseFolder(ByVal StrPath As String, _
                               ByRef Files As Collection, _
                               Optional ByVal strPattern As String = "*.*", _
                               Optional ByVal Attributes As VbFileAttribute = vbNormal, _
                               Optional ByVal Recursive As Boolean = True) As Long
Const vbErr_PathNotFound As Long = 76
Const INVALID_VALUE      As Long = -1
Dim FileAttr             As Long
Dim Filename             As String
Dim hFind                As Long
Dim WFD                  As WIN32_FIND_DATA
    If Mid$(StrPath, Len(StrPath) - 1, 1) <> "\" Then
        StrPath = StrPath & "\"
    End If
    If Files Is Nothing Then
        Set Files = New Collection
    End If
    strPattern = LCase$(strPattern)
    hFind = FindFirstFileA(StrPath & "*", WFD)
    If hFind = INVALID_VALUE Then
        err.Raise vbErr_PathNotFound
    End If
    Do
        Filename = LeftB$(WFD.cFileName, InStrB(WFD.cFileName, vbNullChar))
        FileAttr = GetFileAttributesA(StrPath & Filename)
        If FileAttr And vbDirectory Then
            If Recursive Then
                If FileAttr <> INVALID_VALUE Then
                    If Filename <> "." Then
                        If Filename <> ".." Then
                            AnalyseFolder = AnalyseFolder + AnalyseFolder(StrPath & Filename, Files, strPattern, Attributes)
                        End If
                    End If
                End If
            End If
        Else
            If (FileAttr And Attributes) = Attributes Then
                If LCase$(Filename) Like strPattern Then
                    AnalyseFolder = AnalyseFolder + 1
                    Files.Add StrPath & Filename
                End If
            End If
        End If
    Loop While FindNextFileA(hFind, WFD)
    FindClose hFind
End Function
Private Sub BuildSigns()
Dim Data()   As String
Dim sIn      As String
Dim swords() As String
Dim X        As Long
''Dim Y        As Long
    sIn = FileText(AV.Signature.SignatureFilename)
    swords = Split(sIn, vbNewLine)
    ReDim Preserve swords(UBound(swords) - 1)
    sIn = ""
    For X = LBound(swords) To UBound(swords)
        ReDim Preserve Sign(0 To X) As String
        ReDim Preserve SignVirusStringType(0 To X) As String * 1
        ReDim Preserve SignVirusName(0 To X) As String
        Data = Split(swords(X) & ":" & ":", ":")
        Sign(X) = UCase$(Data(0))
        SignVirusStringType(X) = Data(1)
        SignVirusName(X) = Data(2)
'Y = X + 1
    Next X
    ReDim Preserve Sign(0 To X + 1) As String
    Sign(X + 1) = "#END#"
    AV.Signature.SignatureDate = Sign(0)
    AV.Signature.SignatureCount = UBound(swords) - 1
Exit Sub
    err
    MsgBox "An error has occured while loading the signature File!" & vbNewLine & _
       "This could be caused by an empty or damaged file!" & vbNewLine & _
       "The error message was: " & err.Description, vbCritical + vbOKOnly, LoadResString(140)
End Sub
Public Sub Initialise(ByVal SignatureFile As String, _
                      ByVal UpdatePath As String)
    AV.Signature.SignatureFilename = SignatureFile
    ScriptHeuristic = "DEL,KILL,FORMAT,REN,COPY,XCOPY,OUTLOOK,EMAIL"
    BuildSigns
' MsgBox AV.Signature.SignatureCount
End Sub
Public Function IsFileaScript(strFilename As String) As Boolean
    IsFileaScript = False
    strFilename = UCase$(strFilename)
    If Mid$(strFilename, Len(strFilename) - 3, 3) = ".JS" Then
        IsFileaScript = True
    ElseIf Mid$(strFilename, Len(strFilename) - 3, 3) = "VBS" Then
        IsFileaScript = True
    End If
End Function
'//------------------------
Public Sub RemoveFileWithDialog(ByVal strFilename As String, _
                                ByVal lngHwnd As Long)
Dim Files As String
Dim SFO   As SHFILEOPSTRUCT
    On Error GoTo err
    DoEvents
    Files = strFilename & vbNullChar
    Files = Files & vbNullChar
    With SFO
        .Hwnd = lngHwnd
        .wFunc = FO_DELETE
        .pFrom = Files
        .pTo = "" & vbNullChar
    End With
    SHFileOperation SFO
Exit Sub
err:
    MsgBox err.Description
End Sub
Public Function ScanFile(ByVal strFilename As String) As String
Dim Current     As Long
Dim crc         As String
Dim strFiletext As String
'Dim Zeilen()    As String
    On Error GoTo err
    crc = CalcCRC(strFilename)
    strFiletext = Replace(CStr(FileText(strFilename)), "ß", "-")
    Debug.Print strFilename
    For Current = 1 To 4096
        If Sign(Current) = "#END#" Or LenB(Sign(Current)) = 0 Then
            GoTo Finish
        End If
        If SignVirusStringType(Current) = "E" Then
            If crc = Sign(Current) Then
                DoEvents
                RaiseEvent VirusFound(strFilename, SignVirusName(Current))
                ScanFile = SignVirusName(Current)
                Exit Function
            Else
                ScanFile = "NOTHING"
            End If
        ElseIf SignVirusStringType(Current) = "S" Then
            If FindTermInFile(strFilename, Sign(Current), strFiletext) Then
                DoEvents
                RaiseEvent VirusFound(strFilename, SignVirusName(Current))
                ScanFile = SignVirusName(Current)
                Exit Function
            Else
                ScanFile = "NOTHING"
            End If
        End If
        DoEvents
    Next Current
    ScanFile = "NOTHING"
Finish:
Exit Function
err:
'ToDo: Ein gescheites Errorhandling
End Function
Public Sub ScanFolder(StrPath As String, _
                      Optional ByVal Recursive As Boolean = False)
Dim Candidates As Collection
Dim file       As Variant
Dim Files      As Collection
    Running = True
    Set Files = New Collection
    AnalyseFolder StrPath, Candidates, "*.*", vbNormal, Recursive
    For Each file In Candidates
        ScanFile (file)
        If Running = False Then
            Exit Sub
        End If
    Next file
End Sub
Public Function ScanScript(ByVal strFilename As String) As Boolean
Dim Textin         As String
Dim temp           As Long
Dim Searchstring() As String
Dim Istr           As cInString
    Set Istr = New cInString
    Searchstring = Split(ScriptHeuristic, ",")
    Textin = UCase$(FileText(strFilename))
    ScanScript = False
    For temp = 0 To UBound(Searchstring)
        If InStr(1, Textin, Searchstring(temp), vbTextCompare) <> 0 Then
'If InStr(1, Textin, Searchstring(temp), vbTextCompare) <> 0 Then
            ScanScript = True
            RaiseEvent VirusFound(strFilename, (Searchstring(temp)))
            Exit Function
        End If
    Next temp
    Set Istr = Nothing
End Function
Public Property Get ScriptHeuristicList() As String
    ScriptHeuristicList = ScriptHeuristic
End Property
Public Property Let ScriptHeuristicList(ByVal HeuristicList As String)
    ScriptHeuristic = UCase$(HeuristicList)
End Property
Public Property Get SignatureCount() As Long
    SignatureCount = AV.Signature.SignatureCount
End Property
Public Property Get SignatureDatabaseEntry(EntryNmr As Long) As String
    SignatureDatabaseEntry = SignVirusName(EntryNmr)
End Property
Public Property Get SignatureDate() As Variant
    SignatureDate = AV.Signature.SignatureDate
End Property
Public Property Get SignatureFilename() As String
    SignatureFilename = AV.Signature.SignatureFilename
End Property
Public Property Let SignatureFilename(ByVal strFilename As String)
    Initialise strFilename, AV.Signature.SignatureUpdatepath
End Property
Public Property Get SignatureUpdatepath() As String
    SignatureUpdatepath = AV.Signature.SignatureUpdatepath
End Property
Public Property Let SignatureUpdatepath(ByVal strFilename As String)
    AV.Signature.SignatureUpdatepath = strFilename
End Property
''
''Public Sub ScriptHeuristicListAdd(NewEntry As String)
''
''
''ScriptHeuristicList = ScriptHeuristicList & "," & (NewEntry)
''End Sub
''


