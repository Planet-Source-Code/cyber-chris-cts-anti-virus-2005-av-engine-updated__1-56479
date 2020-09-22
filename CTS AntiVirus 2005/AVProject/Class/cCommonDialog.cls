VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const WM_USER                       As Long = &H400
Private Const WM_DESTROY                    As Long = &H2
Private Const WM_NOTIFY                     As Long = &H4E
Private Const WM_INITDIALOG                 As Long = &H110
Private Const CDN_FIRST                     As Long = (-601)
Private Const CDN_SELCHANGE                 As Long = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE              As Long = (CDN_FIRST - &H2)
Private Const CDN_FILEOK                    As Long = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE                As Long = (CDN_FIRST - &H6)
Private Const CDM_FIRST                     As Double = (WM_USER + 100)
Private Const CDM_GETFILEPATH               As Long = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH             As Long = (CDM_FIRST + &H2)
Private Const SM_CXDLGFRAME                 As Long = 7
Private Const SM_CYDLGFRAME                 As Long = 8
Private Const SM_CYCAPTION                  As Long = 4
Private Const MAX_PATH                      As Long = 260
Public Enum CdlFlags
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private OFN_ALLOWMULTISELECT, OFN_CREATEPROMPT, OFN_ENABLEHOOK, OFN_ENABLETEMPLATE, OFN_ENABLETEMPLATEHANDLE
Private OFN_EXPLORER, OFN_EXTENSIONDIFFERENT, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_LONGNAMES
Private OFN_NOCHANGEDIR, OFN_NODEREFERENCELINKS, OFN_NOLONGNAMES, OFN_NONETWORKBUTTON, OFN_NOREADONLYRETURN
Private OFN_NOTESTFILECREATE, OFN_NOVALIDATE, OFN_OVERWRITEPROMPT, OFN_PATHMUSTEXIST, OFN_READONLY
Private OFN_SHAREAWARE, OFN_SHAREFALLTHROUGH, OFN_SHARENOWARN, OFN_SHAREWARN, OFN_SHOWHELP
#End If
#If False Then
Private OFN_ALLOWMULTISELECT, OFN_CREATEPROMPT, OFN_ENABLEHOOK, OFN_ENABLETEMPLATE, OFN_ENABLETEMPLATEHANDLE
Private OFN_EXPLORER, OFN_EXTENSIONDIFFERENT, OFN_FILEMUSTEXIST, OFN_HIDEREADONLY, OFN_LONGNAMES
Private OFN_NOCHANGEDIR, OFN_NODEREFERENCELINKS, OFN_NOLONGNAMES, OFN_NONETWORKBUTTON, OFN_NOREADONLYRETURN
Private OFN_NOTESTFILECREATE, OFN_NOVALIDATE, OFN_OVERWRITEPROMPT, OFN_PATHMUSTEXIST, OFN_READONLY
Private OFN_SHAREAWARE, OFN_SHAREFALLTHROUGH, OFN_SHARENOWARN, OFN_SHAREWARN, OFN_SHOWHELP
#End If
Public Enum eCCDLG_AlignConstants
    Right = 0
    Bottom = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Right, Bottom
#End If
#If False Then
Private Right, Bottom
#End If
Private Type RECT
    Left                                        As Long
    Top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type
Private Type NMHDR
    hwndFrom                                    As Long
    IDfrom                                      As Long
    Code                                        As Long
End Type
Private Type OPENFILENAME
    lStructSize                                 As Long
    hwndOwner                                   As Long
    hInstance                                   As Long
    lpstrFilter                                 As String
    lpstrCustomFilter                           As String
    nMaxCustFilter                              As Long
    nFilterIndex                                As Long
    lpstrFile                                   As String
    nMaxFile                                    As Long
    lpstrFileTitle                              As String
    nMaxFileTitle                               As Long
    lpstrInitialDir                             As String
    lpstrTitle                                  As String
    flags                                       As Long
    nFileOffset                                 As Long
    nFileExtension                              As Long
    lpstrDefExt                                 As String
    lCustData                                   As Long
    lpfnHook                                    As Long
    lpTemplateName                              As String
End Type
Private m_fOpenFlags                        As CdlFlags
Private m_fSaveFlags                        As CdlFlags
Private m_sOpenFilter                       As String
Private m_sSaveFilter                       As String
Private m_lOpenMaxFileLen                   As Long
Private m_lSaveMaxFileLen                   As Long
Private m_sOpenDialogTitle                  As String
Private m_sSaveDialogTitle                  As String
Private m_sOpenInitialDir                   As String
Private m_sSaveInitialDir                   As String
Private m_cControl                          As Control
Private m_eAlign                            As eCCDLG_AlignConstants
Public Event FileChanged(ByVal FileName As String)
Public Event FolderChanged(ByVal path As String)
Public Event DialogInitialized()
Public Event DialogClosed()
Public Event PressedOKButton()
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
                                                                         Source As Any, _
                                                                         ByVal Length As Long)
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "USER32" (ByVal hwnd As Long, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "USER32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function SendMessageString Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                                  ByVal wMsg As Long, _
                                                                                  ByVal wParam As Long, _
                                                                                  ByVal lParam As String) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOFN As OPENFILENAME) As Long
Private Function addr(ByVal a As Long) As Long
    addr = a
End Function
Private Sub Class_Initialize()
    m_fOpenFlags = OFN_EXPLORER + OFN_ENABLEHOOK + OFN_PATHMUSTEXIST + OFN_HIDEREADONLY + OFN_FILEMUSTEXIST
    m_fSaveFlags = OFN_EXPLORER + OFN_ENABLEHOOK + OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST
    m_sOpenFilter = "Alle Dateien (*.*)|*.*"
    m_sSaveFilter = "Alle Dateien (*.*)|*.*"
    m_lOpenMaxFileLen = MAX_PATH
    m_lSaveMaxFileLen = MAX_PATH
End Sub
Private Property Get ControlAlign() As eCCDLG_AlignConstants
    ControlAlign = m_eAlign
End Property
Private Property Let ControlAlign(ByVal eNew As eCCDLG_AlignConstants)
    m_eAlign = eNew
End Property
Public Property Get ControlToSetNewParent() As Control
    Set ControlToSetNewParent = m_cControl
End Property
Public Property Let ControlToSetNewParent(cNew As Control)
    Set m_cControl = cNew
End Property
Private Function GetDlgPath(ByVal lConst As Long, _
                            hWndDlg As Long) As String
Dim lPos As Long
Dim hwnd As Long
Dim sBuf As String
    hwnd = GetParent(hWndDlg)
    sBuf = String$(MAX_PATH, 0)
    SendMessageString hwnd, lConst, MAX_PATH, sBuf
    lPos = InStr(1, sBuf, vbNullChar)
    If lPos > 0 Then
        GetDlgPath = Left$(sBuf, lPos - 1)
    Else
        GetDlgPath = sBuf
    End If
End Function
Private Property Get OpenDialogTitle() As String
    OpenDialogTitle = m_sOpenDialogTitle
End Property
Private Property Let OpenDialogTitle(ByVal sNew As String)
    m_sOpenDialogTitle = sNew
End Property
Private Property Get OpenFilter() As String
    OpenFilter = m_sOpenFilter
End Property
Private Property Let OpenFilter(ByVal sNew As String)
    m_sOpenFilter = sNew
End Property
Private Property Get OpenFlags() As CdlFlags
    OpenFlags = m_fOpenFlags
End Property
Private Property Let OpenFlags(fNew As CdlFlags)
    m_fOpenFlags = fNew
End Property
Private Property Get OpenInitialDir() As String
    OpenInitialDir = m_sOpenInitialDir
End Property
Private Property Let OpenInitialDir(ByVal sNew As String)
    m_sOpenInitialDir = sNew
End Property
Private Property Get OpenMaxFileLen() As String
    OpenMaxFileLen = m_lOpenMaxFileLen
End Property
Private Property Let OpenMaxFileLen(ByVal lNew As String)
    m_lOpenMaxFileLen = lNew
End Property
Public Sub pIncomingMessage(ByVal hWndDlg As Long, _
                            ByVal Msg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long)
''Dim scHeight     As Single
''Dim scWidth      As Single
Dim tR           As RECT
Dim lBorderSize  As Long
Dim lCaptionSize As Long
Dim s            As String
Static OldhWnd   As Long
Dim tNMH         As NMHDR
    Select Case Msg
    Case WM_INITDIALOG
        If Not m_cControl Is Nothing Then
            m_cControl.Visible = True
            OldhWnd = GetParent(m_cControl.hwnd)
            lBorderSize = GetSystemMetrics(SM_CXDLGFRAME)
            lCaptionSize = GetSystemMetrics(SM_CYCAPTION)
            SetParent m_cControl.hwnd, GetParent(hWndDlg)
            GetWindowRect GetParent(hWndDlg), tR
'scHeight = Screen.Height / Screen.TwipsPerPixelY
'scWidth = Screen.Width / Screen.TwipsPerPixelX
            If m_eAlign = Right Then
                MoveWindow GetParent(hWndDlg), tR.Left, tR.Top, tR.Right - tR.Left + m_cControl.Width \ Screen.TwipsPerPixelX + lBorderSize * 2, tR.Bottom - tR.Top, 1
                MoveWindow m_cControl.hwnd, tR.Right - tR.Left - lBorderSize, 4&, m_cControl.Width \ Screen.TwipsPerPixelX, tR.Bottom - tR.Top - lCaptionSize - GetSystemMetrics(SM_CYDLGFRAME) * 2 - 8, 1
            Else
                MoveWindow GetParent(hWndDlg), tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top + m_cControl.Height \ Screen.TwipsPerPixelY - lBorderSize * 2, 1
                MoveWindow m_cControl.hwnd, 4, tR.Bottom - tR.Top - lCaptionSize - GetSystemMetrics(SM_CYDLGFRAME) * 2 - 10, tR.Right - tR.Left - lBorderSize * 2 - 8, m_cControl.Height \ Screen.TwipsPerPixelY, 1
            End If
        End If
        RaiseEvent DialogInitialized
    Case WM_NOTIFY
        CopyMemory tNMH, ByVal lParam, Len(tNMH)
        Select Case tNMH.Code
        Case CDN_SELCHANGE
            s = GetDlgPath(CDM_GETFILEPATH, hWndDlg)
            If PathIsDirectory(s) = False And CBool(PathFileExists(s)) Then
                RaiseEvent FileChanged(s)
            End If
        Case CDN_FOLDERCHANGE
            RaiseEvent FolderChanged(GetDlgPath(CDM_GETFOLDERPATH, hWndDlg))
        Case CDN_FILEOK
            RaiseEvent PressedOKButton
        Case CDN_TYPECHANGE
        End Select
    Case WM_DESTROY
        If Not m_cControl Is Nothing Then
            SetParent m_cControl.hwnd, OldhWnd
            m_cControl.Visible = False
        End If
        RaiseEvent DialogClosed
    End Select
End Sub
Private Property Get SaveDialogTitle() As String
    SaveDialogTitle = m_sSaveDialogTitle
End Property
Private Property Let SaveDialogTitle(ByVal sNew As String)
    m_sSaveDialogTitle = sNew
End Property
Private Property Get SaveFilter() As String
    SaveFilter = m_sSaveFilter
End Property
Private Property Let SaveFilter(ByVal sNew As String)
    m_sSaveFilter = sNew
End Property
Private Property Get SaveFlags() As CdlFlags
    SaveFlags = m_fSaveFlags
End Property
Private Property Let SaveFlags(fNew As CdlFlags)
    m_fSaveFlags = fNew
End Property
Private Property Get SaveInitialDir() As String
    SaveInitialDir = m_sSaveInitialDir
End Property
Private Property Let SaveInitialDir(ByVal sNew As String)
    m_sSaveInitialDir = sNew
End Property
Private Property Get SaveMaxFileLen() As String
    SaveMaxFileLen = m_lSaveMaxFileLen
End Property
Private Property Let SaveMaxFileLen(ByVal lNew As String)
    m_lSaveMaxFileLen = lNew
End Property
Public Function ShowOpen(ByVal hwndOwner As Long) As String
Dim r    As Long
Dim OFN  As OPENFILENAME
Dim sBuf As String
    sBuf = String$(m_lOpenMaxFileLen, 0)
    If (m_fOpenFlags And OFN_ENABLEHOOK) <> OFN_ENABLEHOOK Then
        m_fOpenFlags = m_fOpenFlags + OFN_ENABLEHOOK
    End If
    With OFN
        .lStructSize = Len(OFN)
        .flags = m_fOpenFlags
        .hInstance = App.hInstance
        .hwndOwner = hwndOwner
        .lpfnHook = addr(AddressOf ComDlgCallback)
        .lpstrFile = sBuf
        .lpstrFilter = Replace(m_sOpenFilter, "|", vbNullChar)
        .lpstrInitialDir = m_sOpenInitialDir
        .lpstrTitle = m_sOpenDialogTitle
        .nFilterIndex = 1
        .nMaxFile = Len(sBuf)
        r = GetOpenFileName(OFN)
        If r <> 0 Then
            ShowOpen = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
        End If
    End With
End Function


