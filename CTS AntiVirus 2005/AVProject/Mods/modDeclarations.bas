Attribute VB_Name = "modDeclarations"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Running                              As Boolean
Public logType(1 To 10)                     As String
Public Type IconType
    cbSize                                      As Long
    picType                                     As PictureTypeConstants
    hIcon                                       As Long
End Type
Public Type CLSIdType
    Id(16)                                      As Byte
End Type
Public Type ShellFileInfoType
    hIcon                                       As Long
    iIcon                                       As Long
    dwAttributes                                As Long
    szDisplayName                               As String * 260
    szTypeName                                  As String * 80
End Type
Public Const Large                          As Long = &H100
Public Const VER_PLATFORM_WIN32_NT          As Long = 2
Public Type OSVERSIONINFO
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
End Type
Private Type TypeSignature
    SignatureFilename                           As String
    SignatureDate                               As String
    SignatureOnlineFilename                     As String
    SignatureCount                              As Long
End Type
Public Enum RM
    Normal = 0
    TrayOnly = 1
    ScanFile = 3
    SecureFile = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Normal, TrayOnly, ScanFile, SecureFile
#End If
#If False Then
Private Normal, TrayOnly, ScanFile
#End If
Private Type Lng
    Lanugage                                    As String
    Translator                                  As String
End Type
Private Type AntiVirus
    AVname                                      As String
    Runmode                                     As RM
    Signature                                   As TypeSignature
    Language                                    As Lng
End Type
Public AV                                   As AntiVirus
Private Type SHItemID
    cb                                          As Long
    abID                                        As Byte
End Type
Public Type ItemIDList
    mkid                                        As SHItemID
End Type
Public Type BROWSEINFO
    hOwner                                      As Long
    pidlRoot                                    As Long
    pszDisplayName                              As String
    lpszTitle                                   As String
    ulFlags                                     As Long
    lpCallbackProc                              As Long
    lParam                                      As Long
    iImage                                      As Long
End Type
Public Enum VirusT
    Executable = 0
    Script = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Executable, Script
#End If
Public Enum pStatus
    Max = 1
    Min = 0
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Max, Min
#End If
Public Const FileTypesofInterrest           As String = "EXEBATCOMPIFDOCVBS.JSSCR"
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Executable, Script
#End If
Private Type TypeVirus
    FileNameShort                               As String
    Reason                                      As String
    FileName                                    As String
Type                                        As VirusT
End Type
Public Type RECT
    Left                                        As Long
    Top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type
''Private Const IDANI_OPEN                      As Long = &H1
Public Const IDANI_CLOSE                    As Long = &H2
Public Const IDANI_CAPTION                  As Long = &H3
Public Virus                                As TypeVirus
#If Win16 Then
Public Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, _
                                            ByVal hWndInsertAfter As Integer, _
                                            ByVal X As Integer, _
                                            ByVal Y As Integer, _
                                            ByVal cx As Integer, _
                                            ByVal cy As Integer, _
                                            ByVal wFlags As Integer)
#Else
Public Declare Function SetWindowPos Lib "USER32" (ByVal hwnd As Long, _
                                                   ByVal hWndInsertAfter As Long, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long, _
                                                   ByVal cx As Long, _
                                                   ByVal cy As Long, _
                                                   ByVal wFlags As Long) As Long
#End If
Public Type FILETIME
    dwLowDateTime                               As Long
    dwHighDateTime                              As Long
End Type
Public Const HKEY_LOCAL_MACHINE             As Long = &H80000002
Public Type WIN32_FIND_DATA
    dwFileAttributes                            As Long
    ftCreationTime                              As FILETIME
    ftLastAccessTime                            As FILETIME
    ftLastWriteTime                             As FILETIME
    nFileSizeHigh                               As Long
    nFileSizeLow                                As Long
    dwReserved0                                 As Long
    dwReserved1                                 As Long
    cFileName                                   As String * 260
    cAlternate                                  As String * 14
End Type
Public Const KEY_ALL_ACCESS                 As Long = &H3F
Public Const KEY_SET_VALUE                  As Long = &H2
Public Const KEY_CREATE_SUB_KEY             As Long = &H4
Public Const REG_PRIMARY_KEY                As String = "Software\Classes\"
Public Const REG_SHELL_KEY                  As String = "Shell\"
Public Const REG_SHELL_OPEN_KEY             As String = "Open\"
Public Const REG_SHELL_OPEN_COMMAND_KEY     As String = "Command"
Public Const REG_ICON_KEY                   As String = "DefaultIcon"
Public Const UpdateWebsite                  As String = "www.cts.sub.cc"
Public Const REG_SZ                         As Long = 1
Public Const REG_OPTION_NON_VOLATILE        As Long = 0
Public Const ERROR_SUCCESS                  As Long = 0
Public Type SECURITY_ATTRIBUTES
    nLength                                     As Long
    lpSecurityDescriptor                        As Long
    bInheritHandle                              As Boolean
End Type
Public SH                                   As New Shell    'reference to shell32.dll class
Public Type NOTIFYICONDATA
    cbSize                                      As Long
    hwnd                                        As Long
    uId                                         As Long
    uFlags                                      As Long
    ucallbackMessage                            As Long
    hIcon                                       As Long
    szTip                                       As String * 64
End Type
Type POINTAPI
    X                                           As Long
    Y                                           As Long
End Type
Public Enum Direction
    dirTop = 1
    dirBottom = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private dirTop, dirBottom
#End If
Public Enum lbBorderStyleTypes
    None = 0
    [Fixed Single] = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private None
#End If
Public Declare Function SetRect Lib "USER32" (lpRect As RECT, _
                                              ByVal X1 As Long, _
                                              ByVal Y1 As Long, _
                                              ByVal X2 As Long, _
                                              ByVal Y2 As Long) As Long
Public Declare Function DrawAnimatedRects Lib "USER32" (ByVal hwnd As Long, _
                                                        ByVal idAni As Long, _
                                                        lprcFrom As RECT, _
                                                        lprcTo As RECT) As Long
''Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
''Private Declare Function WindowFromPoint Lib "USER32" (ByVal lpPointX As Long, ByVal lpPointY As Long) As Long
Public Declare Function MoveWindow Lib "USER32" (ByVal hwnd As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal bRepaint As Long) As Long
''Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
''Private Declare Function GetLogicalDrives Lib "kernel32" () As Long
''Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
''Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function Beep Lib "KERNEL32" (ByVal dwFreq As Long, _
                                             ByVal dwDuration As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function GetShortPathName Lib "KERNEL32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                                                                   ByVal lpszShortPath As String, _
                                                                                   ByVal cchBuffer As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                               ByVal lpSubKey As String, _
                                                                               ByVal ulOptions As Long, _
                                                                               ByVal samDesired As Long, _
                                                                               phkResult As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                   ByVal lpSubKey As String, _
                                                                                   ByVal Reserved As Long, _
                                                                                   ByVal lpClass As String, _
                                                                                   ByVal dwOptions As Long, _
                                                                                   ByVal samDesired As Long, _
                                                                                   lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                   phkResult As Long, _
                                                                                   lpdwDisposition As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, _
                                                                             ByVal lpSubKey As Any, _
                                                                             ByVal dwType As Long, _
                                                                             ByVal lpData As String, _
                                                                             ByVal cbData As Long) As Long
Public Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, _
                                                                     riid As CLSIdType, _
                                                                     ByVal fown As Long, _
                                                                     lpUnk As Object) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                ByVal dwFileAttributes As Long, _
                                                                                psfi As ShellFileInfoType, _
                                                                                ByVal cbFileInfo As Long, _
                                                                                ByVal uFlags As Long) As Long
Public Declare Function GetFileNameFromBrowseW Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, _
                                                                          ByVal lpstrFile As Long, _
                                                                          ByVal nMaxFile As Long, _
                                                                          ByVal lpstrInitialDir As Long, _
                                                                          ByVal lpstrDefExt As Long, _
                                                                          ByVal lpstrFilter As Long, _
                                                                          ByVal lpstrTitle As Long) As Long
Public Declare Function GetFileNameFromBrowseA Lib "shell32" Alias "#63" (ByVal hwndOwner As Long, _
                                                                          ByVal lpstrFile As String, _
                                                                          ByVal nMaxFile As Long, _
                                                                          ByVal lpstrInitialDir As String, _
                                                                          ByVal lpstrDefExt As String, _
                                                                          ByVal lpstrFilter As String, _
                                                                          ByVal lpstrTitle As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
Public Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function FindFirstFileA Lib "KERNEL32" (ByVal lpFileName As String, _
                                                       lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFileA Lib "KERNEL32" (ByVal hFindFile As Long, _
                                                      lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributesA Lib "KERNEL32" (ByVal lpFileName As String) As Long


