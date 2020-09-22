VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "CTS Antivirus 2005"
   ClientHeight    =   4800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8175
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSearch.frx":0CCA
   ScaleHeight     =   4800
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   -1000
      Picture         =   "frmSearch.frx":3BEC
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      ToolTipText     =   "Planet Source Code Superior Coding Contest Winner"
      Top             =   2400
      Width           =   495
   End
   Begin VB.ListBox lstFiles 
      Height          =   255
      ItemData        =   "frmSearch.frx":47CE
      Left            =   -1000
      List            =   "frmSearch.frx":47D0
      TabIndex        =   21
      Top             =   480
      Width           =   135
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   3920
      Left            =   2880
      ScaleHeight     =   3915
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   280
      Width           =   135
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   75
         X2              =   75
         Y1              =   15
         Y2              =   3900
      End
      Begin VB.Line border 
         BorderColor     =   &H00E0E0E0&
         X1              =   150
         X2              =   135
         Y1              =   3870
         Y2              =   15
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   7920
      Top             =   9600
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   12480
      ScaleHeight     =   2955
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   10080
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan the selected File"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checksum:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filesize:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   825
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      Picture         =   "frmSearch.frx":47D2
      ScaleHeight     =   255
      ScaleWidth      =   8415
      TabIndex        =   19
      Top             =   0
      Width           =   8415
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      Picture         =   "frmSearch.frx":DA54
      ScaleHeight     =   615
      ScaleWidth      =   8175
      TabIndex        =   17
      Top             =   4200
      Width           =   8175
      Begin SHDocVwCtl.WebBrowser mnuWebUpdate 
         Height          =   30
         Left            =   -1000
         TabIndex        =   20
         Top             =   -10000
         Width           =   30
         ExtentX         =   53
         ExtentY         =   53
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin SHDocVwCtl.WebBrowser menuBrowser 
      CausesValidation=   0   'False
      Height          =   3975
      Left            =   0
      TabIndex        =   16
      Top             =   240
      Width           =   2925
      ExtentX         =   5159
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser mainBrowser 
      CausesValidation=   0   'False
      Height          =   3975
      Left            =   2880
      TabIndex        =   8
      Top             =   240
      Width           =   5415
      ExtentX         =   9551
      ExtentY         =   7011
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   0
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   8
      X1              =   6840
      X2              =   6840
      Y1              =   -480
      Y2              =   240
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   7
      X1              =   6360
      X2              =   6360
      Y1              =   -480
      Y2              =   240
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuLMainWindow 
         Caption         =   "Open CTS AntiVirus"
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSF 
         Caption         =   "Scan a File"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit CC Antivir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private WithEvents X              As cCommonDialog
Attribute X.VB_VarHelpID = -1
Private CurrentFile               As String
Private sPicScan                  As pStatus
Private SpicOther                 As pStatus
Private sHelpAbout                As pStatus
Private Currentpage               As String
Private isoverlabel               As Boolean             ':( Missing Scope
Private iewindow                  As InternetExplorer    ':( Missing Scope
Private currentwindows            As New ShellWindows
Private WithEvents Doc            As MSHTML.HTMLDocument
Attribute Doc.VB_VarHelpID = -1
Private WithEvents DocMenu        As MSHTML.HTMLDocument
Attribute DocMenu.VB_VarHelpID = -1
Private Const NIM_ADD             As Long = &H0&
''Private Const NIM_MODIFY          As Long = &H1&
Private Const NIM_DELETE          As Long = &H2&
Private Const NIF_MESSAGE         As Long = &H1&
Private Const NIF_ICON            As Long = &H2&
Private Const NIF_TIP             As Long = &H4&
Private Const WM_MOUSEMOVE        As Long = &H200&
Private Const WM_LBUTTONDOWN      As Long = &H201&
Private Const WM_LBUTTONUP        As Long = &H202&
Private Const WM_LBUTTONDBLCLK    As Long = &H203&
Private Const WM_RBUTTONDOWN      As Long = &H204&
Private Const WM_RBUTTONUP        As Long = &H205&
Private Const WM_RBUTTONDBLCLK    As Long = &H206&
Private Type NOTIFYICONDATA
    cbSize                            As Long
    hwnd                              As Long
    uId                               As Long
    uFlags                            As Long
    ucallbackMessage                  As Long
    hIcon                             As Long
    szTip                             As String * 64
End Type
Private TIcon                     As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                   pnid As NOTIFYICONDATA) As Boolean
Private Sub cmdScan_Click()
    CheckFile (CurrentFile)
End Sub
Private Sub Form_Load()
'On Error GoTo err
'Call lblFileScan_Click
    Me.Hide
    App.TaskVisible = False
'mnBar.Visible = False
    With TIcon
        .cbSize = Len(TIcon)
        .hwnd = Me.picIcon.hwnd
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
        .hIcon = frmMain.Icon
        .szTip = "CTS AnChr$(0)irus 2005" & vbNullChar
    End With 'TIcon
    Shell_NotifyIcon NIM_ADD, TIcon
    With App
        LoadPage (.path & "\Gui\SystemStatus.htm")
        Currentpage = .path & "\Gui\SystemStatus.htm"
        menuBrowser.Navigate2 .path & "\gui\menu.htm"
    End With 'App
    logType(1) = "Virus found"
    logType(2) = "Virus action"
    logType(3) = "Error"
    Set X = New cCommonDialog
    Set ccClass = X
    frmMain.Cls
    sPicScan = Max
    SpicOther = Min
    sHelpAbout = Min
    If DateDiff("d", AVE.SignatureDate, CDate(Date)) > 5 Then
        If DateDiff("d", GetSetting(AV.AVname, "Settings", "RemindLater", Date), Date) >= 0 Then
            frmAutoUpdate.Show , Me
        End If
    End If
    Debug.Print DateDiff("d", GetSetting(AV.AVname, "Settings", "RemindLater", Date), Date)
    mnuWebUpdate.Navigate2 UpdateWebsite
Exit Sub
err:
    ErrorFunc err.Number, err.Description, "frMmain.Startup"
End Sub
Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    isoverlabel = True
End Sub
Private Sub Form_QueryUnload(cancel As Integer, _
                             UnloadMode As Integer)
    frmMain.Hide
    If UnloadMode = vbAppWindows Or UnloadMode = vbFormCode Then
        Shell_NotifyIcon NIM_DELETE, TIcon
    Else
        cancel = 1
    End If
End Sub
Private Sub Form_Unload(cancel As Integer)
    End
End Sub
Private Sub lblBug_Click()
    ShellExecute Me.hwnd, "Open", "mailto:cyber_chris235@gmx.net?subject=Bug in " & AV.AVname, vbNullString, "c:\", 1
    MsgBox LoadResString(128) '"Thank you for your help!"
End Sub
Private Sub mainBrowser_NavigateComplete2(ByVal pDisp As Object, _
                                          URL As Variant)
    Set Doc = mainBrowser.Document
End Sub
Private Sub menuBrowser_NavigateComplete2(ByVal pDisp As Object, _
                                          URL As Variant)
    Set DocMenu = menuBrowser.Document
End Sub
Private Sub mnuExit_Click()
    End
End Sub
Private Sub mnuLMainWindow_Click()
    Me.Show
End Sub
Private Sub mnuSF_Click()
    ShowFileSearch
End Sub
Private Sub picAbout_Click()
    frmAbout.Show , Me
End Sub
Private Sub picFastSearchx_Click(lngIndex As Long)
    X.ControlToSetNewParent = Picture1
    Debug.Print X.ShowOpen(Me.hwnd)
End Sub
Private Sub picFileSearch_Click()
    ShowFileSearch
End Sub
Private Sub picIcon_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
Dim Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
    Case WM_MOUSEMOVE
    Case WM_LBUTTONDBLCLK
        Me.Show
    Case WM_LBUTTONDOWN
    Case WM_LBUTTONUP
    Case WM_RBUTTONDBLCLK
        Me.Show
    Case WM_RBUTTONDOWN
    Case WM_RBUTTONUP
        Me.PopupMenu mnuMenu
    End Select
End Sub
Private Sub picPathsearch_Click()
End Sub
Private Sub picSec_Click()
    frmSecFiles.Show
End Sub
Private Sub Picture2_Click()
    ShellExecute Me.hwnd, "Open", "www.cts.sub.cc", vbNullString, "c:\", 1
End Sub
Private Sub picUpdate_Click()
    frmUpdate.Show , Me
End Sub
Public Sub ShowFileSearch()
Dim strFileName As String
    On Error Resume Next
    strFileName = (ShowOpenDlg(Me, , LoadResString(129) & "|*.*", , "Scan File")) 'All files
    Debug.Print Len(strFileName)
    If Len(strFileName) > 3 Then 'avoids Bug on Cancel
        If FileLen(strFileName) <> 0 Then
            CheckFile strFileName, True
        End If
    End If
    On Error GoTo 0
End Sub
Private Sub Timer1_Timer()
Dim currentlocation As String
Dim buffer          As String
Dim ValidData       As String
'Dim c               As Collection
    On Error Resume Next
    Timer1.Enabled = False
    For Each iewindow In currentwindows
        DoEvents
        If iewindow.Busy Then
            GoTo busysignal
        End If
        currentlocation = iewindow.LocationURL
        ValidData = InStr(1, buffer, iewindow.LocationName & "|" & iewindow.LocationURL & "|")
        If ValidData = 0 Then
            If Mid$(currentlocation, 1, 7) = "file://" Then
                currentlocation = Replace(currentlocation, "file:///", vbNullString)
                currentlocation = Replace(currentlocation, "%20", " ")
                currentlocation = Replace(currentlocation, "/", "\")
                currentlocation = Replace(currentlocation, "%C1", "Á")
                currentlocation = Replace(currentlocation, "%C9", "É")
                currentlocation = Replace(currentlocation, "%CD", "Í")
                currentlocation = Replace(currentlocation, "%D3", "Ó")
                currentlocation = Replace(currentlocation, "%DA", "Ú")
                currentlocation = Replace(currentlocation, "%E1", "á")
                currentlocation = Replace(currentlocation, "%E9", "é")
                currentlocation = Replace(currentlocation, "%ED", "í")
                currentlocation = Replace(currentlocation, "%F3", "ó")
                currentlocation = Replace(currentlocation, "%FA", "ú")
                currentlocation = Replace(currentlocation, "%E3", "ã")
                currentlocation = Replace(currentlocation, "%F5", "õ")
                currentlocation = Replace(currentlocation, "%C3", "Ã")
                currentlocation = Replace(currentlocation, "%D5", "Õ")
                currentlocation = Replace(currentlocation, "%E0", "à")
                currentlocation = Replace(currentlocation, "%E8", "è")
                currentlocation = Replace(currentlocation, "%EC", "ì")
                currentlocation = Replace(currentlocation, "%F2", "ò")
                currentlocation = Replace(currentlocation, "%F9", "ù")
                currentlocation = Replace(currentlocation, "%C0", "À")
                currentlocation = Replace(currentlocation, "%C8", "È")
                currentlocation = Replace(currentlocation, "%CC", "Ì")
                currentlocation = Replace(currentlocation, "%D2", "Ò")
                currentlocation = Replace(currentlocation, "%D9", "Ù")
                If currentlocation = "c:\ziptemp" Then
                    Exit Sub
                End If
                AVE.ScanFolder currentlocation
                Debug.Print currentlocation
            End If
        End If
busysignal:
    Next ':( Repeat For-Variable: IEWINDOW IEWINDOW IEWINDOW IEWINDOW
    Timer1.Enabled = True
    On Error GoTo 0
End Sub
Private Sub X_FileChanged(ByVal FileName As String)
    lblFileName.Caption = Mid$(FileName, InStrRev(FileName, "\") + 1)
    lblText(12).Caption = FileLen(FileName) & " Bytes"
    lblText(14).Caption = CalcCRC(FileName)
    CurrentFile = FileName
End Sub

Private Function Doc_onclick() As Boolean


If Doc.activeElement Is Nothing Then
Exit Function '>---> Bottom
Else
If LenB(Doc.activeElement.Id) Then
Select Case Doc.activeElement.Id
Case "auto_protection"
If GetSetting(AV.AVname, "Settings", "Auto Protect", "ON") = "ON" Then
SaveSetting AV.AVname, "Settings", "Auto Protect", "OFF"
Timer1.Enabled = False
Else
SaveSetting AV.AVname, "Settings", "Auto Protect", "ON"
Timer1.Enabled = True
End If
Case "run_on_startup"
With AV
If GetSetting(.AVname, "Settings", "Startup", "OFF") = "OFF" Then
SetKeyValue &H80000001, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", .AVname, App.path & "\" & App.EXEName & ".exe /T", 1
SaveSetting .AVname, "Settings", "Startup", "ON"
Else 'NOT GETSETTING(AV.AVNAME,...
DeleteValue &H80000001, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", .AVname
SaveSetting .AVname, "Settings", "Startup", "OFF"
End If
End With
Case "logging"
If GetSetting(AV.AVname, "Settings", "LogFile", "OFF") = "OFF" Then
SaveSetting AV.AVname, "Settings", "LogFile", "ON"
Else
SaveSetting AV.AVname, "Settings", "LogFile", "OFF"
End If
Case "quarantine"
frmSecFiles.Show , Me
Case "update_now"
frmUpdate.Show , Me
Case "scan_in_files"
ShowFileSearch
Case "full_path_scan"
Checkfolder
Case "fast_file_scan"
X.ControlToSetNewParent = Picture1
Debug.Print X.ShowOpen(Me.hwnd)
Case "password_safe"

End Select
End If
End If
LoadPage (Currentpage)
Doc_onclick = True
End Function


Private Function DocMenu_onclick() As Boolean



If DocMenu.activeElement Is Nothing Then
Exit Function '>---> Bottom
Else
If LenB(DocMenu.activeElement.Id) Then
Select Case DocMenu.activeElement.Id
Case "status"
LoadPage (App.path & "\Gui\SystemStatus.htm")
Currentpage = App.path & "\Gui\SystemStatus.htm"
If DocMenu.All("statistic").parentElement.Style.display = "none" Then
DocMenu.All("statistic").parentElement.Style.display = "block"
DocMenu.All("log").parentElement.Style.display = "block"
Else 'NOT (DOCMENU.ALL("STATISTIC").PARENTELEMENT.STYLE.DISPLAY...
DocMenu.All("statistic").parentElement.Style.display = "none"
DocMenu.All("log").parentElement.Style.display = "none"
End If
Case "statistic"
LoadPage (App.path & "\Gui\statistic.htm")
Currentpage = App.path & "\Gui\statistic.htm"
Case "scan_for_viruses"
LoadPage (App.path & "\Gui\filescan.htm")
Currentpage = App.path & "\Gui\filescan.htm"
With DocMenu
If .All("scan_in_files").parentElement.Style.display = "none" Then
.All("scan_in_files").parentElement.Style.display = "block"
.All("full_path_scan").parentElement.Style.display = "block"
.All("fast_file_scan").parentElement.Style.display = "block"
Else 'NOT (DOCMENU.ALL("SCAN_IN_FILES").PARENTELEMENT.STYLE.DISPLAY...
.All("scan_in_files").parentElement.Style.display = "none"
.All("full_path_scan").parentElement.Style.display = "none"
.All("fast_file_scan").parentElement.Style.display = "none"
End If
End With 'DocMenu
Case "scan_in_files"
ShowFileSearch
Case "full_path_scan"
Checkfolder
Case "fast_file_scan"
X.ControlToSetNewParent = Picture1
Debug.Print X.ShowOpen(Me.hwnd)
Case "update"
LoadPage (App.path & "\Gui\update.htm")
Currentpage = App.path & "\Gui\update.htm"
Case "quarintine"
LoadPage (App.path & "\Gui\quarintine.htm")
Currentpage = App.path & "\Gui\quarintine.htm"
Case "help"
LoadPage (App.path & "\Gui\help.htm")
Currentpage = App.path & "\Gui\help.htm"
Case "about"
frmAbout.Show , Me
Case "log"
frmLog.Show , Me
Case "tools"
LoadPage (App.path & "\Gui\tools.htm")
Currentpage = App.path & "\Gui\tools.htm"
End Select
End If
End If
DocMenu_onclick = True
End Function
''
''
''Private Sub lblAbout_Click()
''
''
''picAbout_Click
''End Sub
''
''
''Private Sub lblCFP_Click()
''
''
''picPathsearch_Click
''End Sub
''
''
''Private Sub lblffs_Click()
''
''
''picFastSearchx_Click 0
''End Sub
''
''
''Private Sub lblLogFile_Click()
''
''
''frmLog.Show
''End Sub
''
''
''Private Sub lblSecured_Click()
''
''
''picSec_Click
''End Sub
''
''
''Private Sub lblSif_Click()
''
''
''picFileSearch_Click
''End Sub
''
''
''Private Sub lblupdate_Click()
''
''
''picUpdate_Click
''End Sub
''


