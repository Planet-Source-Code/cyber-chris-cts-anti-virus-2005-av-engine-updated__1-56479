VERSION 5.00
Begin VB.Form frmScriptAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CTS Antivirus 2005: Realtime Scriptchecker"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   825
      Left            =   6000
      Picture         =   "frmScriptAlert.frx":0000
      ScaleHeight     =   825
      ScaleWidth      =   645
      TabIndex        =   3
      Top             =   960
      Width           =   645
   End
   Begin VB.Line lline 
      Index           =   3
      X1              =   6840
      X2              =   6840
      Y1              =   2040
      Y2              =   480
   End
   Begin VB.Line lline 
      Index           =   2
      X1              =   480
      X2              =   6840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lline 
      Index           =   1
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   2040
   End
   Begin VB.Line lline 
      Index           =   0
      X1              =   480
      X2              =   6840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "The specified file contains one or more lines of destructive code! This might be a Virus. Please decide carefully how to proceed!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Possible Script Virus found!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmScriptAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sub BeepAlert()
    Beep 4000, 220
    Beep 3000, 200
    Beep 4000, 220
    Beep 3000, 200
End Sub
Private Sub cmdIgnore_Click()
    Log "Alert ignored: " & Virus.Reason, 2
    Unload Me
End Sub
Private Sub cmdRemove_Click()
    Log "File removed: " & lblText(1).Caption, 2
    AVE.RemoveFileWithDialog lblText(1).Caption, Me.hwnd
    End
End Sub
Private Sub cmdRun_Click()
Dim Result
Dim AA
Dim strResult As String
    strResult = MsgBox("Do you really want to execute that file?", vbCritical + vbYesNo, "Alert!")
    If strResult = vbYes Then
        AA = Space$(255)
        Result = GetShortPathName(lblText(1).Caption, AA, Len(AA))
        Shell "c:\windows\System32\WScript.exe """"" & Mid$(AA, 1, Result) & """"""
    End If
    End
End Sub
Private Sub cmdSecure_Click()
Dim sXor As New clsSimpleXOR
    On Error Resume Next
    MsgBox "The File will be secured, that means everytime you want to start it, you'll get a prompt." & vbNewLine & _
       "This will avoid unwanted starts!", vbInformation + vbOKOnly
    sXor.EncryptFile Virus.FileName, Virus.FileName, AV.AVname
    Set sXor = Nothing
    MkDir App.path & "\Secure\"
    FileCopy Virus.FileName, App.path & "\Secure\" & Mid$(Virus.FileNameShort, 1, Len(Virus.FileNameShort) - 1) & ".secure"
    Kill Virus.FileName
    With frmSecFiles
        .Visible = False
        .Show
        SaveSetting AV.AVname, "Settings", "Quarintine", .flSec.ListCount
    End With 'frmSecFiles
    Unload frmSecFiles
    Log "File moved to quarintine: " & Virus.FileName, 2
    On Error GoTo 0
End Sub
Private Sub Form_Load()
Dim R1  As RECT
Dim R2  As RECT
Dim TPP As Long
    TPP = Screen.TwipsPerPixelX
    SetRect R1, Screen.Width / TPP, Screen.Height / TPP, Screen.Width / TPP, Screen.Height / TPP
    SetRect R2, 0, 0, Me.Width / TPP, Me.Height / TPP
    DrawAnimatedRects Me.hwnd, IDANI_CLOSE Or IDANI_CAPTION, R1, R2
    KeepOnTop Me
'translatelabel hpOnline, 150
'translatelabel lblText(0), 103
'translatelabel lblText(3), 104
'translatelabel lblText(4), 105
'translatelabel lblText(6), 106
'translatelabel lblText(5), 107
'translatelabel cmdIgnore, 132
'translatelabel cmdRemove, 133
'translatelabel cmdSecure, 134
    DoEvents
    BeepAlert
End Sub
Private Sub hpOnline_Click()
'http://www.viruslist.com/eng/viruslistfind.html?findTxt=code+red
    ShellExecute Me.hwnd, "Open", "http://www.viruslist.com/eng/viruslistfind.html?findTxt=" & Replace(Virus.Reason, " ", "+"), vbNullString, "c:\", 1
End Sub


