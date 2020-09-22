VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivir 2004"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdsecure 
      Caption         =   "&Secure"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   585
      Left            =   6000
      ScaleHeight     =   585
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
      Caption         =   "xxxxxx"
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
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
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
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "File size:"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   1215
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
      Caption         =   "xxxxxx"
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
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   4815
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
      Caption         =   "Virus found!"
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
Attribute VB_Name = "frmAlert"
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
Private Sub BuildAlert()
    On Error Resume Next
    lblText(1).Caption = Virus.FileName
    lblText(1).ToolTipText = Virus.FileName & "  (" & FileLen(Virus.FileName) & " Bytes )"
    lblText(2).Caption = Virus.Reason
    lblText(8).Caption = FileLen(Virus.FileName) & " Bytes"
    If Virus.Type = Executable Then
        lblText(7).Caption = "Executable File"
    End If
    If Virus.Type = Script Then
        lblText(7).Caption = "Script"
    End If
    picIcon.Picture = LoadIcon(Large, Virus.FileName)
    On Error GoTo 0
End Sub
Private Sub cmdIgnore_Click()
    Log "Alert ignored: " & Virus.Reason, 2
    Unload Me
End Sub
Private Sub cmdRemove_Click()
    Log "File removed: " & Virus.FileName, 2
    AVE.RemoveFileWithDialog Virus.FileName, Me.hwnd
    Unload Me
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
    BuildAlert
    KeepOnTop Me
    DoEvents
    BeepAlert
End Sub


