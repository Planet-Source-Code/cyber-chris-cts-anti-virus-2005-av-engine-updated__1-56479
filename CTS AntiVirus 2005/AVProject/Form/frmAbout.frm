VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton lblAboutTheEngine 
      Caption         =   "&About the Engine"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   4455
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3015
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      ExtentX         =   12303
      ExtentY         =   5318
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
      Location        =   "http:///"
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   8520
      Y1              =   3120
      Y2              =   3120
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    wb.Navigate App.path & "\gui\about.htm"
End Sub
Private Sub Form_Unload(cancel As Integer)
Dim myArticleAddr As String
    If MsgBox("Would you please vote on PSC Website in case you like this program?", vbQuestion + vbYesNo, "Your vote will be very well appreciated ...") = vbYes Then
        myArticleAddr = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=56479&optCodeRatingValue=5"
        ShellExecute Me.hwnd, "Open", myArticleAddr, vbNullString, vbNullString, 1
        MsgBox "Thank you very much. I really appreciate that :-) ", , "Thanks a million..."
    End If
End Sub
Private Sub lblAboutTheEngine_Click()
    AVE.About
End Sub
''
''Private Sub cmdExit_Click()
''
''
''Unload Me
''End Sub
''
''
''Private Sub lblCopyright_Click(lngIndex As Long)
''
''
''
''
''
''ShellExecute Me.hwnd, "Open", "mailto:cyber_chris235@gmx.net", vbNullString, "c:\", 1
''
''End Sub
''
''
''Private Sub lbllCopyright_Click()
''
''
''ShellExecute Me.hwnd, "Open", "mailto:cyber_chris235@gmx.net", vbNullString, "c:\", 1
''
''End Sub
''
''
''Private Sub lblthanks2_Click()
''
''
''ShellExecute Me.hwnd, "Open", "mailto:wpsjr1@succeed.net", vbNullString, "c:\", 1
''
''End Sub
''
''
''Private Sub lblThanks3_Click()
''
''
''ShellExecute Me.hwnd, "Open", "mailto:sharmaq@terra.com.br", vbNullString, "c:\", 1
''
''End Sub
''
''
''Private Sub lblThanks_Click()
''
''
''ShellExecute Me.hwnd, "Open", "mailto:dude@patabugen.co.uk", vbNullString, "c:\", 1
''
''End Sub
''
