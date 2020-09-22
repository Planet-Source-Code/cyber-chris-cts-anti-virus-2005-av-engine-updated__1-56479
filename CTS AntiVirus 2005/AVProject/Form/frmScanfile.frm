VERSION 5.00
Begin VB.Form frmScanfile 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Scaning File"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrUnload 
      Interval        =   1000
      Left            =   960
      Top             =   360
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin CCAntivir2004.ProgressBar pBar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      Picture         =   "frmScanfile.frx":0000
      ForeColor       =   0
      BarPicture      =   "frmScanfile.frx":001C
      Value           =   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
End
Attribute VB_Name = "frmScanfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnd_Click()
    Running = False
    MsgBox "Scan process terminated by user!", vbInformation
End Sub

Private Sub tmrUnload_Timer()
    If pBar.Text = "Scan complete!" Then Unload Me
End Sub
