VERSION 5.00
Begin VB.Form frmLanguage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Language"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin CCAntivir2004.DMSXpButton cmdOK 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Ok"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.FileListBox flst 
      Height          =   1845
      Left            =   600
      Pattern         =   "*.lng"
      TabIndex        =   2
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ComboBox cmbLanguage 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please select the Language you want to use:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
If cmbLanguage.Text <> "" Then
    AV.Language.Lanugage = cmbLanguage.Text
    SaveSetting AV.AVname, "Settings", "Language", cmbLanguage.Text
    modLanguage.BuildTranslation
    BuildUI
    MsgBox "Changes need to restart the program to take effect!", vbOKOnly + vbInformation, "Lanugage"
    End
    Unload Me
Else
    MsgBox "Please select a Language!", vbOKOnly + vbInformation, "Lanugage"
End If
End Sub

Private Sub Form_Load()
Dim count As Integer
    KeepOnTop Me
    flst.path = App.path & "\language\"
    flst.Refresh
    cmbLanguage.Clear
    For count = 0 To flst.ListCount - 1
        If flst.List(count) <> ".lng" Then
            cmbLanguage.AddItem UCase(Mid((flst.List(count)), 1, 1)) & Mid(flst.List(count), 2, Len(flst.List(count)) - 5)
        End If
    Next count
End Sub
