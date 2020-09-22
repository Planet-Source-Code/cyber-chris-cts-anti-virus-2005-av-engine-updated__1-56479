VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Update"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin InetCtlsObjects.Inet Inet 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Sub DownStatus(ByVal strStatus As String)
    lblStatus.Caption = strStatus
End Sub
Private Sub Form_Load()
    On Error GoTo err
    DownStatus "Starting Download"
    DownloadFile AV.Signature.SignatureOnlineFilename, App.path & "\temp.$$$"
    If FileLen(App.path & "\temp.$$$") <> 0 Then
        DownStatus "Download Complete"
        Kill AV.Signature.SignatureFilename
        FileCopy App.path & "\temp.$$$", AV.Signature.SignatureFilename
        Kill App.path & "\temp.$$$"
        Main
    End If
Exit Sub
err:
    DownStatus "Error while downloading"
End Sub
