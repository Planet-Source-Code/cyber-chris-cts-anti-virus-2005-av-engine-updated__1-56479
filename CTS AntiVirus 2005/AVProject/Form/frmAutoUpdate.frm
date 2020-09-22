VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmAutoUpdate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Update Reminder"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   5106
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
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
''Private iewindow         As InternetExplorer    ':( Missing Scope
''Private currentwindows   As New ShellWindows
Private WithEvents Doc   As MSHTML.HTMLDocument
Attribute Doc.VB_VarHelpID = -1
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Function Doc_onclick() As Boolean
    If Doc.activeElement Is Nothing Then
        Exit Function '>---> Bottom
    Else
        If LenB(Doc.activeElement.Id) Then
            Select Case Doc.activeElement.Id
            Case "yes"
                frmUpdate.Show , Me
                Unload Me
            Case "no"
                Unload Me
            End Select
        End If
    End If
    Doc_onclick = True
End Function
Private Sub Form_Load()
    wb.Navigate2 App.path & "\gui\signatureupdate.htm"
End Sub
Private Sub wb_NavigateComplete2(ByVal pDisp As Object, _
                                 URL As Variant)
    Set Doc = wb.Document
End Sub
''
''Private Sub txtRemindLater_KeyDown(ByVal KeyCode As Long, ByVal Shift As Long)
''
''
''
''
''
''
''
''
'''48..57
''End Sub
''


