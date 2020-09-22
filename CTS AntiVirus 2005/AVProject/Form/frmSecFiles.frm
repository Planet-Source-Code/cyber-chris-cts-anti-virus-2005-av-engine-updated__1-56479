VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSecFiles 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Secured Files"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmSecFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton cmdDesecure 
      Caption         =   "&Desecure"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox flSec 
      Height          =   1455
      Left            =   8880
      Pattern         =   "*.secure"
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ListView lvQuarintine 
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   6455
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Type"
         Object.Width           =   3810
      EndProperty
   End
End
Attribute VB_Name = "frmSecFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private Sub cmdDesecure_Click()
Dim sXor As New clsSimpleXOR
    If MsgBox("Do you really want to desecure the file?", vbYesNo + vbCritical) = vbYes Then
        Log "File desecured: " & flSec.FileName, 2
        With App
            FileCopy .path & "\secure\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1) & ".secure", .path & "\secure\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1)
            Kill .path & "\secure\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1) & ".secure"
            sXor.DecryptFile .path & "\secure\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1), .path & "\secure\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1), AV.AVname
        End With
        Set sXor = Nothing
        flSec.Refresh
        Form_Load
        SaveSetting AV.AVname, "Settings", "Quarintine", flSec.ListCount
    End If
End Sub
Private Sub cmdRemove_Click()
    AVE.RemoveFileWithDialog App.path & "\secure\" & lvQuarintine.SelectedItem & "." & lvQuarintine.SelectedItem.SubItems(1) & ".secure", Me.hwnd
End Sub
Private Sub Form_Load()
Dim CurEntry As ListItem
Dim Counter  As Long
'translatelabel cmdRemove, 112
'translatelabel cmdDesecure, 113
    Caption = LoadResString(114)
    Me.flSec.path = App.path & "\secure\"
    flSec.Refresh
    For Counter = 0 To flSec.ListCount - 1
        Set CurEntry = Me.lvQuarintine.ListItems.Add
        CurEntry.Text = Mid$(flSec.List(Counter), 1, InStr(1, flSec.List(Counter), ".") - 1)
        CurEntry.SubItems(1) = Mid$(flSec.List(Counter), InStr(1, flSec.List(Counter), ".") + 1, 3)
    Next '  COUNTER COUNTER COUNTER COUNTER
End Sub
Private Sub Form_Unload(cancel As Integer)
    SaveSetting AV.AVname, "Settings", "Quarintine", flSec.ListCount
End Sub
