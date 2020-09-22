VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Virus Search"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  '2D
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "c:\"
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Search in:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    txtPath.Text = BrowseForFolder(Me.hwnd, "Please select the Path you want to search in!")
End Sub

