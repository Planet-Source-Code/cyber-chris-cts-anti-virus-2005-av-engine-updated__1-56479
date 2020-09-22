VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "About"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "enginefrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "www.cts.sub.cc       cyber_chris235@gmx.net"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright by Cyber Chris"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblText 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Corona"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "CTS Antivirus Engine 2005"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
End Sub
Private Sub Form_Load()
    Me.lblText(0).Caption = App.Major & "." & App.Minor & "." & App.Revision & " Corona"
End Sub
