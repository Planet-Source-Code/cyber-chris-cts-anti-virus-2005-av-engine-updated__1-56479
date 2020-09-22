VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Log View"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11940
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   11940
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btnDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   10560
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   10095
      Begin VB.TextBox tbDateFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tbDateTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tbTimeFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tbTimeTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tbData 
         Height          =   285
         Left            =   7680
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date : "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time : "
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFFFF&
         Caption         =   " to "
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFFFF&
         Caption         =   " to "
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data : "
         Height          =   255
         Left            =   6840
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox tbOutput 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmLog.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const Log_file   As String = "log.txt"
Public Sub AddLog(ByVal typ As Long, _
                  ByVal skt As Long, _
                  st As String)
Dim ns As String
    ns = Trim$(CStr(skt))
    If Len(ns) < 3 Then
        ns = Left$("000", 3 - Len(ns)) & ns
    End If
    st = Trim$(st)
    st = ns & ": [" & logType(typ) & "] " & st
    st = Format$(Now, "MM/DD/YYYY HH:MM") & " " & st
    WriteLog st
End Sub
Private Sub btnDone_Click()
    Form_Unload 0
End Sub
Private Sub btnSearch_Click()
    QueryLog
End Sub
Public Sub Form_Load()
'Dim i As Long
    FileCopy App.path & "/" & Log_file, App.path & "/server.bak"
End Sub
Private Sub Form_Unload(cancel As Integer)
    On Error Resume Next
    Kill App.path & "/server.bak"
    If err.Number <> 0 Then
        err.Clear
    End If
    Unload Me
    On Error GoTo 0
End Sub
Private Sub QueryLog()
Dim srchSdate As Date
Dim srchEdate As Date
Dim srchSock  As Long
Dim foundline
Dim srchMod   As String
Dim srchData  As String
Dim lineDate  As Date
Dim lineSock  As Long
Dim lineMod   As String
Dim lineData  As String
Dim st        As String
Dim temp      As String
Dim CRLF      As String
'Dim foundLin  As Boolean
''Dim ctr       As Long
'Dim i         As Long
'Dim e         As Long
    CRLF = vbNewLine
'CREATE QUERY
    If LenB(tbTimeFrom.Text) = 0 Then
        st = "00:00"
    Else
        st = tbTimeFrom.Text
    End If
    If LenB(tbDateFrom.Text) = 0 Then
        temp = "01/01/2000"
    Else
        temp = tbDateFrom.Text
    End If
    temp = temp & " " & st
    srchSdate = CDate(Format$(temp, "MM/DD/YYYY HH:MM"))
    If LenB(tbTimeTo.Text) = 0 Then
        If LenB(tbDateTo.Text) = 0 Then
            st = Format$(Now, "HH:MM")
        Else
            st = "23:59"
        End If
    Else
        st = tbTimeTo.Text
    End If
    If LenB(tbDateTo.Text) = 0 Then
        temp = Format$(Now, "MM/DD/YYYY")
    Else
        temp = tbDateTo.Text
    End If
    temp = temp & " " & st
    srchEdate = CDate(Format$(temp, "MM/DD/YYYY HH:MM"))
    srchSock = 0
    srchMod = "ALL"
    srchData = Trim$(tbData.Text)
'DISPLAY QUERY
    tbOutput.Text = ""
    tbOutput.Visible = False
    foundline = False
'ctr = 0
'DISPLAY QUERY RESULTS
    Open App.path & "/server.bak" For Input As #8
    Do While (Not EOF(8))
        st = ""
        temp = ""
        Do While (temp <> vbLf) And (Not EOF(8))
            temp = Input$(1, #8)
            st = st & temp
        Loop
        If (InStr(1, st, vbCr) > 0) Then
            st = Mid$(st, 1, InStr(1, st, vbCr) - 1)
        End If
        If (Trim$(CStr(Mid$(st, 1, 2))) > 0) Then
            lineDate = CDate(Format$(Mid$(st, 1, 16), "MM/DD/YYYY HH:MM"))
        Else
            lineDate = CDate(Format$("01/01/1990 01:00", "MM/DD/YYYY HH:MM"))
        End If
        lineSock = Trim$(CStr(Mid$(st, 18, 3)))
        lineMod = Mid$(st, 24, MAX_LEN)
        lineData = Trim$(Mid$(st, 26 + MAX_LEN, Len(st)))
        If (lineDate >= srchSdate) And (lineDate <= srchEdate) And ((lineSock = srchSock) Or (srchSock = 0)) And ((lineMod = srchMod) Or (srchMod = "ALL")) And (InStr(1, lineData, srchData) > 0) Then
            tbOutput.Text = tbOutput.Text & st & CRLF
            foundline = True
        End If
        DoEvents
    Loop
    Close #8
    If Not foundline Then
        tbOutput.Text = "No Lines match your Query."
    End If
    tbOutput.Visible = True
End Sub
Private Sub tbData_GotFocus()
    SendKeys "{home}+{end}"
End Sub
Private Sub tbDateFrom_GotFocus()
    SendKeys "{home}+{end}"
End Sub
Private Sub tbDateFrom_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And (KeyAscii <> 8) And (Chr$(KeyAscii) <> "/") Then
        KeyAscii = 0
    End If
End Sub
Private Sub tbDateTo_GotFocus()
    If LenB(tbDateFrom.Text) <> 0 Then
        If tbDateTo.Text = 0 Then
            tbDateTo.Text = tbDateFrom.Text
        End If
    End If
    SendKeys "{home}+{end}"
End Sub
Private Sub tbDateTo_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And (KeyAscii <> 8) And (Chr$(KeyAscii) <> "/") Then
        KeyAscii = 0
    End If
End Sub
Private Sub tbTimeFrom_GotFocus()
    SendKeys "{home}+{end}"
End Sub
Private Sub tbTimeFrom_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And (KeyAscii <> 8) And (Chr$(KeyAscii) <> ":") Then
        KeyAscii = 0
    End If
End Sub
Private Sub tbTimeTo_GotFocus()
    If LenB(tbDateFrom.Text) <> 0 Then
        If tbDateTo.Text = 0 Then
            tbTimeTo.Text = tbTimeFrom.Text
        End If
    End If
    SendKeys "{home}+{end}"
End Sub
Private Sub tbTimeTo_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And (KeyAscii <> 8) And (Chr$(KeyAscii) <> ":") Then
        KeyAscii = 0
    End If
End Sub
Public Sub WriteLog(st As String)
    st = Trim$(st)
    Open App.path & "\" & Log_file For Append As #9
    Print #9, st
    Close #9
End Sub
''
''Private Sub tbSocket_GotFocus()
''
''
''SendKeys "{home}+{end}"
''End Sub
''
''
''Private Sub tbSocket_KeyPress(KeyAscii As Long)
''
''
''
''If Not IsNumeric(Chr$(KeyAscii)) And (KeyAscii <> 8) Then
''KeyAscii = 0
''End If
''End Sub
''


