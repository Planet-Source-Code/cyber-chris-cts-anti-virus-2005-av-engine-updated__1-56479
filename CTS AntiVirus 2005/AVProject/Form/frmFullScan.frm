VERSION 5.00
Begin VB.Form frmFullScan 
   Caption         =   "Drive Scan"
   ClientHeight    =   3570
   ClientLeft      =   2655
   ClientTop       =   3390
   ClientWidth     =   6750
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   3570
   ScaleWidth      =   6750
   Begin VB.ListBox List1 
      Height          =   2985
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   6690
      TabIndex        =   0
      Top             =   3315
      Width           =   6750
   End
   Begin VB.Menu mnuFindFiles 
      Caption         =   "Find File(s)..."
   End
   Begin VB.Menu mnuFolderInfo 
      Caption         =   "Folder &Info..."
   End
End
Attribute VB_Name = "frmFullScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PicHeight%, hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%

Dim WFD As WIN32_FIND_DATA, hItem&, hFile&

Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46
 

Private Sub Form_Load()
    ScaleMode = vbPixels
    PicHeight% = Picture1.Height
    hLB& = List1.hwnd
    TranslateLabel frmfullsearch, 149
    SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
    Move (Screen.Width - Width) * 0.5, (Screen.Height - Height) * 0.5
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Running% Then Running% = False
End Sub

Private Sub Form_Resize()
    MoveWindow hLB&, 0, 0, ScaleWidth, ScaleHeight - PicHeight%, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    End
End Sub

Private Sub mnuFindFiles_Click()
  
    If Running% Then: Running% = False: Exit Sub
    
    Dim drvbitmask&, maxpwr%, pwr%
    On Error Resume Next
    
    FileSpec$ = InputBox("Enter a file spec:" & vbCrLf & vbCrLf & _
                                    "Searching will begin at drive A and continue " & _
                                    "until no more drives are found.  " & _
                                    "Click Stop! at any time." & vbCrLf & _
                                    "The * and ? wildcards can be used.", _
                                    "Find File(s)", "*.exe")
    
    If Len(FileSpec$) = 0 Then Exit Sub
    
    MousePointer = 11
    Running% = True
    UseFileSpec% = True
    mnuFindFiles.Caption = "&Stop!"
    mnuFolderInfo.Enabled = False
    List1.Clear
       
    drvbitmask& = GetLogicalDrives()
    If drvbitmask& Then
        
        maxpwr% = Int(Log(drvbitmask&) / Log(2))   ' a little math...
        For pwr% = 0 To maxpwr%
            If Running% And (2 ^ pwr% And drvbitmask&) Then _
                Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
        Next
    End If
    
    Running% = False
    UseFileSpec% = False
    mnuFindFiles.Caption = "&Find File(s)..."
    mnuFolderInfo.Enabled = True
    MousePointer = 0

    Picture1.Cls
    Picture1.Print "Find File(s): " & List1.ListCount & " items found matching " & """" & FileSpec$ & """"
    Beep
    
End Sub

Private Sub mnuFolderInfo_Click()

    If Running% Then: Running% = False: Exit Sub
    
    Dim searchpath$
    On Error Resume Next

    searchpath$ = InputBox("Enter a valid explicit path:", "Folder Info", "C:\")
    If Len(searchpath$) < 2 Then Exit Sub
    If Mid$(searchpath$, 2, 1) <> ":" Then Exit Sub
    
    If Right$(searchpath$, 1) <> vbBackslash Then searchpath$ = searchpath$ & vbBackslash
    If FindClose(FindFirstFile(searchpath$ & vbAllFiles, WFD)) = False Then
        MsgBox searchpath$, vbInformation, "Path is invalid": Exit Sub
    End If

    MousePointer = 11
    Running% = True
    mnuFolderInfo.Caption = "&Stop!"
    mnuFindFiles.Enabled = False
    List1.Clear

    TotalDirs% = 0
    TotalFiles% = 0
    Call SearchDirs(searchpath$)
    
    Running% = False
    mnuFolderInfo.Caption = "&Folder Info..."
    mnuFindFiles.Enabled = True
    Picture1.Cls
    MousePointer = 0

    MsgBox "Total folders: " & vbTab & TotalDirs% & vbCrLf & _
                 "Total files: " & vbTab & TotalFiles%, , _
                 "Folder Info for: " & searchpath$
    
End Sub
 

Private Sub SearchDirs(curpath$)
    Dim dirs%, dirbuf$(), i%
    
    Picture1.Cls
    Picture1.Print "Searching " & curpath$
    
    DoEvents
    If Not Running% Then Exit Sub
    
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    TotalDirs% = TotalDirs% + 1
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        
        Loop While FindNextFile(hItem&, WFD)
        
        Call FindClose(hItem&)
    
    End If

    If UseFileSpec% Then
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
  
End Sub

Private Sub SearchFileSpec(curpath$)
    
    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        
        Do
            DoEvents
            If Not Running% Then Exit Sub
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        Loop While FindNextFile(hFile&, WFD)
        Call FindClose(hFile&)
    
    End If

End Sub
