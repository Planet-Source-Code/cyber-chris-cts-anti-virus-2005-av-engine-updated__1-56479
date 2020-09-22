Attribute VB_Name = "modInterface"
Option Explicit
Public Currentpage   As String
Public Report        As String
Public Sub LoadPage(ByVal StrPath As String)
Dim file
Dim temp  As String
Dim Temp2 As String
Dim i     As Long
    file = FreeFile
    Open StrPath For Binary As #1
    temp = Input$(FileLen(StrPath), #1)
    Close #1
    With AV
        temp = Replace(temp, "#001#", GetSetting(.AVname, "Settings", "Auto Protect", "ON"), , 1)
        temp = Replace(temp, "#002#", GetSetting(.AVname, "Settings", "Startup", "OFF"), , 1)
        temp = Replace(temp, "#003#", GetSetting(.AVname, "Settings", "LogFile", "OFF"), , 1)
        temp = Replace(temp, "#004#", GetSetting(.AVname, "Settings", "Quarintine", 0), , 1)
        temp = Replace(temp, "#005#", AVE.SignatureDate, , 1)
        temp = Replace(temp, "#006#", AVE.SignatureCount, , 1)
        temp = Replace(temp, "#007#", GetSetting(.AVname, "Settings", "countFiles", 0), , 1)
        temp = Replace(temp, "#008#", GetSetting(.AVname, "Settings", "countVirus", 0), , 1)
        If DateDiff("d", AVE.SignatureDate, CDate(Date)) > 5 Then
            temp = Replace(temp, "#009#", " color: #FF0000; ", , 1)
        Else
            temp = Replace(temp, "#009#", vbNullString, , 1)
        End If
    End With
    Temp2 = ""
    For i = 0 To (frmMain.lstFiles.ListCount - 1)
        Temp2 = Temp2 & vbNewLine & "<p class=""""Stil11"""">" & frmMain.lstFiles.List(i) & "</p>"
    Next i
    temp = Replace(temp, "#010#", Temp2, , 1)
    If LenB(Report) Then
        temp = Replace(temp, "#011#", Report, , 1)
        Report = ""
    End If
    file = FreeFile
    Open App.path & "\gui\temp.000" For Append As file
    Print #file, temp
    Close file
    DoEvents
    frmMain.mainBrowser.Navigate App.path & "\gui\temp.000"
    Currentpage = StrPath
    DoEvents
    Kill App.path & "\gui\temp.000"
End Sub
