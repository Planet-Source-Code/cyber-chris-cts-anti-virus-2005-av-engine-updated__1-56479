Attribute VB_Name = "modError"
Option Explicit
Public Sub ErrorFunc(ByVal Err_Number As Long, _
                     Err_Description As String, _
                     Err_Routine As String, _
                     Optional ByVal RoutineVariables As String)
    Debug.Print Now & " Error occured! System halted"
    Log LoadResString(154) & ": " & Err_Description & " " & Err_Routine, 3, True
    MsgBox LoadResString(153) & ": " & Err_Description, vbCritical + vbOKOnly
End Sub


