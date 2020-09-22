Attribute VB_Name = "modUpdate"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Sub DownloadFile(ByVal srcFileName As String, _
                        ByVal targetFileName As String)
'This Downloads the latest version from the Internet
Dim b() As Byte
Dim FID As Byte
    frmUpdate.DownStatus "Conecting..."
    b() = frmUpdate.Inet.OpenURL(srcFileName, icByteArray)
    FID = FreeFile
    Open targetFileName For Binary Access Write As #FID
    Put #FID, , b()
    Close #FID
    frmUpdate.DownStatus "Writing Data to HD..."
    DoEvents
End Sub


