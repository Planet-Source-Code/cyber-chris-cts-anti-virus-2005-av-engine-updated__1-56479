VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_hDC       As Long
Private m_hBmp      As Long
Private m_hBmpOld   As Long
Private m_lWidth    As Long
Private m_lHeight   As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 ByVal lpDeviceName As String, _
                                                                 ByVal lpOutput As String, _
                                                                 lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Sub Class_Terminate()
    pDestroy
End Sub
Public Property Get hdc(lngWidth As Long, _
                        lngHeight As Long) As Long
    If m_hDC = 0 Then
        pCreate lngWidth, lngHeight
    Else
        If lngWidth > m_lWidth Or lngHeight > m_lHeight Then
            pCreate lngWidth, lngHeight
        End If
    End If
    hdc = m_hDC
End Property
Private Sub pCreate(ByVal lngWidth As Long, _
                    ByVal lngHeight As Long)
Dim lhDCC As Long
    pDestroy
    lhDCC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    If Not (lhDCC = 0) Then
        m_hDC = CreateCompatibleDC(lhDCC)
        If Not (m_hDC = 0) Then
            m_hBmp = CreateCompatibleBitmap(lhDCC, lngWidth, lngHeight)
            If Not (m_hBmp = 0) Then
                m_hBmpOld = SelectObject(m_hDC, m_hBmp)
                If Not (m_hBmpOld = 0) Then
                    m_lWidth = lngWidth
                    m_lHeight = lngHeight
                    DeleteDC lhDCC
                    Exit Sub
                End If
            End If
        End If
        DeleteDC lhDCC
        pDestroy
    End If
End Sub
Private Sub pDestroy()
    If Not m_hBmpOld = 0 Then
        SelectObject m_hDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not m_hBmp = 0 Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If Not m_hDC = 0 Then
        DeleteDC m_hDC
        m_hDC = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
End Sub
''
''Public Sub Draw(ByVal lngHdc As Long, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0)
''
''
''If WidthSrc <= 0 Then
''WidthSrc = m_lWidth
''End If
''If HeightSrc <= 0 Then
''HeightSrc = m_lHeight
''End If
''BitBlt lngHdc, xDst, yDst, WidthSrc, HeightSrc, m_hDC, xSrc, ySrc, vbSrcCopy
''End Sub
''


