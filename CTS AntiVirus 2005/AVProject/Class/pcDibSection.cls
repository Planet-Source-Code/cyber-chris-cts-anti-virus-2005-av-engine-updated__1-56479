VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcDibSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type SAFEARRAYBOUND
    cElements                                       As Long
    lLbound                                         As Long
End Type
Private Type SAFEARRAY2D
    cDims                                           As Long
    fFeatures                                       As Long
    cbElements                                      As Long
    cLocks                                          As Long
    pvData                                          As Long
    Bounds(0 To 1)                                  As SAFEARRAYBOUND
End Type
Private Type RGBQUAD
    rgbBlue                                         As Byte
    rgbGreen                                        As Byte
    rgbRed                                          As Byte
    rgbReserved                                     As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize                                          As Long
    biWidth                                         As Long
    biHeight                                        As Long
    biPlanes                                        As Long
    biBitCount                                      As Long
    biCompression                                   As Long
    biSizeImage                                     As Long
    biXPelsPerMeter                                 As Long
    biYPelsPerMeter                                 As Long
    biClrUsed                                       As Long
    biClrImportant                                  As Long
End Type
Private Type BITMAPINFO
    bmiHeader                                       As BITMAPINFOHEADER
    bmiColors                                       As RGBQUAD
End Type
Private Type BLENDFUNCTION
    BlendOp                                         As Byte
    BlendFlags                                      As Byte
    SourceConstantAlpha                             As Byte
    AlphaFormat                                     As Byte
End Type
Private Const AC_SRC_OVER                       As Long = &H0
Private Const AC_SRC_ALPHA                      As Long = &H1
Private Const BI_RGB                            As Long = 0
Private Const DIB_RGB_COLORS                    As Long = 0 '  color table in RGBs
Private Type BITMAP
    bmType                                          As Long
    bmWidth                                         As Long
    bmHeight                                        As Long
    bmWidthBytes                                    As Long
    bmPlanes                                        As Long    '  color table in RGBs
    bmBitsPixel                                     As Long    '  color table in RGBs
    bmBits                                          As Long
End Type
Private Type BITMAPFILEHEADER
    bfType                                          As Long
'- type  ="BM" i.e &H4D42 - 2 '  color table in RGBs
    bfSize                                          As Long    ' - size in bytes of file - 6
    bfReserved1                                     As Long
' - reserved, must be 0 - 8 '  color table in RGBs
    bfReserved2                                     As Long
' - reserved, must be 0 - 10 '  color table in RGBs
    bfOffBits                                       As Long
' offset from this structure to the bitmap bits - 14
End Type
''Private Const GMEM_FIXED                           As Long = &H0
''Private Const GMEM_ZEROINIT                        As Long = &H40
Private Const FORMAT_MESSAGE_FROM_SYSTEM        As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS     As Long = &H200
Private m_hDIb                                  As Long
Private m_hBmpOld                               As Long
Private m_hDC                                   As Long
Private m_hDD                                   As Long
Private m_lPtr                                  As Long
Private m_tBI                                   As BITMAPINFO
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "USER32" () As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
                                                       pBitmapInfo As BITMAPINFO, _
                                                       ByVal un As Long, _
                                                       lplpVoid As Long, _
                                                       ByVal handle As Long, _
                                                       ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "MSIMG32.dll" (ByVal hdcDest As Long, _
                                                       ByVal nXOriginDest As Long, _
                                                       ByVal nYOriginDest As Long, _
                                                       ByVal nWidthDest As Long, _
                                                       ByVal nHeightDest As Long, _
                                                       ByVal hdcSrc As Long, _
                                                       ByVal nXOriginSrc As Long, _
                                                       ByVal nYOriginSrc As Long, _
                                                       ByVal nWidthSrc As Long, _
                                                       ByVal nHeightSrc As Long, _
                                                       ByVal lBlendFunction As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long
Private Declare Function FormatMessage Lib "KERNEL32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                                                              lpSource As Any, _
                                                                              ByVal dwMessageId As Long, _
                                                                              ByVal dwLanguageId As Long, _
                                                                              ByVal lpBuffer As String, _
                                                                              ByVal nSize As Long, _
                                                                              Arguments As Long) As Long
Private Declare Function DrawDibOpen Lib "msvfw32.dll" () As Long
Private Declare Function DrawDibClose Lib "msvfw32.dll" (ByVal hDD As Long) As Long
Private Declare Function DrawDibDraw Lib "msvfw32.dll" (ByVal hDD As Long, _
                                                        ByVal hdc As Long, _
                                                        ByVal xDst As Long, _
                                                        ByVal yDst As Long, _
                                                        ByVal dxDst As Long, _
                                                        ByVal dyDst As Long, _
                                                        lpBI As Any, _
                                                        lpBits As Any, _
                                                        ByVal xSrc As Long, _
                                                        ByVal ySrc As Long, _
                                                        ByVal dxSrc As Long, _
                                                        ByVal dySrc As Long, _
                                                        ByVal wFlags As Long) As Long
Private Function ApiError(ByVal e As Long) As String
Dim s As String
Dim c As Long
    s = String$(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, e, 0&, s, Len(s), ByVal 0)
    If c Then
        ApiError = Left$(s, c)
    End If
End Function
Public Property Get BytesPerScanLine() As Long
    BytesPerScanLine = m_tBI.bmiHeader.biWidth * 4
End Property
Private Sub Class_Terminate()
    ClearUp
End Sub
Public Sub Clear(Optional ByVal lAlpha As Byte = 0)
Dim bDib() As Byte
Dim X      As Long
Dim Y      As Long
Dim tSA    As SAFEARRAY2D
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    For X = 0 To BytesPerScanLine - 1 Step 4
        For Y = 0 To m_tBI.bmiHeader.biHeight - 1
            bDib(X + 3, Y) = 0 ' transparent
            bDib(X, Y) = 0
            bDib(X + 1, Y) = 0
            bDib(X + 2, Y) = 0
        Next Y
    Next X
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
End Sub
Public Sub ClearUp()
    If m_hDC <> 0 Then
        If m_hDIb <> 0 Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0
    m_hDIb = 0
    m_hBmpOld = 0
    m_lPtr = 0
    If Not (m_hDD = 0) Then
        DrawDibClose m_hDD
        m_hDD = 0
    End If
End Sub
Public Function Create(ByVal lWidth As Long, _
                       ByVal lHeight As Long) As Boolean
Dim bDrawDib As Boolean
    bDrawDib = UseDrawDib()
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If m_hDC <> 0 Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            UseDrawDib = bDrawDib
            Create = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Function CreateDIB(ByVal lHDC As Long, _
                          ByVal lWidth As Long, _
                          ByVal lHeight As Long, _
                          ByRef hDib As Long) As Boolean
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection(lHDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
    CreateDIB = (hDib <> 0)
End Function
Public Sub CreateFromHBitmap(ByVal hBmp As Long)
Dim lHDC        As Long
Dim lhDCDesktop As Long
Dim lhBmpOld    As Long
Dim tBMP        As BITMAP
    GetObjectAPI hBmp, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhDCDesktop = GetDC(GetDesktopWindow())
        If lhDCDesktop <> 0 Then
            lHDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If lHDC <> 0 Then
                lhBmpOld = SelectObject(lHDC, hBmp)
                LoadPictureBlt lHDC
                SelectObject lHDC, lhBmpOld
                DeleteObject lHDC
            End If
        End If
    End If
End Sub
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Property Get hdc() As Long
    hdc = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDIb
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property
Private Sub HLSToRGB(ByVal H As Single, _
                     ByVal s As Single, _
                     ByVal l As Single, _
                     r As Long, _
                     g As Long, _
                     b As Long)
Dim rR  As Single
Dim rG  As Single
Dim rB  As Single
Dim Min As Single
Dim Max As Single
    If s = 0 Then
' Achromatic case:
        rR = l
        rG = l
        rB = l
    Else
' Chromatic case:
' delta = Max-Min
        If l <= 0.5 Then
's = (Max - Min) / (Max + Min)
' Get Min value:
            Min = l * (1 - s)
        Else
's = (Max - Min) / (2 - Max - Min)
' Get Min value:
            Min = l - s * (1 - l)
        End If
' Get the Max value:
        Max = 2 * l - Min
' Now depending on sector we can evaluate the h,l,s:
        If H < 1 Then
            rR = Max
            If H < 0 Then
                rG = Min
                rB = rG - H * (Max - Min)
            Else 'NOT (H...
                rB = Min
                rG = H * (Max - Min) + rB
            End If
        ElseIf (H < 3) Then 'NOT (H...
            rG = Max
            If H < 2 Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
            Else 'NOT (H...
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If H < 4 Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
            Else 'NOT (H...
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
    End If
    r = rR * 255
    g = rG * 255
    b = rB * 255
End Sub
Public Sub LoadPictureBlt(ByVal lHDC As Long, _
                          Optional ByVal lSrcLeft As Long = 0, _
                          Optional ByVal lSrcTop As Long = 0, _
                          Optional ByVal lSrcWidth As Long = -1, _
                          Optional ByVal lSrcHeight As Long = -1, _
                          Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    If lSrcWidth < 0 Then
        lSrcWidth = m_tBI.bmiHeader.biWidth
    End If
    If lSrcHeight < 0 Then
        lSrcHeight = m_tBI.bmiHeader.biHeight
    End If
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
End Sub
Private Function Maximum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single
    If rR > rG Then
        If rR > rB Then
            Maximum = rR
        Else 'NOT (RR...
            Maximum = rB
        End If
    Else 'NOT (RR...
        If rB > rG Then
            Maximum = rB
        Else 'NOT (RB...
            Maximum = rG
        End If
    End If
End Function
Private Function Minimum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single
    If rR < rG Then
        If rR < rB Then
            Minimum = rR
        Else 'NOT (RR...
            Minimum = rB
        End If
    Else 'NOT (RR...
        If rB < rG Then
            Minimum = rB
        Else 'NOT (RB...
            Minimum = rG
        End If
    End If
End Function
Private Sub RGBToHLS(ByVal r As Long, _
                     ByVal g As Long, _
                     ByVal b As Long, _
                     H As Single, _
                     s As Single, _
                     l As Single)
Dim Max   As Single
Dim Min   As Single
Dim delta As Single
Dim rR    As Single
Dim rG    As Single
Dim rB    As Single
    rR = r / 255
    rG = g / 255
    rB = b / 255
'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2    '{This is the lightness}
'{Next calculate saturation}
    If Max = Min Then
'begin {Acrhomatic case}
        s = 0
        H = 0
'end {Acrhomatic case}
    Else
'begin {Chromatic case}
'{First calculate the saturation.}
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
'{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            H = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            H = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
    End If
End Sub
Public Property Get UseDrawDib() As Boolean
    UseDrawDib = Not (m_hDD = 0)
End Property
Public Property Let UseDrawDib(ByVal bState As Boolean)
    If bState Then
        If m_hDD = 0 Then
            m_hDD = DrawDibOpen()
        End If
    Else
        If Not (m_hDD = 0) Then
            DrawDibClose m_hDD
        End If
    End If
End Property
Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
''
''Public Sub AlphaPaintPicture(ByVal lHDC As Long, Optional ByVal lDestLeft As Long = 0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, Optional ByVal lDestHeight As Long = -1, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lConstantAlpha As Byte = 255)
''
''
''Dim lBlend As Long
''Dim bf     As BLENDFUNCTION
''Dim lR     As Long
''If lDestWidth < 0 Then
''lDestWidth = m_tBI.bmiHeader.biWidth
''End If
''If lDestHeight < 0 Then
''lDestHeight = m_tBI.bmiHeader.biHeight
''End If
''With bf
''.BlendOp = AC_SRC_OVER
''.BlendFlags = 0
''.SourceConstantAlpha = lConstantAlpha
''.AlphaFormat = AC_SRC_ALPHA
''End With 'bf
''CopyMemory lBlend, bf, 4
''lR = AlphaBlend(lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, lDestWidth, lDestHeight, lBlend)
''If lR = 0 Then
''Debug.Print ApiError(err.LastDllError)
''End If
''End Sub
''
''
''Public Sub CodeAlphaBlend(cDst As pcDibSection, cSrc As pcDibSection, Optional ByVal lDestLeft As Long = 0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, Optional ByVal lDestHeight As Long = -1, Optional ByVal SourceConstantAlpha As Long)
''
''
''
''Dim tSADst    As SAFEARRAY2D
''Dim bDibDst() As Byte
''Dim X         As Long
''
''
''
''
''Dim Y         As Long
''Dim xDest     As Long
''Dim yDest     As Long
''Dim xEnd      As Long
''Dim yEnd      As Long
''Dim tSASrc    As SAFEARRAY2D
''Dim bDibSrc() As Byte
''' Get the bits in the from DIB section:
''With tSASrc
''.cbElements = 1
''.cDims = 2
''.Bounds(0).lLbound = 0
''.Bounds(0).cElements = cSrc.Height
''.Bounds(1).lLbound = 0
''.Bounds(1).cElements = cSrc.BytesPerScanLine()
''.pvData = cSrc.DIBSectionBitsPtr
''End With
''CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4
''' Get the bits in the from DIB section:
''With tSADst
''.cbElements = 1
''.cDims = 2
''.Bounds(0).lLbound = 0
''.Bounds(0).cElements = cDst.Height
''.Bounds(1).lLbound = 0
''.Bounds(1).cElements = cDst.BytesPerScanLine()
''.pvData = cDst.DIBSectionBitsPtr
''End With
''CopyMemory ByVal VarPtrArray(bDibDst()), VarPtr(tSADst), 4
''If lDestWidth = -1 Then
''lDestWidth = cDst.Width
''If lDestWidth > cSrc.Width Then
''lDestWidth = cSrc.Width
''xEnd = cSrc.BytesPerScanLine
''Else 'NOT (LDESTWIDTH...
''xEnd = cDst.BytesPerScanLine
''End If
''End If
''If lDestHeight = -1 Then
''lDestHeight = cDst.Height
''If lDestHeight > cSrc.Height Then
''lDestHeight = cSrc.Height
''yEnd = cSrc.Height
''Else 'NOT (LDESTHEIGHT...
''yEnd = cDst.Height
''End If
''End If
''For X = 0 To xEnd - 4 Step 4
''xDest = X + lDestLeft * 4
''For Y = 0 To yEnd - 1
''yDest = Y + lDestTop
''bDibDst(xDest, yDest) = bDibSrc(X, Y) * SourceConstantAlpha / 255 + bDibDst(xDest, yDest) * (1 - SourceConstantAlpha / 255)
''bDibDst(xDest + 1, yDest) = bDibSrc(X + 1, Y) * SourceConstantAlpha / 255 + bDibDst(xDest + 1, yDest) * (1 - SourceConstantAlpha / 255)
''bDibDst(xDest + 2, yDest) = bDibSrc(X + 2, Y) * SourceConstantAlpha / 255 + bDibDst(xDest + 2, yDest) * (1 - SourceConstantAlpha / 255)
''Next Y
''Next X
''CopyMemory ByVal VarPtrArray(bDibDst), 0&, 4
''CopyMemory ByVal VarPtrArray(bDibSrc), 0&, 4
''End Sub
''
''
''Public Sub CreateFromPicture(ByRef picThis As StdPicture)
''
''
''CreateFromHBitmap picThis.handle
''End Sub
''
''
''Public Sub Lighten(ByVal lAmount As Long)
''
''
''Dim bDib() As Byte
''Dim X      As Long
''
''
''
''
''Dim Y      As Long
''Dim xMax   As Long
''Dim yMax   As Long
''Dim lB     As Long
''Dim lG     As Long
''Dim lR     As Long
''Dim H      As Single
''Dim s      As Single
''Dim l      As Single
''Dim tSA    As SAFEARRAY2D
''' have the local matrix point to bitmap pixels
''With tSA
''.cbElements = 1
''.cDims = 2
''.Bounds(0).lLbound = 0
''.Bounds(0).cElements = Height
''.Bounds(1).lLbound = 0
''.Bounds(1).cElements = BytesPerScanLine
''.pvData = DIBSectionBitsPtr
''End With
''CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
''yMax = Height - 1
''xMax = Width - 1
''For X = 0 To (xMax * 4) Step 4
''For Y = 0 To yMax
''RGBToHLS bDib(X + 2, Y), bDib(X + 1, Y), bDib(X, Y), H, s, l
''l = l * (1 + (lAmount / 100))
''If l > 1 Then
''l = 1
''End If
''HLSToRGB H, s, l, lR, lG, lB
''bDib(X, Y) = lB
''bDib(X + 1, Y) = lG
''bDib(X + 2, Y) = lR
''Next Y
''Next X
''' clear the temporary array descriptor
''' without destroying the local temporary array
''CopyMemory ByVal VarPtrArray(bDib), 0&, 4
''End Sub
''
''
''Public Sub ModifyHLS(ByVal fHue As Double, ByVal fLuminance As Double, ByVal fSaturation As Double)
''
''
''Dim bDib() As Byte
''Dim X      As Long
''
''
''
''
''Dim Y      As Long
''Dim xMax   As Long
''Dim yMax   As Long
''Dim lB     As Long
''Dim lG     As Long
''Dim lR     As Long
''Dim H      As Single
''Dim s      As Single
''Dim l      As Single
''Dim tSA    As SAFEARRAY2D
''' have the local matrix point to bitmap pixels
''With tSA
''.cbElements = 1
''.cDims = 2
''.Bounds(0).lLbound = 0
''.Bounds(0).cElements = Height
''.Bounds(1).lLbound = 0
''.Bounds(1).cElements = BytesPerScanLine
''.pvData = DIBSectionBitsPtr
''End With
''CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
''yMax = Height - 1
''xMax = Width - 1
''For X = 0 To (xMax * 4) Step 4
''For Y = 0 To yMax
''RGBToHLS bDib(X + 2, Y), bDib(X + 1, Y), bDib(X, Y), H, s, l
''H = H * fHue
''l = l * fLuminance
''If l > 1 Then
''l = 1
''End If
''s = s * fSaturation
''If s > 1 Then
''s = 1
''End If
''HLSToRGB H, s, l, lR, lG, lB
''bDib(X, Y) = lB
''bDib(X + 1, Y) = lG
''bDib(X + 2, Y) = lR
''Next Y
''Next X
''' clear the temporary array descriptor
''' without destroying the local temporary array
''CopyMemory ByVal VarPtrArray(bDib), 0&, 4
''End Sub
''
''
''Public Sub PaintPicture(ByVal lHDC As Long, Optional ByVal lDestLeft As Long = 0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, Optional ByVal lDestHeight As Long = -1, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
''
''
''If lDestWidth < 0 Then
''lDestWidth = m_tBI.bmiHeader.biWidth
''End If
''If lDestHeight < 0 Then
''lDestHeight = m_tBI.bmiHeader.biHeight
''End If
''If Not (m_hDD = 0) Then
''' DrawDib method:
''DrawDibDraw m_hDD, lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_tBI, ByVal m_lPtr, lSrcLeft, lSrcTop, lDestWidth, lDestHeight, 0
''Else
''BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
''End If
''End Sub
''
''
''Public Sub RandomiseBits(Optional ByVal bGray As Boolean = False)
''
''
''Dim bDib() As Byte
''Dim X      As Long
''
''
''
''
''Dim Y      As Long
''Dim lC     As Long
''Dim tSA    As SAFEARRAY2D
''' Get the bits in the from DIB section:
''With tSA
''.cbElements = 1
''.cDims = 2
''.Bounds(0).lLbound = 0
''.Bounds(0).cElements = m_tBI.bmiHeader.biHeight
''.Bounds(1).lLbound = 0
''.Bounds(1).cElements = BytesPerScanLine()
''.pvData = m_lPtr
''End With
''CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
''' random:
''Randomize Timer
''If bGray Then
''For Y = 0 To m_tBI.bmiHeader.biHeight - 1
''For X = 0 To BytesPerScanLine - 1 Step 4
''lC = Rnd * 255
''bDib(X + 3, Y) = Rnd * 255 '255 * (y / m_tBI.bmiHeader.biHeight)
''bDib(X, Y) = lC * bDib(X + 3, Y) / 255
''bDib(X + 1, Y) = lC * bDib(X + 3, Y) / 255
''bDib(X + 2, Y) = lC * bDib(X + 3, Y) / 255
''Next X
''Next Y
''Else '(BGRAY) = FALSE/0
''For Y = 0 To m_tBI.bmiHeader.biHeight - 1
''For X = 0 To BytesPerScanLine - 1 Step 4
''bDib(X + 3, Y) = Rnd * 255 '255 * (y / m_tBI.bmiHeader.biHeight)
''bDib(X, Y) = Rnd * 255 * bDib(X + 3, Y) / 255
''bDib(X + 1, Y) = Rnd * 255 * bDib(X + 3, Y) / 255
''bDib(X + 2, Y) = Rnd * 255 * bDib(X + 3, Y) / 255
''Next X
''Next Y
''End If
''' Clear the temporary array descriptor
''' (This does not appear to be necessary, but
''' for safety do it anyway)
''CopyMemory ByVal VarPtrArray(bDib), 0&, 4
''End Sub
''


