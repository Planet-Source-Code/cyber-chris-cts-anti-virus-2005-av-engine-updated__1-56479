VERSION 5.00
Begin VB.UserControl Hyperlink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "Hyperlink.ctx":0000
   MousePointer    =   99  'Benutzerdefiniert
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3990
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4410
      Top             =   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   225
      Left            =   0
      Shape           =   4  'Gerundetes Rechteck
      Top             =   0
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      MouseIcon       =   "Hyperlink.ctx":08CA
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   0
      Top             =   0
      Width           =   990
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Option Explicit
Private Type SepRGB
    Red                      As Long
    Green                    As Long
    Blue                     As Long
End Type
Private ForeIdle         As Long
Private ForeMouse        As Long
Private ShowRec          As Boolean
Private FadeOut          As Boolean
Private FadeOut2         As Boolean
Private SC               As Boolean
Private FC               As Boolean
Private PendingCaption   As String
Public Event Click()
Public Event Change()
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function SetCapture Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "USER32" (ByVal hWndLock As Long) As Integer
Private Property Get Alignment() As AlignmentConstants
    Alignment = Label1.Alignment
End Property
Private Property Let Alignment(NewAlignment As AlignmentConstants)
    Label1.Alignment = NewAlignment
    If NewAlignment = vbRightJustify Then
        Shape1.Left = ScaleWidth - Shape1.Width
        Label1.Left = ScaleWidth - Label1.Width
    ElseIf NewAlignment = vbLeftJustify Then
        Shape1.Left = 0
    ElseIf NewAlignment = vbCenter Then
        Shape1.Left = 0
        Shape1.Width = Width
        Label1.AutoSize = False
        Label1.Width = Width
    End If
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = Label1.BackColor
End Property
Public Property Let BackColor(newValue As OLE_COLOR)
    If FC Then
        newValue = vbWhite
    End If
    Label1.BackColor = newValue
    UserControl.BackColor = newValue
End Property
Public Property Get Caption() As String
    If FC Then
        Caption = Label1.Caption
    Else
        Caption = Mid$(Label1.Caption, 1, Len(Label1.Caption))
    End If
End Property
Public Property Let Caption(ByVal newValue As String)
'<:-) :WARNING: 'ByVal ' inserted for Parameter 'newValue As String'
    If Label1.Caption = newValue Then
        Exit Property
    End If
    If FC Then
        PendingCaption = newValue
        Timer2.Enabled = True
        FadeOut2 = True
        Exit Property
    End If
    With Label1
        .Caption = newValue
        Shape1.Width = .Width + TextWidth(" ")
        If .Alignment = vbCenter Then
            .AutoSize = False
            Shape1.Width = Width
        End If
    End With 'Label1
    If Label1.Alignment = vbRightJustify Then
        Shape1.Left = ScaleWidth - Shape1.Width
    ElseIf Label1.Alignment = vbCenter Then
        Shape1.Left = 0
    End If
End Property
Private Sub DoMouseActions(ByVal MouseIn As Boolean)
'<:-) :WARNING: 'ByVal ' inserted for Parameter 'MouseIn As Boolean'
    If FC Then
        Exit Sub
    End If
    If MouseIn Then
        If ShowRec Then
            If Shape1.Visible = False Then
                Shape1.Visible = True
            End If
        End If
'Label1.ForeColor = ForeMouse
        FadeOut = False
        Timer1.Enabled = True
        If SC Then
            Image1.Visible = False
            Image2.Visible = True
        End If
    Else
        Shape1.Visible = False
'Label1.ForeColor = ForeIdle
        FadeOut = True
        Timer1.Enabled = True
        If SC Then
            Image1.Visible = True
            Image2.Visible = False
        End If
    End If
End Sub
Public Property Get Enabled() As Boolean
    Enabled = Label1.Enabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
'<:-) :WARNING: 'ByVal ' inserted for Parameter 'newValue As Boolean'
    Label1.Enabled = newValue
End Property
Private Property Get FadeChange() As Boolean
    FadeChange = FC
End Property
Private Property Let FadeChange(newValue As Boolean)
    If newValue Then
        ShowCarrot = False
        ForeColorIdle = 0
        BackColor = vbWhite
    End If
    FC = newValue
End Property
Public Property Get Font() As IFontDisp
    Set Font = Label1.Font
End Property
Public Property Set Font(newValue As IFontDisp)
    Set Label1.Font = newValue
    Set UserControl.Font = newValue
End Property
Private Property Get ForeColorIdle() As OLE_COLOR
    ForeColorIdle = ForeIdle
End Property
Private Property Let ForeColorIdle(NewColor As OLE_COLOR)
    If FC Then
        NewColor = 0
    End If
    Label1.ForeColor = NewColor
    ForeIdle = NewColor
End Property
Private Property Get ForeColorMouse() As OLE_COLOR
    ForeColorMouse = ForeMouse
End Property
Private Property Let ForeColorMouse(newValue As OLE_COLOR)
    ForeMouse = newValue
End Property
Private Function GetRGB(ByVal LongValue As Long) As SepRGB
    LongValue = Abs(LongValue)
    With GetRGB
        .Red = LongValue And 255
        .Green = (LongValue \ 256) And 255
        .Blue = (LongValue \ 65536) And 255
    End With 'GetRGB
End Function
Private Sub Label1_Change()
    RaiseEvent Change
End Sub
Private Sub Label1_Click()
    RaiseEvent Click
End Sub
Private Sub Label1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub
Private Property Get ShowCarrot() As Boolean
    ShowCarrot = SC
End Property
Public Property Let ShowCarrot(ByVal blnValue As Boolean)
'<:-) :WARNING: 'ByVal ' inserted for Parameter 'blnValue As Boolean'
'<:-) :WARNING: Poorly named Parameters 'Value' renamed to 'blnValue'
    SC = blnValue
    If blnValue = False Then
        Label1.Left = 0
    Else
        Image1.Visible = True
        Label1.Left = Image2.Width
        ShowRectangle = False
    End If
End Property
Private Property Get ShowRectangle() As Boolean
    ShowRectangle = ShowRec
End Property
Private Property Let ShowRectangle(ByVal newValue As Boolean)
'<:-) :WARNING: 'ByVal ' inserted for Parameter 'newValue As Boolean'
    ShowRec = newValue
End Property
Private Property Get Speed() As Long
'<:-) :WARNING: Property upgraded to Long
    Speed = Timer1.Interval
End Property
Private Property Let Speed(ByVal newValue As Long)
'<:-) :WARNING: 'ByVal ' inserted for Parameter 'newValue As Long'
'<:-) :WARNING: Integer paramemter 'newValue' upgraded to Long.
    On Error GoTo InvalidProp
    Timer1.Interval = newValue
Exit Property
InvalidProp:
    MsgBox "Invalid property value", vbCritical
End Property
Private Sub Timer1_Timer()
'<:-) :WARNING: Large Code procedure (105 lines of code)
Dim CurRGB       As SepRGB
Dim InBy         As SepRGB
Dim ForeIdleRGB  As SepRGB
Dim ForeMouseRGB As SepRGB
    On Error Resume Next
    ForeMouseRGB = GetRGB(ForeMouse)
    ForeIdleRGB = GetRGB(ForeIdle)
    If FadeOut Then
        CurRGB = GetRGB(Label1.ForeColor)
        With InBy
            .Red = Abs(CurRGB.Red - ForeIdleRGB.Red)
            .Green = Abs(CurRGB.Green - ForeIdleRGB.Green)
            .Blue = Abs(CurRGB.Blue - ForeIdleRGB.Blue)
            .Red = .Red / 7
            .Green = .Green / 7
            .Blue = .Blue / 7
            If .Red = 0 Then
                If .Green = 0 Then
                    If .Blue = 0 Then
                        Timer1.Enabled = False
                        LockWindowUpdate UserControl.hwnd
                        Label1.ForeColor = ForeIdle
                        LockWindowUpdate 0
                        Exit Sub
'<:-) :WARNING: Exiting a procedure from within a With Structure can lead to memory leaks
'<:-) It is advised that you re-structure the code around this line.
                    End If
                End If
            End If
        End With
        With CurRGB
            If ForeIdleRGB.Red <> .Red Then
                If ForeIdleRGB.Red < .Red Then
                    .Red = .Red - InBy.Red
                Else
                    .Red = .Red + InBy.Red
                End If
            End If
            If ForeIdleRGB.Green <> .Green Then
                If ForeIdleRGB.Green < .Green Then
                    .Green = .Green - InBy.Green
                Else
                    .Green = .Green + InBy.Green
                End If
            End If
            If ForeIdleRGB.Blue <> .Blue Then
                If ForeIdleRGB.Blue < .Blue Then
                    .Blue = .Blue - InBy.Blue
                Else
                    .Blue = .Blue + InBy.Blue
                End If
            End If
            LockWindowUpdate UserControl.hwnd
            Label1.ForeColor = RGB(.Red, .Green, .Blue)
            LockWindowUpdate 0
        End With
    Else
        CurRGB = GetRGB(Label1.ForeColor)
        With InBy
            .Red = Abs(CurRGB.Red - ForeMouseRGB.Red)
            .Green = Abs(CurRGB.Green - ForeMouseRGB.Green)
            .Blue = Abs(CurRGB.Blue - ForeMouseRGB.Blue)
            .Red = .Red / 4
            .Green = .Green / 4
            .Blue = .Blue / 4
            If .Red = 0 Then
                If .Green = 0 Then
                    If .Blue = 0 Then
                        Timer1.Enabled = False
                        LockWindowUpdate UserControl.hwnd
                        Label1.ForeColor = ForeMouse
                        LockWindowUpdate 0
                        Exit Sub
'<:-) :WARNING: Exiting a procedure from within a With Structure can lead to memory leaks
'<:-) It is advised that you re-structure the code around this line.
                    End If
                End If
            End If
        End With
        With CurRGB
            If ForeMouseRGB.Red <> .Red Then
                If ForeMouseRGB.Red < .Red Then
                    .Red = .Red - InBy.Red
                Else
                    .Red = .Red + InBy.Red
                End If
            End If
            If ForeMouseRGB.Green <> .Green Then
                If ForeMouseRGB.Green < .Green Then
                    .Green = .Green - InBy.Green
                Else
                    .Green = .Green + InBy.Green
                End If
            End If
            If ForeMouseRGB.Blue <> .Blue Then
                If ForeMouseRGB.Blue < .Blue Then
                    .Blue = .Blue - InBy.Blue
                Else
                    .Blue = .Blue + InBy.Blue
                End If
            End If
            LockWindowUpdate UserControl.hwnd
            Label1.ForeColor = RGB(.Red, .Green, .Blue)
            LockWindowUpdate 0
        End With
    End If
    On Error GoTo 0
End Sub
Private Sub Timer2_Timer()
Dim SepRGB As SepRGB
    On Error Resume Next
    If FadeOut2 Then
        SepRGB = GetRGB(Label1.ForeColor)
        With SepRGB
            .Blue = .Blue + 10
            .Red = .Red + 20
            .Green = .Green + 30
            Label1.ForeColor = RGB(.Red, .Green, .Blue)
            If err Or Label1.ForeColor = vbWhite Then
                Label1.ForeColor = vbWhite
                FadeOut2 = False
                Label1.Caption = PendingCaption
            End If
        End With
    Else
        SepRGB = GetRGB(Label1.ForeColor)
        With SepRGB
            .Blue = .Blue - 10
            .Red = .Red - 20
            .Green = .Green - 30
            Label1.ForeColor = RGB(.Red, .Green, .Blue)
            If err Or Label1.ForeColor = 0 Then
                Label1.ForeColor = 0
                Timer2.Enabled = False
            End If
        End With
    End If
    On Error GoTo 0
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_InitProperties()
Dim TransparentBack '<:-) Missing Dim Auto-inserted(Unable to Auto-Type)
    ForeColorIdle = 0
    ForeColorMouse = &H80FF&
    BackColor = vbButtonFace
    Caption = Ambient.DisplayName
    Set Font = Parent.Font
    Enabled = True
    Alignment = vbLeftJustify
    ShowRectangle = False
    TransparentBack = False
    Speed = 50
    ShowCarrot = False
    FadeChange = False
End Sub
Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
Static LastOne As Long
'<:-) :SUGGESTION: Static is very memory hungry; try using a Private Module level variable instead
    LastOne = SetCapture(hwnd)
    If LastOne = hwnd Then
        LastOne = 0
    End If
    If X < 0 Or X > UserControl.Width Or Y < 0 Or Y > UserControl.Height Then
        DoMouseActions False
        If LastOne = 0 Then
            ReleaseCapture
        Else
            SetCapture LastOne
        End If
    Else
        DoMouseActions True
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    RaiseEvent Click
    UserControl_MouseMove 0, 0, 1, 1
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ForeColorIdle = .ReadProperty("ForeColorIdle", 0)
        ForeColorMouse = .ReadProperty("ForeColorMouse", &H80FF&)
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Set Font = .ReadProperty("Font", Parent.Font)
        Enabled = .ReadProperty("Enabled", True)
        Alignment = .ReadProperty("Alignment", vbLeftJustify)
        ShowRectangle = .ReadProperty("ShowRec", False)
        Speed = .ReadProperty("Speed", 50)
        ShowCarrot = .ReadProperty("ShowCarrot", False)
        FadeChange = .ReadProperty("FC", False)
    End With 'PropBag
End Sub
Private Sub UserControl_Resize()
    If Alignment = vbCenter Then
        Label1.AutoSize = False
        Shape1.Width = Width
        Label1.Width = Width
    ElseIf Alignment = vbRightJustify Then
        Label1.Left = ScaleWidth - Label1.Width
    Else
        Label1.AutoSize = True
        Shape1.Width = Label1.Width + TextWidth(" ")
    End If
    Label1.Height = Height
    If Label1.Alignment = vbCenter Then
        Shape1.Width = Width
    End If
    Shape1.Height = Height
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "ForeColorIdle", ForeIdle, 0
        .WriteProperty "ForeColorMouse", ForeMouse, &H80FF&
        .WriteProperty "BackColor", Label1.BackColor, vbButtonFace
    End With 'PropBag
    If FC Then
        PropBag.WriteProperty "Caption", Label1.Caption, " " & Ambient.DisplayName
    Else
        PropBag.WriteProperty "Caption", Mid$(Label1.Caption, 1, Len(Label1.Caption)), Ambient.DisplayName
    End If
    With PropBag
        .WriteProperty "Font", Label1.Font, Parent.Font
        .WriteProperty "Enabled", Label1.Enabled, True
        .WriteProperty "Alignment", Label1.Alignment, vbLeftJustify
        .WriteProperty "ShowRec", ShowRec, False
        .WriteProperty "Speed", Timer1.Interval, 50
        .WriteProperty "ShowCarrot", SC, False
        .WriteProperty "FC", FC, False
    End With 'PropBag
    On Error GoTo 0
End Sub
''
''Private Sub DoMouseOutFade(ByVal ReturnImed As Boolean)
''
'''<:-) :WARNING: Unused Sub 'DoMouseOutFade'
'''<:-) :WARNING: 'ByVal ' inserted for Parameter 'ReturnImed As Boolean'
''Label1.ForeColor = ForeMouse
''FadeOut = True
''Timer1.Enabled = True
''Timer1_Timer
''If ReturnImed = False Then
''Do Until Timer1.Enabled = False
''DoEvents
''Loop
''End If
''End Sub
''
''
''Private Function TextBoxHWND() As Long
''
'''<:-) :WARNING: Unused Function 'TextBoxHWND'
''TextBoxHWND = Text1.hwnd
''End Function
''
':)Code Fixer V2.5.3 (02.10.2004 06:24:44) 20 + 565 = 585 Lines Thanks Ulli for inspiration and lots of code.


