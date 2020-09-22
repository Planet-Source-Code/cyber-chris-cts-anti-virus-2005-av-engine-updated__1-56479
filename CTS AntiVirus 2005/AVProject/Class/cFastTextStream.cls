VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Faster Reading of larger files
'Jonathan Orgel, kalonymous, Inc.

Option Explicit
Const VSS_ID = "$Header: /FileIO Time Test/Forms/main.frm 5     1/05/99 6:49p Joni $"
'$NoKeywords: $


Const SkipFromStart = 0
Const SkipFromCurrent = 1
Const SkipFromEnd = 2

'Local variables
Private Buffer As String * 63000        'The buffer to use with get. Declared here to allow bigger buffer
Private FileNumber As Integer           'The filenumber of the opened file
Private LinesArray() As String          'Array with the strings
Private LinesInArray As Long           'Number of strings in the LinesArray
Private NextLine As Long                'NextLine to give out with ReadLine


'local variable(s) to hold property value(s)
Private mvarPath As String              'Name (including path) of the open file)
Private mvarAtEndOfFile As Boolean      'True if  beyond the last character in the file
Private mvarSize As Long                'Size of the file in bytes
Private mvarCurrentPosition As Long     'Current byte position in file

Public Function ReadLine() As Variant
    'Read a line (from the LinesArray buffer).
    'Return the string as a variant or Empty if at EOF
    Dim LineLength As Long
    Dim Line As String
    
    If NextLine > LinesInArray Then
        'No more lines in buffer. Get some more.
        ReadBuffer
    End If
    
    If mvarAtEndOfFile Then
        ReadLine = Empty
        Exit Function
    End If
    
    Line = LinesArray(NextLine)
    LineLength = Len(Line)
    'Line read is length of string + 1 for LF
    mvarCurrentPosition = mvarCurrentPosition + LineLength + 1
    
    'skip CR if present
    If Right$(Line, 1) = vbCr Then
        ReadLine = Left$(Line, LineLength - 1)
    Else
        ReadLine = Line
    End If
    NextLine = NextLine + 1
End Function


Public Sub Skip(Optional NumberOfBytes As Long = 1, Optional StartingPoint As Integer = SkipFromCurrent)
    'Skip 'NumberOfBytes' (may be negative) from current position,
    'start of file or end of file
    Dim NewPosition As Long
   
    Select Case StartingPoint
    Case SkipFromCurrent
        NewPosition = mvarCurrentPosition + NumberOfBytes
    Case SkipFromEnd
        NewPosition = mvarSize - NumberOfBytes
    Case Else 'Includes SkipFromStart
        NewPosition = NumberOfBytes
    End Select
    
    If (NewPosition < 0 Or NewPosition > mvarSize) Then
        'Ignore illegal request
        Exit Sub
    End If
    
    Seek FileNumber, NewPosition + 1
    mvarCurrentPosition = NewPosition
    ReadBuffer
    
End Sub



Public Function OpenTextFileForReading(Path As String) As Boolean
    'Returns True for success and false for an error

    If FileNumber > 0 Then
        'Already have an opened file
        Exit Function
        OpenTextFileForReading = False
    End If
    
    On Error GoTo Failed    'Catch any type of open statement error.
    
    mvarSize = FileLen(Path)
    mvarPath = Path
    
    FileNumber = FreeFile()
    Open Path For Binary As FileNumber
    
    ReadBuffer 'Get first group of lines in to LinesArray buffer.
    mvarCurrentPosition = 0
    OpenTextFileForReading = True
    Exit Function
    
Failed:
    OpenTextFileForReading = False
   
End Function


Public Property Get CurrentPosition() As Long
    CurrentPosition = mvarCurrentPosition
End Property


Public Property Get Size() As Long
    Size = mvarSize
End Property


Public Property Get AtEndOfFile() As Boolean
    AtEndOfFile = mvarAtEndOfFile
End Property


Public Property Get Path() As String
    Path = mvarPath
End Property

Private Sub ReadBuffer()
    'Refill the LinesArray buffer and set the value of
    'mvarAtEndOfFile if at the end of the stream
    If Loc(FileNumber) >= mvarSize Then
        mvarAtEndOfFile = True
    Else
        Get FileNumber, , Buffer
        LinesArray = Split(Buffer, vbLf)
        LinesInArray = UBound(LinesArray)
        NextLine = 0
    End If
End Sub


Private Sub initializeValues()
    'Set values to "undefined"
    FileNumber = -1
    mvarPath = ""
    mvarAtEndOfFile = False     'One could argue about this
    mvarSize = -1
    mvarCurrentPosition = -1
    LinesInArray = 0
    NextLine = -1
End Sub


Private Sub Class_Terminate()
    If FileNumber > 0 Then
            Close FileNumber
    End If
End Sub

    
'The following functions are only provided for debugging purposes.
'Eliminate them by setting the conFTS_DEBUG constant to False

#Const conFTS_DEBUG = True

#If conFTS_DEBUG Then

    Public Function getBufferLine(LineNumber As Integer) As String
        If LineNumber < LBound(LinesArray) Or LineNumber > UBound(LinesArray) Then
            getBufferLine = "Error: LineNumber out of bounds"
        Else
            getBufferLine = LinesArray(LineNumber)
        End If
    End Function
    
    Public Property Get NumberOfBufferLines() As Long
        NumberOfBufferLines = LinesInArray
    End Property
    
    Public Property Get IndexNextBufferLine() As Long
        IndexNextBufferLine = NextLine
    End Property

    Public Property Get RealFilePosition() As Long
        If FileNumber > 0 Then
            RealFilePosition = Loc(FileNumber)
        Else
            RealFilePosition = -1
        End If
    End Property
#End If '#If FTS_DEBUG Then



