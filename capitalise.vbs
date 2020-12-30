Option Explicit
Const module_name = "QuickStuff"
Const module_ver = "1.00"

Sub CapitaliseAll
  Dim MyDoc
  Dim InputString, OutputString, PrevChar, ThisChar 'Strings
  Dim i, j 'Counters

  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'Process each line
  For i = 1 To MyDoc.linesCount
    'Home of current line
    MyDoc.caretY(i)
    'Get whole line 
    InputString = MyDoc.lineText
    'Initialise OutputString
    OutputString = vbNullString
    For j = 1 To Len(InputString)
      'Get next character
      ThisChar = Mid(InputString, j, 1)
      'Get previous character
      If j>1 Then
        'Not beginning of line: read previous character
        PrevChar = Mid(InputString, j - 1, 1)
      Else
        'Beginning of line: assume needs capitalising
        PrevChar= " "
      End If
      'Following a space?
      If PrevChar <> " " Then
        'No: copy as is
        OutputString = OutputString & ThisChar
      Else
        'Yes: capitalise
        OutputString = OutputString & UCase(ThisChar)
      End If
    Next
    'Write capitalised conversion in current line
    MyDoc.lineText(OutputString)
  Next
End Sub

Sub CapitaliseLine
  Dim MyDoc
  Dim InputString, OutputString, PrevChar, ThisChar 'Strings
  Dim i 'Counter

  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'read current line
  InputString = MyDoc.lineText
  'Initialise OutputString
  OutputString = vbNullString
  For i = 1 To Len(InputString)
    'Get next character
    ThisChar = Mid(InputString, i, 1)
    'Get previous character
    If i>1 Then
      'Not beginning of line: read previous character
      PrevChar = Mid(InputString, i - 1, 1)
    Else
      'Beginning of line: assume needs capitalising
      PrevChar= " "
    End If
    'Following a space?
    If PrevChar <> " " Then
      'No: copy as is
      OutputString = OutputString & ThisChar
    Else
      'Yes: capitalise
      OutputString = OutputString & UCase(ThisChar)
    End If
  Next
  'Write capitalised conversion in current line
  MyDoc.lineText(OutputString)
End Sub

Sub Init
  addMenuItem "Capitalise (line)", "", "CapitaliseLine"
  addMenuItem "Capitalise (all)", "", "CapitaliseAll"  
End Sub