Option Explicit
Const module_name = "MSX"
Const module_ver = "1.00"

' Convert:
' 0,47446,20,0,power
' To:
' 1,B956,14,power
'The CHT file format is the following :
'Enable, Address, Value, Comment

Sub bluemsxTOfmsx
  Dim  MyDoc
  Dim InputString, OutputString, NextChunk 'Strings
  Dim i, j 'Counter

  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'Process each line
  For i = 1 To MyDoc.linesCount
    'Home of current line
    MyDoc.caretY(i)
    'Comment? Delete line
    If Left(MyDoc.lineText, 2) = "0," Then
      'Get whole line 
      InputString = MyDoc.lineText
      'Get address
      NextChunk = Mid(InputString, 3, 5)
      'Initialise output and add Address in hex
      OutputString = "1," & Hex(CLng(NextChunk)) & ","
      'Get value
      j = InStr(9, InputString, ",")
      NextChunk = Mid(InputString, 9, j - 9)
      'Add Value to output
      OutputString = OutputString & Hex(CLng(NextChunk)) & ","
      'Get comment
      j = InStr(j + 1, InputString, ",")
      NextChunk = Mid(InputString, j + 1, 999)
      'Add comment to output
      OutputString = OutputString & NextChunk
      'Write fMSX conversion in current line
      MyDoc.lineText(OutputString)
    Else
      'Not a poke declaration: delete line
      MyDoc.lineText(vbNullString)
    End If
  Next
  'Remove empty lines
  runPSPadAction "aRemoveBlankLines"
End Sub

Sub Init
  addMenuItem "BlueMSX->fMSX", module_name, "bluemsxTOfmsx"
End Sub