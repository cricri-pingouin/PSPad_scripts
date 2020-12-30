Option Explicit
Const module_name = "ClrMAME"
Const module_ver = "1.00"

Const username = "user" 'User name
Const userpwd = "password" 'User password
Const romsdestinationfolder = "D:\Destination\" 'Destination path for downloaded ROMs, must finish with \
Const outputpath = "queue.xml" 'Path and name of output queue file
Const romspath = "xxxxxxxxxxxxx" 'FTP server romspath in FileZilla format goes here
Const samplespath = "xxxxxxxxxxxxx" 'FTP server samplespath in FileZilla format goes here

Sub ClrMAMEtoFilezilla
  Dim MyDoc, GetWhat, InputString, OutputString, i, j, k, IsCHD

  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  GetWhat = MsgBox("What files do you want to add to the download queue?" & vbCrLf & "Yes: ROMs only, No: samples only, Cancel: both", vbYesNoCancel + vbQuestion, module_name)
  'Process each line
  For i = 1 To MyDoc.linesCount
    'Home of current line
    MyDoc.caretY(i)
    'Check if line declares a new missing file
    InputString = MyDoc.lineText
    j = InStr(InputString, "folder: ")
    'Default: assume not a CHD
    IsCHD = False
    'If missing file, check if CHD
    If j Then
      'Start searching at current row
      k = i
      Do
        'Go to next line
        k = k + 1
        MyDoc.caretY(k)
        'If CHD, set flag
        If Left(MyDoc.lineText, 13) = "missing chd: " Then IsCHD = True
      Loop until Len(MyDoc.lineText) < 3 Or IsCHD 'Stop when found or short lines less than 3 chars, e.g. vbCrLf
    End If
    'Check if missing file and not a CHD
    If j And Not IsCHD Then
      'Extract file name
      InputString = Mid(InputString, j + 8)
      InputString = Mid(InputString, 1, InStr(InputString, " ") - 1)
      'Generate queue item
      OutputString = "<File><LocalFile>" & romsdestinationfolder & InputString & _
      ".zip</LocalFile><RemoteFile>" & InputString & ".zip</RemoteFile><RemotePath>"
      'Get next line to check file type
      MyDoc.caretY(i + 1)
      InputString = MyDoc.lineText
      'Back to current line
      MyDoc.caretY(i)
      'Write either ROMs or samples path accordingly. CHDs not supported
      If Left(InputString, 16) = "missing sample: " Then
        If GetWhat <> vbYes Then
          OutputString = OutputString & samplespath & "</RemotePath><Download>1</Download><Priority>2</Priority><TransferMode>1</TransferMode></File>"
        Else
          OutputString = vbNullString
        End If
      Else
        If GetWhat <> vbNo Then
          OutputString = OutputString & romspath & "</RemotePath><Download>1</Download><Priority>2</Priority><TransferMode>1</TransferMode></File>"
        Else
          OutputString = vbNullString
        End If
      End If
      'Write FileZilla XML queue item in current line
      MyDoc.lineText(OutputString)
    Else
      'Not a missing file declaration: delete line
      MyDoc.lineText(vbNullString)
    End If
  Next
  'Clean up non-XML lines
  For i = 1 To MyDoc.linesCount
    MyDoc.caretY(i)
    'If line does not start with < delete line
    If Left(MyDoc.lineText, 1) <> "<" Then MyDoc.lineText(vbNullString)
  Next
  'Remove empty lines
  runPSPadAction "aRemoveBlankLines"
  'Write XML header tags
  MyDoc.text("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes"" ?><FileZilla3><Queue><Server><Host>ftp.server_name.com</Host><Port>123</Port><Protocol>3</Protocol><Type>0</Type><Logontype>1</Logontype><User>" _
  & username & "</User><Pass>" & userpwd & _
  "</Pass><TimezoneOffset>0</TimezoneOffset><PasvMode>MODE_DEFAULT</PasvMode><MaximumMultipleConnections>1</MaximumMultipleConnections><EncodingType>Auto</EncodingType>" _
  & vbCrLf & MyDoc.text & vbCrLf & "</Server></Queue></FileZilla3>")
  'Save as FileZilla XML queue file
  MyDoc.saveFileAs(outputpath)
End Sub

Sub KeepMatches
  Dim i, j
  'Script requires at least 2 opened documents
  If editorsCount < 2 Then
    MsgBox "You need at least 2 opened documents!" & vbCrLf & "The document to process (which must be active) and a document containing the keywords."
    Exit sub
  End If
  'Set document to process
  Dim ProcessDoc
  Set ProcessDoc = newEditor
  ProcessDoc.assignActiveEditor
  'Set source document
  Dim KeywordsDoc
  Set KeywordsDoc = newEditor
  For i = 0 To editorsCount - 1
    KeywordsDoc.assignEditorByIndex(i)
    If KeywordsDoc.fileName <> ProcessDoc.fileName Then Exit For
  Next
  'Populate keywords array
  Dim Prefix, Postfix
  Prefix = InputBox("This will keep lines in " & ProcessDoc.fileName & " that contain any keyword from " & KeywordsDoc.fileName & vbCrLf & "Enter prefix/postfix pattern in the form prefix*postfix:", , ">*.zip")
  i = InStr(Prefix, "*")
  If Prefix = vbEmpty Then
    Exit Sub
  ElseIf i > 0 Then
    Postfix = Right(Prefix, Len(Prefix) - i)
    Prefix = Left(Prefix, i - 1)
  Else
    Postfix = vbNullString
    Prefix = vbNullString
  End If
  Dim KeywordsList()
  ReDim KeywordsList(0)
  For Each i In KeywordsDoc
      KeywordsList(UBound(KeywordsList)) = Prefix & i & Postfix
      ReDim Preserve KeywordsList(UBound(KeywordsList) + 1)
  Next
  Dim KeywordsCount
  KeywordsCount = UBound(KeywordsList)
  'Process each line
  Dim ProcessString, HasKeyword
  For i = 1 To ProcessDoc.linesCount
    'Get next line
    ProcessDoc.caretY(i)
    ProcessString = ProcessDoc.lineText
    'Default: assume line does not contains a keyword
    HasKeyword = False
    For j = 1 To KeywordsCount
      If InStr(ProcessString, KeywordsList(j-1)) > 0 Then
        'Contains keyword: set flag
        HasKeyword = True
        'And exit loop to avoid redundant loops 
        Exit For
      End If
    Next
    'ONLY LINE CHANGING BETWEEN Keep AND Remove. No keyword found: delete line
    If Not HasKeyword Then ProcessDoc.lineText(vbNullString)
  Next
  'Remove empty lines, including lines deleted earlier
  runPSPadAction "aRemoveBlankLines"
End Sub

Sub RemoveMatches
  Dim i, j
  'Script requires at least 2 opened documents
  If editorsCount < 2 Then
    MsgBox "You need at least 2 opened documents!" & vbCrLf & "The document to process (which must be active) and a document containing the keywords."
    Exit sub
  End If
  'Set document to process
  Dim ProcessDoc
  Set ProcessDoc = newEditor
  ProcessDoc.assignActiveEditor
  'Set source document
  Dim KeywordsDoc
  Set KeywordsDoc = newEditor
  For i = 0 To editorsCount - 1
    KeywordsDoc.assignEditorByIndex(i)
    If KeywordsDoc.fileName <> ProcessDoc.fileName Then Exit For
  Next
  'Populate keywords array
  Dim Prefix, Postfix
  Prefix = InputBox("This will keep lines in " & ProcessDoc.fileName & " that contain any keyword from " & KeywordsDoc.fileName & vbCrLf & "Enter prefix/postfix pattern in the form prefix*postfix:", , ">*.zip")
  i = InStr(Prefix, "*")
  If Prefix = vbEmpty Then
    Exit Sub
  ElseIf i > 0 Then
    Postfix = Right(Prefix, Len(Prefix) - i)
    Prefix = Left(Prefix, i - 1)
  Else
    Postfix = vbNullString
    Prefix = vbNullString
  End If
  Dim KeywordsList()
  ReDim KeywordsList(0)
  For Each i In KeywordsDoc
      KeywordsList(UBound(KeywordsList)) = Prefix & i & Postfix
      ReDim Preserve KeywordsList(UBound(KeywordsList) + 1)
  Next
  Dim KeywordsCount
  KeywordsCount = UBound(KeywordsList)
  'Process each line
  Dim ProcessString, HasKeyword
  For i = 1 To ProcessDoc.linesCount
    'Get next line
    ProcessDoc.caretY(i)
    ProcessString = ProcessDoc.lineText
    'Default: assume line does not contains a keyword
    HasKeyword = False
    For j = 1 To KeywordsCount
      If InStr(ProcessString, KeywordsList(j-1)) > 0 Then
        'Contains keyword: set flag
        HasKeyword = True
        'And exit loop to avoid redundant loops 
        Exit For
      End If
    Next
    'ONLY LINE CHANGING BETWEEN Keep AND Remove. Keyword found: delete line
    If HasKeyword Then ProcessDoc.lineText(vbNullString)
  Next
  'Remove empty lines, including lines deleted earlier
  runPSPadAction "aRemoveBlankLines"
End Sub

Sub ExtractRomsListFromDriver
  Dim MyDoc, InputString, i, j, k

  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'Process each line
  For i = 1 To MyDoc.linesCount
    'Home of current line
    MyDoc.caretY(i)
    InputString = MyDoc.lineText
    j = InStr(InputString, "ROM_START(")
    If j Then
      'Found: locate end of ROM name
      k = InStr(InputString, ")")
      'Get cleaned/trimmed ROM name
      MyDoc.lineText(Trim(Mid(InputString, j+10, k-j-10)))
    Else
      MyDoc.lineText(vbNullString)
    End If
  Next
  'Cleanup lines
  runPSPadAction "aRemoveBlankLines"
  runPSPadAction "aRemoveSpaces"
  runPSPadAction "aSort"
End Sub

Sub LinesToArrayTest
  Dim MyDoc
  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  Dim KeywordsCount
  
  'Using For Each on document object: almost instantaneous for 20000 rows 
  Dim item
  Dim arr1()
  ReDim arr1(0)
  For Each item In MyDoc
      arr1(UBound(arr1)) = item
      ReDim Preserve arr1(UBound(arr1) + 1)
  Next
  KeywordsCount = UBound(arr1)

  'Using caret loop on document lines: around 2 minutes for 20000 rows 
  KeywordsCount = MyDoc.linesCount
  Dim arr2()
  redim arr2(KeywordsCount)
  Dim i
  For i = 1 To KeywordsCount
    MyDoc.caretY(i)
    arr2(i-1) = MyDoc.lineText
  Next
End Sub

Sub Init
  addMenuItem "ClrMAME log->FileZilla 3 queue", module_name, "ClrMAMEtoFilezilla"
  addMenuItem "Keep lines including keywords", module_name, "KeepMatches"
  addMenuItem "Remove lines including keywords", module_name, "RemoveMatches"
  addMenuItem "Extract ROMs list from MAME driver", module_name, "ExtractRomsListFromDriver"
End Sub