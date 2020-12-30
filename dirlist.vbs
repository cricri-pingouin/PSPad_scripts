Option Explicit
Const module_name = "DirList"
Const module_ver = "1.00"

Sub ListFiles
  Dim MyDoc, objFSO, objFolder, objFile, colFiles
  Dim OutputFormat, OutputItems, OutputString, FileN, FileE, FileP, FileF, FileB, FileK, FileM, FileD, FileA, i

'N: name w/o extension
'E: full name w/ extension
'P: name with path
'F: folder name
'B: size in bytes
'K: size in KB
'M: size in MB
'D: date last modified
'A: attributes
'Other: as typed
  'Initialise document
  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'User input source folder
  Dim SourceFolder
  SourceFolder = SelectFolder("")
  'No folder selected? Exit
  If SourceFolder = vbNull Then Exit Sub
  'Get format string
  OutputFormat = InputBox("N: name w/o extension" & vbCrLf & "E: full name w/ extension"  & vbCrLf & "P: name with path" _
  & vbCrLf & "F: folder name" & vbCrLf & "B: size in bytes"  & vbCrLf & "K: size in KB"  & vbCrLf & "M: size in MB"  _
  & vbCrLf & "D: date last modified" & vbCrLf & "A: attributes" & vbCrLf & "Other: as typed", _
  "Enter the output format string", "N,1,M")
  'Parse format string
  If Len(OutputFormat) = 0 Then
      Exit Sub
  Else
      OutputFormat = Split(UCase(OutputFormat), ",")
  End If
  OutputItems = UBound(OutputFormat)
  'Prepare file object
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objFSO.GetFolder(SourceFolder)
  Set colFiles = objFolder.Files
  OutputString=""
  'Process each file in user selected folder
  For Each objFile in colFiles
    'Prepare information for file
    FileN = objFSO.getbasename(objFile.Name)
    FileE = objFile.Name
    FileP = objFile.Path
    FileF = objFolder.Name
    FileB = objFile.Size
    FileK = Round(objFile.Size/1024,1)
    FileM = Round(objFile.Size/1048576,1)
    FileD = objFile.DateLastModified
    FileA = objFile.Attributes
    'Build output string for file
    For i=0 to OutputItems
      Select Case OutputFormat(i)
        Case "N"
          OutputString = OutputString & FileN 
        Case "E"
          OutputString = OutputString & FileE
        Case "P"
          OutputString = OutputString & FileP
        Case "F"
          OutputString = OutputString & FileF
        Case "B"
          OutputString = OutputString & FileB
        Case "K"
          OutputString = OutputString & FileK
        Case "M"
          OutputString = OutputString & FileM
        Case "D"
          OutputString = OutputString & FileD
        Case "A"
          OutputString = OutputString & FileA
        Case Else
          OutputString = OutputString & OutputFormat(i)
      End Select
      If i<OutputItems Then OutputString = OutputString & ","
    Next
    OutputString = OutputString & vbCrLf
  Next
  'Append to file
  MyDoc.text(MyDoc.text & OutputString)
End Sub

Sub ListFolders
  Dim MyDoc, objFSO, objFolder, colSubfolders, objSubfolder, colFiles, objFile
  Dim OutputFormat, OutputItems, OutputString, FileN, FileE, FileP, FileF, FileB, FileK, FileM, FileD, FileA, i

'N: name w/o extension
'E: full name w/ extension
'P: name with path
'F: folder name
'B: size in bytes
'K: size in KB
'M: size in MB
'D: date last modified
'A: attributes
'Other: as typed
  'Initialise document
  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'User input source folder
  Dim SourceFolder
  SourceFolder = SelectFolder("")
  'No folder selected? Exit
  If SourceFolder = vbNull Then Exit Sub
  'Get format string
  OutputFormat = InputBox("N: name w/o extension" & vbCrLf & "E: full name w/ extension"  & vbCrLf & "P: name with path" _
  & vbCrLf & "F: folder name" & vbCrLf & "B: size in bytes"  & vbCrLf & "K: size in KB"  & vbCrLf & "M: size in MB"  _
  & vbCrLf & "D: date last modified" & vbCrLf & "A: attributes" & vbCrLf & "Other: as typed", _
  "Enter the output format string", "N,F,1,M")
  'Parse format string
  If Len(OutputFormat) = 0 Then
      Exit Sub
  Else
      OutputFormat = Split(UCase(OutputFormat), ",")
  End If
  OutputItems = UBound(OutputFormat)
  'Prepare file object
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objFSO.GetFolder(SourceFolder)
  Set colSubfolders = objFolder.SubFolders
  OutputString=""
  'Process each subfolder in user selected folder
  For Each objSubfolder in colSubfolders
    'Process each file in subfolder
    Set colFiles = objSubfolder.Files
    For Each objFile in colFiles
      'Prepare information for file
      FileN = objFSO.getbasename(objFile.Name)
      FileE = objFile.Name
      FileP = objFile.Path
      FileF = objSubfolder.Name
      FileB = objFile.Size
      FileK = Round(objFile.Size/1024,1)
      FileM = Round(objFile.Size/1048576,1)
      FileD = objFile.DateLastModified
      FileA = objFile.Attributes
      'Build output string for file
      For i=0 to OutputItems
        Select Case OutputFormat(i)
          Case "N"
            OutputString=OutputString & FileN
          Case "E"
            OutputString=OutputString & FileE
          Case "P"
            OutputString=OutputString & FileP
          Case "F"
            OutputString=OutputString & FileF
          Case "B"
            OutputString=OutputString & FileB
          Case "K"
            OutputString=OutputString & FileK
          Case "M"
            OutputString=OutputString & FileM
          Case "D"
            OutputString=OutputString & FileD
          Case "A"
            OutputString=OutputString & FileA
          Case Else
            OutputString=OutputString & OutputFormat(i)
        End Select
        If i<OutputItems Then OutputString = OutputString & ","
      Next
      OutputString = OutputString & vbCrLf
    Next
  Next
  'Append to file
  MyDoc.text(MyDoc.text & OutputString)
End Sub

Function SelectFolder(myStartFolder)
  Dim objFolder, objItem, objShell

  On Error Resume Next
  SelectFolder = vbNull
  'Create a dialog object
  Set objShell = CreateObject( "Shell.Application" )
  Set objFolder = objShell.BrowseForFolder( 0, "Select Folder", 0, myStartFolder )
  'Return the path of the selected folder
  If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path
  Set objFolder = Nothing
  Set objshell = Nothing
  On Error Goto 0
End Function

Sub Init
  addMenuItem "List files in selected folder", module_name, "ListFiles"
  addMenuItem "List files in subfolders of selected folder", module_name, "ListFolders"
End Sub