Option Explicit
Const module_name = "MoveFiles"
Const module_ver = "1.00"

Sub MoveMatches
  'Initialise document
  Dim MyDoc
  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'Populate keywords array
  Dim Prefix, Postfix, i
  Prefix = InputBox("This will move any file which name is in the active document." & vbCrLf & "Enter prefix/postfix pattern in the form prefix*postfix:", , "*.zip")
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
  'Build file names list
  Dim FilesNames()
  ReDim FilesNames(0)
  For Each i In MyDoc
      FilesNames(UBound(FilesNames)) = Prefix & i & Postfix
      ReDim Preserve FilesNames(UBound(FilesNames) + 1)
  Next
  'Count records
  Dim FilesCount
  FilesCount = UBound(FilesNames)
  If FilesCount<1 Then Exit sub
  'Folders user selection
  Dim SourceFolder, DestinationFolder
  SourceFolder = SelectFolder("", "Select source folder:")
  If SourceFolder = vbNull Then Exit Sub
  DestinationFolder = SelectFolder("", "Select destination folder:")
  If DestinationFolder = vbNull Then Exit Sub
  'Terminate folders strings
  If Left(SourceFolder,1)<> "\" Then SourceFolder = SourceFolder + "\"
  If Left(DestinationFolder,1)<> "\" Then DestinationFolder = DestinationFolder + "\"
  'Initialise counters
  Dim Success
  Success = 0
  Dim Fail
  Fail = 0
  'Prepare file object
  Dim objFSO
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  'Process each file in user selected folder
  For i = 0 to FilesCount
    'File exists?
    If objFSO.FileExists(SourceFolder & FilesNames(i)) Then
      'Yes: move
      objFSO.MoveFile SourceFolder & FilesNames(i), DestinationFolder & FilesNames(i)
      Success = Success + 1
    Else
      Fail = Fail + 1
    End If 
  Next
  'Display results
  Msgbox "Files found/moved: " & Success & vbCrLf &"Files not found/moved: " & Fail,, module_name
End Sub

Sub KeepMatches
  'Initialise document
  Dim MyDoc
  Set MyDoc = newEditor
  MyDoc.assignActiveEditor
  If MyDoc Is Nothing Then Exit Sub
  'Populate keywords array
  Dim Prefix, Postfix, i
  Prefix = InputBox("This will move any file which name is NOT in the active document." & vbCrLf & "Enter prefix/postfix pattern in the form prefix*postfix:", , "*.zip")
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
  'Build file names list
  Dim FilesNames()
  ReDim FilesNames(0)
  For Each i In MyDoc
      FilesNames(UBound(FilesNames)) = Prefix & i & Postfix
      ReDim Preserve FilesNames(UBound(FilesNames) + 1)
  Next
  'Count records
  Dim FilesCount
  FilesCount = UBound(FilesNames)
  If FilesCount<1 Then Exit sub
  'Folders user selection
  Dim SourceFolder, DestinationFolder
  SourceFolder = SelectFolder("", "Select source folder:")
  If SourceFolder = vbNull Then Exit Sub
  DestinationFolder = SelectFolder("", "Select destination folder:")
  If DestinationFolder = vbNull Then Exit Sub
  'Terminate folders strings
  If Left(SourceFolder,1)<> "\" Then SourceFolder = SourceFolder + "\"
  If Left(DestinationFolder,1)<> "\" Then DestinationFolder = DestinationFolder + "\"
  'Initialise counters
  Dim Success
  Success = 0
  Dim Fail
  Fail = 0
  'Prepare file object
  Dim objFSO, objFolder, colFiles, objFile
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objFSO.GetFolder(SourceFolder)
  Set colFiles = objFolder.Files
  'Process each file in user selected folder
  Dim FileInList
  For Each objFile in colFiles
    FileInList = False
    For i = 0 to FilesCount
      'Prepare information for file
      If objFile.Name = FilesNames(i) Then
        FileInList = True
        Exit For
      End If
    Next
    'File exists?
    If Not FileInList Then
      'No: move
      objFSO.MoveFile SourceFolder & objFile.Name, DestinationFolder & objFile.Name
      Success = Success + 1
    Else
      Fail = Fail + 1
    End If
  Next
  'Display results
  Msgbox "Files not found/moved: " & Success & vbCrLf &"Files found/not moved: " & Fail,, module_name
End Sub

Function SelectFolder(myStartFolder, Message)
  Dim objFolder, objItem, objShell

  On Error Resume Next
  SelectFolder = vbNull
  'Create a dialog object
  Set objShell = CreateObject( "Shell.Application" )
  Set objFolder = objShell.BrowseForFolder( 0, Message, 0, myStartFolder )
  'Return the path of the selected folder
  If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path
  Set objFolder = Nothing
  Set objshell = Nothing
  On Error Goto 0
End Function

Sub Init
  addMenuItem "Move matches", module_name, "MoveMatches"
  addMenuItem "Keep matches", module_name, "KeepMatches"
End Sub