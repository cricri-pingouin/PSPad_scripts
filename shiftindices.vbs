Option Explicit
Const module_name = "Indices shifting"
Const module_ver = "1.00"

Sub ShiftIndices
    Dim StartVal, FinishVal, ShiftVal, i, MyDoc, IndexType

    Set MyDoc = newEditor()
    MyDoc.assignActiveEditor()
    If MyDoc Is Nothing Then Exit Sub
    StartVal = CInt(InputBox("Enter start index value:", module_name, 0))
    If StartVal < 0 Then Exit Sub
    FinishVal = CInt(InputBox("Enter finish index value:", module_name, 99))
    If StartVal > FinishVal Then Exit Sub
    ShiftVal = CInt(InputBox("Enter shift value (can be negative):", module_name, 1))
    If ShiftVal = 0 Then Exit Sub
    IndexType = MsgBox("What indices should be changed (note: [Column][Row]) ?" & vbCrLf & "Yes: columns, No: rows, Cancel: both", vbYesNoCancel + vbQuestion, module_name)
    If ShiftVal > 0 Then 
      For i = FinishVal To StartVal Step -1
        runPSPadAction "aSelectAll"
        If IndexType = vbYes Then MyDoc.selText Replace(MyDoc.selText, "[" & i & "][", "[" & (i+ShiftVal) & "][")
        If IndexType = vbNo Then MyDoc.selText Replace(MyDoc.selText, "][" & i & "]", "][" & (i+ShiftVal) & "]")
        If IndexType = vbCancel Then MyDoc.selText Replace(MyDoc.selText, "[" & i & "]", "[" & (i+ShiftVal) & "]")
      Next
    Else
      For i = StartVal To FinishVal
        runPSPadAction "aSelectAll"
        If IndexType = vbYes Then MyDoc.selText Replace(MyDoc.selText, "[" & i & "][", "[" & (i+ShiftVal) & "][")
        If IndexType = vbNo Then MyDoc.selText Replace(MyDoc.selText, "][" & i & "]", "][" & (i+ShiftVal) & "]")
        If IndexType = vbCancel Then MyDoc.selText Replace(MyDoc.selText, "[" & i & "]", "[" & (i+ShiftVal) & "]")
      Next    
    End If
End Sub

Sub ShiftVarIndex
    Dim VarName, StartVal, FinishVal, ShiftVal, i, MyDoc

    Set MyDoc = newEditor()
    MyDoc.assignActiveEditor()
    If MyDoc Is Nothing Then Exit Sub
    VarName = InputBox("Enter variable name:", module_name)
    If VarName = "" Then Exit Sub
    StartVal = CInt(InputBox("Enter start index value:", module_name, 0))
    If StartVal < 0 Then Exit Sub
    FinishVal = CInt(InputBox("Enter finish index value:", module_name, 99))
    If StartVal > FinishVal Then Exit Sub
    ShiftVal = CInt(InputBox("Enter shift value (can be negative):", module_name))
    If ShiftVal = 0 Then Exit Sub
    If ShiftVal > 0 Then 
      For i = FinishVal To StartVal Step -1
        runPSPadAction "aSelectAll"
        MyDoc.selText Replace(MyDoc.selText, VarName & "[" & i & "]", VarName & "[" & (i+ShiftVal) & "]")
      Next
    Else
      For i = StartVal To FinishVal
        runPSPadAction "aSelectAll"
        MyDoc.selText Replace(MyDoc.selText, VarName & "[" & i & "]", VarName & "[" & (i+ShiftVal) & "]")
      Next    
    End If
End Sub

Sub Init
    addMenuItem "Shift [i][j] indices", module_name, "ShiftIndices"
    addMenuItem "Shift Var[i] indices", module_name, "ShiftVarIndex"
End Sub