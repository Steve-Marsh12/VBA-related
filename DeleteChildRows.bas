Attribute VB_Name = "DeleteChildRows"
Sub Delete_Child_Rows()
Attribute Delete_Child_Rows.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'
    Dim parentIDCell As String
    Dim parentIDCellValue As Integer
    Dim childIDCell As String
    Dim childIDCellValue As Integer

    
    Worksheets("RAG Rating").Activate
    Range("A1").Select
    Cells.Find(What:="Count Current Responses", After:=ActiveCell, LookIn:= _
        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(1, 0).Activate
    
    parentIDCell = "A" & CStr(ActiveCell.Row)
    childIDCell = "B" & CStr(ActiveCell.Row)
    
    parentIDCellValue = Range(parentIDCell).Value
    childIDCellValue = Range(childIDCell).Value

    While (Not (ActiveCell = ""))
    
'    MsgBox ("parentIDCell is " & parentIDCell)
'    MsgBox ("childIDCell is " & childIDCell)
'    MsgBox ("parentIDCellValue is " & CStr(parentIDCellValue))
'    MsgBox ("childIDCellValue is " & CStr(childIDCellValue))

        If (Not (parentIDCellValue = childIDCellValue)) Then
            ActiveCell.EntireRow.Select
            Selection.Delete
        End If
        
    ActiveCell.Offset(1, 0).Activate
    
    parentIDCell = "A" & CStr(ActiveCell.Row)
    childIDCell = "B" & CStr(ActiveCell.Row)
    
    parentIDCellValue = Range(parentIDCell).Value
    childIDCellValue = Range(childIDCell).Value
    
    Wend

    Range("A1").Activate
    
    Worksheets("ORSA_DB").Activate
    Range("A1").Activate

End Sub
