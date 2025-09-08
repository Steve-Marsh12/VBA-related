Attribute VB_Name = "deleteObseleteRows"
Option Explicit

Sub Delete_Obselete_Rows()

Dim endCell As String

Range("A1").Activate

Cells.Find(What:="Trust Code", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    
endCell = Range(ActiveCell.Address).Offset(-1, 0).Address

Range("A1:" & endCell).Select
Selection.EntireRow.Delete

End Sub
