Attribute VB_Name = "deleteEmptyRows"
Option Explicit

' This macro locates the the last populated row on a spreadsheet (the spreadsheet is determined by the value of
' selectedSheet - a global variable declared in the sub Copy_And_Paste_Cluster_Data) and selects that row.

Sub Delete_Empty_Rows()

Dim checkCell As String
Dim startColumn As Integer
Dim endColumn As Integer
Dim columnOffset As Integer

Cells.Find(What:="PHE Centre", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    

    
checkCell = ActiveCell.Address
startColumn = ActiveCell.Column
    
ActiveCell.SpecialCells(xlLastCell).Activate
endColumn = ActiveCell.Column
columnOffset = (endColumn - startColumn)

ActiveCell.Offset(0, -columnOffset).Activate

Do Until ActiveCell.Row = 1
    If ActiveCell.Value = "" Or ActiveCell.Value = " " Then
'        MsgBox ("Empty Row")
        ActiveCell.EntireRow.Select
        Selection.Delete
        ActiveCell.Offset(-1, startColumn - 1).Activate
    Else
        ActiveCell.Offset(-1, 0).Activate
    End If
Loop

End Sub
