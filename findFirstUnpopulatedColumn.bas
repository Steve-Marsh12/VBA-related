Attribute VB_Name = "findFirstUnpopulatedColumn"
Option Explicit

' This macro locates the the last populated row on a spreadsheet (the spreadsheet is determined by the value of
' selectedSheet - a global variable declared in the sub Copy_And_Paste_Cluster_Data) and selects that row.

Sub Find_First_Unpopulated_Column()

Public lastColumn As String
Public lastCell As String


Range("A1").Activate
ActiveCell.EntireColumn.Select             'Select the entire 1st row (as the active cell is "A1"

Do Until (WorksheetFunction.CountA(Selection) = 0)
    ActiveCell.Offset(0, 1).Activate
    ActiveCell.EntireRow.Select
Loop

lastCell = ActiveCell.Address
lastColumn = CStr(ActiveCell.Column)
'MsgBox ("first unpopulated cell is " & lastCell +vbCrLf _
& "first unpopulated column is " & lastColumn)

End Sub
