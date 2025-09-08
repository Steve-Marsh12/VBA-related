Attribute VB_Name = "findLastPopulatedRow"
Option Explicit

' This macro locates the the last populated row on a spreadsheet (the spreadsheet is determined by the value of
' selectedSheet - a global variable declared in the sub Copy_And_Paste_Cluster_Data) and selects that row.

Sub Find_Last_Populated_Row()

Public lastRow As String

Sheets("List of Expected Responses").Activate
Range("A1").Activate
ActiveCell.EntireRow.Select             'Select the entire 1st row (as the active cell is "A1"

Do Until (WorksheetFunction.CountA(Selection) = 0)
    ActiveCell.Offset(1, 0).Activate
    ActiveCell.EntireRow.Select
Loop

ActiveCell.Offset(-1, 0).Activate
lastRow = CStr(ActiveCell.Row)
'MsgBox ("last populated row is " & lastRow)

End Sub
