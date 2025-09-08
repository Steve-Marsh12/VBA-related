Attribute VB_Name = "outstandingResponses"
Option Explicit


Sub Outstanding_Responses()
Attribute Outstanding_Responses.VB_ProcData.VB_Invoke_Func = " \n14"
'
'
'

'

    Dim lastRow As String
    Dim copyRange As String
    
    Sheets("Outstanding Responses").Activate
    Cells.Select
    Selection.Delete Shift:=xlUp

    Sheets("List of Expected Responses").Activate
    ActiveSheet.Range("$A:$L").AutoFilter Field:=1, Criteria1:="="
    
    Range("A1").Activate
    ActiveCell.EntireRow.Select             'Select the entire 1st row (as the active cell is "A1"
    
    Do Until (WorksheetFunction.CountA(Selection) = 0)
        ActiveCell.Offset(1, 0).Activate
        ActiveCell.EntireRow.Select
    Loop
    
    ActiveCell.Offset(-1, 0).Activate
    lastRow = CStr(ActiveCell.Row)
    
    copyRange = "C1:F" & lastRow
    Range(copyRange).Select
    Selection.Copy
    
    Sheets("Outstanding Responses").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Cells.Find(What:="ORSA (Sep 12) RAG", After:=ActiveCell, LookIn:= _
    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
    xlNext, MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.EntireColumn.Delete
    
    Selection.Replace What:="DB Name as advised by Cluster", Replacement:= _
        "DB Name", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    Selection.Replace What:= _
        "Reason for inclusion (as per email reply or last month performance report)", _
        Replacement:="Category", LookAt:=xlPart, SearchOrder:=xlByRows, _
        MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Columns("A:C").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Columns("A:C").Select
    ActiveWorkbook.Worksheets("Outstanding Responses").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Outstanding Responses").Sort.SortFields.Add Key:= _
        Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Outstanding Responses").Sort.SortFields.Add Key:= _
        Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Outstanding Responses").Sort.SortFields.Add Key:= _
        Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Outstanding Responses").Sort
        .SetRange Range("A:C")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Selection.ColumnWidth = 100#
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
    
    Cells.Select
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    
    Columns("A:C").Select
    Selection.AutoFilter
    Range("A1").Select

    Application.DisplayAlerts = False
    
    Sheets("Outstanding Responses").Select
    Sheets("Outstanding Responses").Copy
    ChDir "C:\"
    ActiveWorkbook.SaveAs Filename:="C:\Users\marshst\Documents\ORSA Daily Email Docs\Outstanding Responses.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True

    Sheets("List of Expected Responses").Select
    Cells.Select
    ActiveSheet.ShowAllData
    Range("A1").Select

End Sub

