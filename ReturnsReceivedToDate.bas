Attribute VB_Name = "ReturnsReceivedToDate"
Sub Returns_Received_To_Date()
Attribute Returns_Received_To_Date.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Dim areaCell As String
    Dim trustNameCell As String
    Dim healthSectorCell As String
    Dim lastSubmissionTimeStampCell As String
    
    
    Worksheets("Returns received to date").Activate
    Cells.Select
    Selection.Delete Shift:=xlUp
    
    Worksheets("ORSA_DB").Activate
    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="Area", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    areaCell = ActiveCell.Address

    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="DesignatedBody", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    trustNameCell = ActiveCell.Address

    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="HealthSector", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    healthSectorCell = ActiveCell.Address
    
    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="LastSubmissionTimeStamp", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    lastSubmissionTimeStampCell = ActiveCell.Address

    Columns(Range(areaCell).Column).Select
    Selection.Copy
    Sheets("Returns received to date").Select
    Columns("A:A").Select
    ActiveSheet.Paste

    Sheets("ORSA_DB").Select
    Columns(Range(trustNameCell).Column).Select
    Selection.Copy
    Sheets("Returns received to date").Select
    Columns("B:B").Select
    ActiveSheet.Paste

    Sheets("ORSA_DB").Select
    Columns(Range(healthSectorCell).Column).Select
    Selection.Copy
    Sheets("Returns received to date").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    
    Sheets("ORSA_DB").Select
    Columns(Range(lastSubmissionTimeStampCell).Column).Select
    Selection.Copy
    Sheets("Returns received to date").Select
    Columns("D:D").Select
    ActiveSheet.Paste

    Cells.Replace What:="DesignatedBody", Replacement:="TrustName", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Cells.Select
    Application.CutCopyMode = False
'    Columns("A:D").Select
'    Range("D1").Activate
    ActiveWorkbook.Worksheets("Returns received to date").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Returns received to date").Sort.SortFields.Add Key _
        :=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    ActiveWorkbook.Worksheets("Returns received to date").Sort.SortFields.Add Key _
        :=Range("D:D"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Returns received to date").Sort
        .SetRange Range("A:D")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Worksheets("Daily Email").Activate
    Range("A1").Select
    Range("A1").Activate
    
    Worksheets("Returns received to date").Activate
    Range("A1").Select
    Range("A1").Activate
    
    Columns("A:D").Select
    Selection.AutoFilter
    Range("A1").Select
    
    Application.DisplayAlerts = False
    
    Sheets("Returns received to date").Select
    Sheets("Returns received to date").Copy
    ChDir "C:\"
    ActiveWorkbook.SaveAs Filename:="C:\Users\marshst\Documents\ORSA Daily Email Docs\Returns Received To Date.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True

End Sub
