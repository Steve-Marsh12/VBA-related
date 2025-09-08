Attribute VB_Name = "transposeData"
Option Explicit

Sub Transpose_Data()

Dim copySheetName As String
Dim lastColumn As Integer
Dim firstColumn As Integer
Dim noLoops As Integer

Dim mainRangeStartCell As String
Dim mainRangeEndCell As String
Dim mainRange As String
Dim mainPasteCell As String

Dim mainRangeStartCell2 As String
Dim mainRange2 As String
Dim mainPasteCell2 As String
Dim pasteSheetName As String

Dim subRangeStartCell As String
Dim subRangeEndCell As String
Dim subRange As String

Dim pasteRowOffset As Integer
Dim i As Integer

Dim subRangePasteStartCell As String
Dim subRangePasteEndCell As String
Dim subRangePaste As String

Dim dateAddress As String
Dim dateString As String
Dim monthString As String
Dim yearString As String

Dim headerNameCell As String

Dim saveFileName As String

Dim applicableDate As String
Dim extractDate As String
Dim reportType As String

Dim lastRowCell As String
Dim derivedColumnCell As String

applicableDate = Year(Now()) & "0401"
extractDate = Year(Now()) & Month(Now()) & Day(Now())
reportType = "Dummy"

copySheetName = ActiveSheet.Name

Range("A1").Activate
ActiveCell.SpecialCells(xlLastCell).Select
lastColumn = ActiveCell.Column

Range("A1").Activate

Cells.Find(What:="Trust Code", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    
mainRangeStartCell = ActiveCell.Address

Cells.Find(What:="Trust Name", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
    
headerNameCell = ActiveCell.Address

firstColumn = Range(ActiveCell.Offset(0, 1).Address).Column

noLoops = lastColumn - firstColumn

    
ActiveCell.Select
Selection.End(xlDown).Select

pasteRowOffset = Selection.Row

mainRangeEndCell = Selection.Address
mainRange = mainRangeStartCell & ":" & mainRangeEndCell

mainRangeStartCell2 = Range(mainRangeStartCell).Offset(1, 0).Address
mainRange2 = mainRangeStartCell2 & ":" & mainRangeEndCell

    Range(mainRangeEndCell).Activate
    Range(ActiveCell.Address).End(xlUp).Select
    subRangePasteStartCell = Selection.Offset(1, 3).Address
    subRangePasteEndCell = Range(mainRangeEndCell).Offset(0, 3).Address
    subRangePaste = subRangePasteStartCell & ":" & subRangePasteEndCell


Range(mainRange).Select
Selection.Copy

Sheets.Add After:=Sheets(Sheets.Count)

pasteSheetName = ActiveSheet.Name
mainPasteCell = Range("A1").Address

Sheets(pasteSheetName).Select
Range(mainPasteCell).Select
ActiveSheet.Paste

mainPasteCell = Range("A2").Address
pasteRowOffset = pasteRowOffset - 1

Sheets(copySheetName).Activate
Range(mainRange2).Select
Selection.Copy


For i = 1 To (noLoops - 1)

    mainPasteCell = Range(mainPasteCell).Offset(pasteRowOffset, 0).Address

    Sheets(pasteSheetName).Select
    Range(mainPasteCell).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False


    Sheets(copySheetName).Activate
    Range(mainRange2).Select
    Selection.Copy

Next



For i = 1 To noLoops

    Sheets(copySheetName).Activate
    
    dateAddress = Range("A1").Offset(0, i + 3).Address
    dateString = Range(dateAddress).Value
    monthString = MonthName(Month(dateString))
    yearString = Year(dateString)

'    MsgBox ("Month is " & monthString & " and year is " & yearString)
    
    Range(mainRangeEndCell).Offset(0, i).Select
    subRangeStartCell = Selection.End(xlUp).Offset(1, 0).Address
    subRangeEndCell = Range(mainRangeEndCell).Offset(0, i).Address
    subRange = subRangeStartCell & ":" & subRangeEndCell
    Range(subRange).Select
    Selection.Copy

    Sheets(pasteSheetName).Select
    Range(subRangePaste).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range(subRangePaste).Offset(0, -1).Value = monthString
    Range(subRangePaste).Offset(0, -2).Value = yearString

    subRangePasteStartCell = Range(subRangePasteStartCell).Offset(pasteRowOffset, 0).Address
    subRangePasteEndCell = Range(subRangePasteEndCell).Offset(pasteRowOffset, 0).Address
    subRangePaste = subRangePasteStartCell & ":" & subRangePasteEndCell

    Sheets(copySheetName).Activate

Next

Sheets(pasteSheetName).Activate
Range("A1").Activate
ActiveCell.SpecialCells(xlLastCell).Activate
lastRowCell = ActiveCell.Address

ActiveCell.Offset(-((ActiveCell.Row) - 1), 1).Activate
derivedColumnCell = ActiveCell.Address

Range(ActiveCell.Address & ":" & Range(lastRowCell).Offset(0, 1).Address).Value = applicableDate

Range(derivedColumnCell).Offset(0, 1).Activate
Range(ActiveCell.Address & ":" & Range(lastRowCell).Offset(0, 2).Address).Value = extractDate

Range(derivedColumnCell).Offset(1, 2).Activate
'Range(ActiveCell.Address & ":" & Range(lastRowCell).Offset(0, 3).Address).Value = reportType
Range(ActiveCell.Address).Value = reportType

'Range(headerNameCell).Offset(0, 1).Value = "Year"
'Range(headerNameCell).Offset(0, 2).Value = "Month"
'Range(headerNameCell).Offset(0, 3).Value = "Cases"

Cells.Select
Cells.EntireColumn.AutoFit
    
Application.DisplayAlerts = False
Range("A1").EntireRow.Delete
Range("A1").Activate

ActiveSheet.Name = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)

Sheets(ActiveSheet.Name).Move

saveFileName = ("Y:\MRSA Prov\Input\CSVs\" & ActiveSheet.Name & ".csv")

ActiveWorkbook.SaveAs Filename:=saveFileName, FileFormat _
        :=xlCSV, CreateBackup:=False

Application.DisplayAlerts = False

ActiveWorkbook.Close

Application.DisplayAlerts = True
    
    


End Sub
