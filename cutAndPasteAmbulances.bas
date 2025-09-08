Attribute VB_Name = "cutAndPasteAmbulances"
Option Explicit

Sub Cut_And_Paste_Ambulances()

Dim ambTrustName As String             'variable to hold name of DB to be removed
Dim checkOffsetCount As Integer
Dim startCheckCell As String
Dim pasteRow As Integer

checkOffsetCount = 1

pasteRow = 2

'Sheets("Amb Trusts Removed from Subs").Select
'Cells.Select
'Selection.Delete Shift:=xlUp
'ActiveWindow.ScrollWorkbookTabs Position:=xlFirst

Sheets("ORSA_DB").Select
Rows("1:1").Select
Selection.Copy

Sheets("Amb Trusts Removed from Subs").Select
Rows("1:1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Sheets("ORSA_DB").Activate
Range("A1").Activate

'Find "Designated Body" and activate that cell
Cells.Find(What:="DesignatedBody", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
startCheckCell = ActiveCell.Address                     'gets the string address of the "ORSA Question"
'                                                       'header cell and assigns the value to startCell

Range(startCheckCell).Offset(checkOffsetCount, 0).Activate    'Activates the cell offset from the start cell
                                                              'by offsetCount rows.  Initially this is the next
                                                              'row after the header and this increases by 1
                                                              'row each loop.

Do While Not (ActiveCell = " " Or IsEmpty(ActiveCell))
    If ((InStr(1, ActiveCell.Value, "mbulance", 1)) And (ActiveCell.Offset(0, 35).Value = 0)) Then
        ActiveCell.EntireRow.Select
        Selection.Copy
    
        Sheets("Amb Trusts Removed from Subs").Activate
        Range("A1").Activate
        ActiveCell.EntireRow.Select                 'Select the entire 1st row (as the active cell is "A1"
            Do Until (WorksheetFunction.CountA(Selection) = 0)
                ActiveCell.Offset(1, 0).Activate
                ActiveCell.EntireRow.Select
            Loop
        ActiveSheet.Paste
        Application.CutCopyMode = False
        ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
        
        Sheets("ORSA_DB").Select
        Selection.Delete Shift:=xlUp
        pasteRow = pasteRow + 1
        checkOffsetCount = checkOffsetCount + 1
        Range(startCheckCell).Offset(checkOffsetCount, 0).Activate
    Else: checkOffsetCount = checkOffsetCount + 1
    Range(startCheckCell).Offset(checkOffsetCount, 0).Activate
    End If
    
Loop
    
Sheets("Amb Trusts Removed from Subs").Activate
Cells.Select
Cells.EntireColumn.AutoFit
Range("A1").Activate
Range("A1").Select
    
Sheets("ORSA_DB").Activate
Range("A1").Activate
Range("A1").Select

End Sub


