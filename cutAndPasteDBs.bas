Attribute VB_Name = "cutAndPasteDBs"
Option Explicit

Sub Cut_And_Paste_DBs()

Dim sourceOffsetCount As Integer          'records the number of loops.  This directly corresponds with the number of rows to offset
Dim startSourceCell As String             'this is the first cell containing the header for the ORSA question column.
Dim nonRegionDBName As String             'variable to hold name of DB to be removed
Dim checkOffsetCount As Integer
Dim startCheckCell As String
Dim checkDBName As String
Dim copyRowCellAddress As String
Dim copyRowCellRow As String
Dim pasteRow As Integer
Dim pasteRowstring As String


sourceOffsetCount = 1               'initialised to 1 and increases by 1 each loop.  This number is
                                    'used to offset from the startcell each loop to get the next question

checkOffsetCount = 1

pasteRow = 2


'Sheets("Non-reg Removed from Subs").Select
'Cells.Select
'Selection.Delete Shift:=xlUp
'ActiveWindow.ScrollWorkbookTabs Position:=xlFirst

Sheets("ORSA_DB").Select
Rows("1:1").Select
Selection.Copy

Sheets("Non-reg Removed from Subs").Select
Rows("1:1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Sheets("ORSA_DB").Activate
Range("A1").Activate

'Find "Designated Body" and activate that cell
Cells.Find(What:="DesignatedBody", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
startCheckCell = ActiveCell.Address                      'gets the string address of the "ORSA Question"
                                                    'header cell and assigns the value to startCell

Range(startCheckCell).Offset(checkOffsetCount, 0).Activate    'Activates the cell offset from the start cell
                                                              'by offsetCount rows.  Initially this is the next
                                                              'row after the header and this increases by 1
                                                              
                                                              'row each loop.


Sheets("Non-reg to Remove from Subs").Activate
Range("A1").Activate

'Find "Designated Body" and activate that cell
Cells.Find(What:="Designated Body", After:=ActiveCell, LookIn:=xlFormulas, _
    LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False).Activate
startSourceCell = ActiveCell.Address                'gets the string address of the "ORSA Question"
                                                    'header cell and assigns the value to startCell

Range(startSourceCell).Offset(sourceOffsetCount, 0).Activate    'Activates the cell offset from the start cell
                                                    'by offsetCount rows.  Initially this is the next
                                                    'row after the header and this increases by 1
                                                    'row each loop.

Do While Not (ActiveCell = " " Or IsEmpty(ActiveCell))
    nonRegionDBName = Trim(ActiveCell.Value)
    '    checkOffsetCount = checkOffsetCount + 1
    Sheets("ORSA_DB").Activate
    Range(startCheckCell).Offset(checkOffsetCount, 0).Activate



            Do While Not (ActiveCell = " " Or IsEmpty(ActiveCell))
                checkDBName = Trim(ActiveCell.Value)
                    
                    If (nonRegionDBName = checkDBName) Then
                        ActiveCell.EntireRow.Select
                        Selection.Copy
                    
                        Sheets("Non-reg Removed from Subs").Activate
                        Range("A1").Activate
                        ActiveCell.EntireRow.Select             'Select the entire 1st row (as the active cell is "A1"
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
                    Else: 'Do nothing
                    End If
                ActiveCell.Offset(1, 0).Activate
            Loop
            
    sourceOffsetCount = sourceOffsetCount + 1
    Sheets("Non-reg to Remove from Subs").Activate
    Range(startSourceCell).Offset(sourceOffsetCount, 0).Activate
    
'    MsgBox ("nonRegionDBName is " & nonRegionDBName + vbCrLf _
'    & "checkDBName is " & checkDBName)
    
    Loop
        

End Sub
