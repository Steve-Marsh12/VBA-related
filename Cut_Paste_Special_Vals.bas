Attribute VB_Name = "Cut_Paste_Special_Vals"
Option Explicit

Public Sub cut_and_paste_special_values()

Dim waitTimesWkBk As Workbook
Set waitTimesWkBk = ActiveWorkbook

Dim intLoop As Integer              'declare the variable intLoop to be type Integer.
                                        'This is used to hold the count for the number of loops
Dim noOfSheets As Integer           'declare the variable noOfSheets to be type Integer.
                                        'This is used to hold the count for the number of worksheets
                             
noOfSheets = waitTimesWkBk.Worksheets.Count   'This initialises the variable noOfSheets to
                                                'the count of the number of worksheets


For intLoop = 1 To noOfSheets           'Loop for 1 to the count of the number of sheets
                                            'i.e. from the first to the last
                                            
Worksheets(intLoop).Activate
Range("A1").Activate
                                            
Cells.Select                            'selects entire worksheet
Selection.Copy                          'copies the selection (entire worksheet)
    
    
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False                           'paste specials values
        
Range("A1").Select
    

Next intLoop                            'ends the loop

ActiveSheet.Range("A1").Select
Selection.Copy
ActiveSheet.Paste
Application.CutCopyMode = False

Worksheets("Weekly Outstanding by mod").Activate



End Sub



