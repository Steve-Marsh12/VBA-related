Attribute VB_Name = "Print_Sheets"
Option Explicit
 
Sub Print_Weekly_Sheets()

Dim countAll As Integer
Dim intLoop As Integer
Dim noCopies As Integer

noCopies = CInt(InputBox(Prompt:="Please enter the number of copies required.  If no value is entered 9 copies will be produced", Title:="No of Copies", Default:="9"))

countAll = ActiveWorkbook.Sheets.Count


'For intLoop = 1 To countAll                       'Loop for first to last sheet.
'
'
'    Sheets(intLoop).Activate                        'select the worksheet equivalent to the count of the
'                                                    'number of loops i.e. on the first loop select the first
'                                                    'worksheet
'
'    If Not ((ActiveSheet.Name = "Exam Counts") Or (ActiveSheet.Name = "Exams") Or (ActiveSheet.Name = "Query Data") Or (ActiveSheet.Name = "Demand Data")) Then
'    Sheets(intLoop).PrintOut Copies:=noCopies
'    End If
'
'Next intLoop                                  'ends the loop


For intLoop = 1 To countAll                       'Loop for first to last sheet.


    Sheets(intLoop).Activate                        'select the worksheet equivalent to the count of the
                                                    'number of loops i.e. on the first loop select the first
                                                    'worksheet

    If ((ActiveSheet.Name = "Weekly Outstanding by mod") Or (ActiveSheet.Name = "Appointments") _
    Or (ActiveSheet.Name = "Pending") Or (ActiveSheet.Name = "Combined Appt and Pend") Or (ActiveSheet.Name = "Demand")) Then
    Sheets(intLoop).PrintOut Copies:=noCopies
    End If

Next intLoop                                  'ends the loop


Sheets(1).Activate


End Sub

