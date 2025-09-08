Attribute VB_Name = "Delete_Hidden_Worksheets"
Option Explicit
 
Sub Delete_Hidden_Sheets()

Dim countAll As Integer
Dim intLoop As Integer
Dim noCopies As Integer

countAll = ActiveWorkbook.Sheets.Count

MsgBox ("Count of Sheets is: " & countAll)

For intLoop = 1 To countAll                       'Loop for first to last sheet.

MsgBox ("Sheet " & intLoop & " is " & (Sheets(intLoop).Visible))

If (Not (Sheets(intLoop).Visible)) Then
    Sheets(intLoop).Delete
End If


'    Sheets(intLoop).Activate                        'select the worksheet equivalent to the count of the
'                                                    'number of loops i.e. on the first loop select the first
'                                                    'worksheet
'
'    If ((ActiveSheet.Name = "Weekly Outstanding by mod") Or (ActiveSheet.Name = "Appointments") _
'    Or (ActiveSheet.Name = "Pending") Or (ActiveSheet.Name = "Combined Appt and Pend") Or (ActiveSheet.Name = "Demand")) Then
'    Sheets(intLoop).PrintOut Copies:=noCopies
'    End If

Next intLoop                                  'ends the loop


Sheets(1).Activate


End Sub



