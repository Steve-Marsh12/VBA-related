Attribute VB_Name = "DeleteReferences"
Sub Delete_References()

Application.DisplayAlerts = False

Dim countAll As Integer
Dim intLoop As Integer

countAll = ActiveWorkbook.Sheets.Count

'MsgBox ("Count of Sheets is: " & countAll)

For intLoop = 1 To countAll


Sheets(intLoop).Activate                        'select the worksheet equivalent to the count of the
                                                    'number of loops i.e. on the first loop select the first
                                                    'worksheet

If (ActiveSheet.Name = "References") Then
Worksheets("References").Delete
End If

Next intLoop                                  'ends the loop

Application.DisplayAlerts = True

End Sub
