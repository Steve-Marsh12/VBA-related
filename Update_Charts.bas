Attribute VB_Name = "Update_Charts"
Sub Update_All_Charts()

Dim waitTimesWkBk As Workbook
Set waitTimesWkBk = ActiveWorkbook

Dim intLoop As Integer              'declare the variable intLoop to be type Integer.
                                        'This is used to hold the count for the number of loops
Dim noOfSheets As Integer           'declare the variable noOfSheets to be type Integer.
                                        'This is used to hold the count for the number of worksheets
                                        
Dim intChartLoop As Integer
Dim noOfCharts As Integer

Dim WkBkName As String
Dim FullPath As String
                             
noOfSheets = waitTimesWkBk.Worksheets.Count   'This initialises the variable noOfSheets to
                                                'the count of the number of worksheets


For intLoop = 1 To noOfSheets           'Loop for 1 to the count of the number of sheets
                                            'i.e. from the first to the last
    Worksheets(intLoop).Activate
    noOfCharts = ActiveSheet.ChartObjects.Count
    
'    MsgBox ("No of Charts is: " & noOfCharts)
    
        For intChartLoop = 1 To noOfCharts                       'Loop for first to last chart.


        ActiveSheet.ChartObjects(intChartLoop).Activate     'select the worksheet equivalent to the count of the
                                                            'number of loops i.e. on the first loop select the first
                                                            'worksheet
        WkBkName = ActiveWorkbook.Name
        FullPath = """" & "='" & WkBkName & "'!"

'        MsgBox ("The ActiveChart Name is: " & ActiveChart.Name)

        Select Case ActiveChart.Name

            Case "All mods - MR Charts Chart 1"

'                MsgBox ("All Mods Chart")
                
                ActiveChart.SeriesCollection(1).XValues = FullPath & "All_Mods_Label" & """"
                ActiveChart.SeriesCollection(1).Values = FullPath & "All_Mods_Appt" & """"
                ActiveChart.SeriesCollection(1).Name = "Appt"
                
                ActiveChart.SeriesCollection(2).XValues = FullPath & "All_Mods_Label" & """"
                ActiveChart.SeriesCollection(2).Values = FullPath & "All_Mods_Pend" & """"
                ActiveChart.SeriesCollection(2).Name = "Pend"
                
                ActiveChart.SeriesCollection(3).XValues = FullPath & "All_Mods_Label" & """"
                ActiveChart.SeriesCollection(3).Values = FullPath & "All_Mods_Combined" & """"
                ActiveChart.SeriesCollection(3).Name = "Combined"
                
            Case "All mods - MR Charts Chart 2"
                
'                MsgBox ("MR Chart")
                ActiveChart.SeriesCollection(1).XValues = FullPath & "MR_Label" & """"
                ActiveChart.SeriesCollection(1).Values = FullPath & "MR_Appt" & """"
                ActiveChart.SeriesCollection(1).Name = "Appt"
                
                ActiveChart.SeriesCollection(2).XValues = FullPath & "MR_Label" & """"
                ActiveChart.SeriesCollection(2).Values = FullPath & "MR_Pend" & """"
                ActiveChart.SeriesCollection(2).Name = "Pend"
                
                ActiveChart.SeriesCollection(3).XValues = FullPath & "MR_Label" & """"
                ActiveChart.SeriesCollection(3).Values = FullPath & "MR_Combined" & """"
                ActiveChart.SeriesCollection(3).Name = "Combined"
            
            Case "US - Fluoro Charts Chart 1"

'                MsgBox ("US Chart")
                ActiveChart.SeriesCollection(1).XValues = FullPath & "US_Label" & """"
                ActiveChart.SeriesCollection(1).Values = FullPath & "US_Appt" & """"
                ActiveChart.SeriesCollection(1).Name = "Appt"
                
                ActiveChart.SeriesCollection(2).XValues = FullPath & "US_Label" & """"
                ActiveChart.SeriesCollection(2).Values = FullPath & "US_Pend" & """"
                ActiveChart.SeriesCollection(2).Name = "Pend"
                
                ActiveChart.SeriesCollection(3).XValues = FullPath & "US_Label" & """"
                ActiveChart.SeriesCollection(3).Values = FullPath & "US_Combined" & """"
                ActiveChart.SeriesCollection(3).Name = "Combined"
                
            Case "US - Fluoro Charts Chart 2"
                
'                MsgBox ("Fluoro Chart")
                ActiveChart.SeriesCollection(1).XValues = FullPath & "Fluoro_Label" & """"
                ActiveChart.SeriesCollection(1).Values = FullPath & "Fluoro_Appt" & """"
                ActiveChart.SeriesCollection(1).Name = "Appt"
                
                ActiveChart.SeriesCollection(2).XValues = FullPath & "Fluoro_Label" & """"
                ActiveChart.SeriesCollection(2).Values = FullPath & "Fluoro_Pend" & """"
                ActiveChart.SeriesCollection(2).Name = "Pend"
                
                ActiveChart.SeriesCollection(3).XValues = FullPath & "Fluoro_Label" & """"
                ActiveChart.SeriesCollection(3).Values = FullPath & "Fluoro_Combined" & """"
                ActiveChart.SeriesCollection(3).Name = "Combined"
                
            Case "CT - Inter Charts Chart 1"

'                MsgBox ("CT Mods Chart")
                ActiveChart.SeriesCollection(1).XValues = FullPath & "CT_Label" & """"
                ActiveChart.SeriesCollection(1).Values = FullPath & "CT_Appt" & """"
                ActiveChart.SeriesCollection(1).Name = "Appt"
                
                ActiveChart.SeriesCollection(2).XValues = FullPath & "CT_Label" & """"
                ActiveChart.SeriesCollection(2).Values = FullPath & "CT_Pend" & """"
                ActiveChart.SeriesCollection(2).Name = "Pend"
                
                ActiveChart.SeriesCollection(3).XValues = FullPath & "CT_Label" & """"
                ActiveChart.SeriesCollection(3).Values = FullPath & "CT_Combined" & """"
                ActiveChart.SeriesCollection(3).Name = "Combined"
                
            Case "CT - Inter Charts Chart 2"
                
'                MsgBox ("Inter Chart")
                ActiveChart.SeriesCollection(1).XValues = FullPath & "Inter_Label" & """"
                ActiveChart.SeriesCollection(1).Values = FullPath & "Inter_Appt" & """"
                ActiveChart.SeriesCollection(1).Name = "Appt"
                
                ActiveChart.SeriesCollection(2).XValues = FullPath & "Inter_Label" & """"
                ActiveChart.SeriesCollection(2).Values = FullPath & "Inter_Pend" & """"
                ActiveChart.SeriesCollection(2).Name = "Pend"
                
                ActiveChart.SeriesCollection(3).XValues = FullPath & "Inter_Label" & """"
                ActiveChart.SeriesCollection(3).Values = FullPath & "Inter_Combined" & """"
                ActiveChart.SeriesCollection(3).Name = "Combined"
                
        End Select
        
        ActiveSheet.Range("A47").Activate
    
        Next intChartLoop


Next intLoop

Worksheets("Weekly Outstanding by mod").Activate
Range("A1").Activate

End Sub
