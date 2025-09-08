Attribute VB_Name = "InsertPictureInOutlook"
Sub insert_Picture_In_OLBody()

Dim rng As Range
Dim olApp As Object
Dim todaysEMail As Object

Worksheets("Daily Email").Activate
'Set rng = Worksheets("Daily Email").UsedRange
'MsgBox ("rng is " & rng.Address)
'rng.Select
'Selection.Copy


''' Set Range you want to export to file
Set rng = Worksheets("Daily Email").UsedRange

''' Copy range as picture onto Clipboard
rng.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
''' Create an empty chart with exact size of range copied
 With ActiveSheet.ChartObjects.Add(Left:=rng.Left, Top:=rng.Top, Width:=rng.Width, Height:=rng.Height)
    .Name = "DailyEmailChart"
    .Activate
End With

''' Paste into chart area, export to file, delete chart.
ActiveChart.Paste
ActiveSheet.ChartObjects("DailyEmailChart").Chart.Export "C:\DailyEmailChart.jpg"
ActiveSheet.ChartObjects("DailyEmailChart").Delete

End Sub




