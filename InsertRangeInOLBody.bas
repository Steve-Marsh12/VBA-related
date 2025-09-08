Attribute VB_Name = "InsertRangeInOLBody"
Sub insert_range_In_OLBody()


Dim rng As Range
Dim olApp As Object
Dim todaysEMail As Object
Dim FSObj As FileSystemObject
Dim TStream1 As TextStream
Dim TStream2 As TextStream
Dim TStream3 As TextStream
Dim strHTMLBody As String

Worksheets("Daily Email - by Cluster").Activate

''' Set Range you want to export to file
Set rng = Worksheets("Daily Email - by Cluster").UsedRange

'Now create the HTML file
ActiveWorkbook.PublishObjects.Add(xlSourceRange, "C:DailyEmailChart.htm", rng.Parent.Name, rng.Address, xlHtmlStatic).Publish True

'Create an instance of Outlook (or use existing instance if it already exists
Set olApp = CreateObject("Outlook.Application")

'Create a mail item
Set todaysEMail = olApp.CreateItemFromTemplate("C:\Users\marshst\AppData\Roaming\Microsoft\Templates\ORSA - Current reported position.oft")

'Open the HTML file using the FilesystemObject into a TextStream object
Set FSObj = New Scripting.FileSystemObject
Set TStream1 = FSObj.OpenTextFile("C:\DailEmailTextPart1.htm", ForReading)
Set TStream2 = FSObj.OpenTextFile("C:\DailyEmailChart.htm", ForReading)
Set TStream3 = FSObj.OpenTextFile("C:\DailEmailTextPart2.htm", ForReading)

'Now set the HTMLBody property of the message to the text contained in the TextStream object
strHTMLBody = TStream1.ReadAll & TStream2.ReadAll & TStream3.ReadAll

'By default the range will be centred. This line left aligns it and you can
'comment it out if you want the range centred.
strHTMLBody = Replace(strHTMLBody, "align=center", "align=left", , , vbTextCompare)

With todaysEMail
    .HTMLBody = strHTMLBody
'    .Attachments.Add "C:\Users\marshst\Documents\ORSA Daily Email Docs\Returns Received To Date.xlsx"
    .Attachments.Add "C:\Users\marshst\Documents\ORSA Daily Email Docs\Submissions to date.xlsx"
'    .Attachments.Add "C:\Users\marshst\Documents\ORSA Daily Email Docs\Outstanding Responses.xlsx"
    .Display
End With


Worksheets("ORSA_DB").Activate
Range("A1").Activate

End Sub





