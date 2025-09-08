Attribute VB_Name = "SendDailyEmailSheet"
Sub send_Daily_Email_Sheet()

Sheets("Daily Email").Select
Sheets("Daily Email").Copy
Sheets("Daily Email").Select
Sheets("Daily Email").Name = "Current ORSA responses"
Range("A1").Select
ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\marshst\Desktop\Current ORSA responses.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
' Show the envelope on the ActiveWorkbook.
ActiveWorkbook.EnvelopeVisible = True

' Set the optional introduction field thats adds
' some header text to the email body. It also sets
' the To and Subject lines. Finally the message
' is sent.
With ActiveSheet.MailEnvelope
'EmployeeName = "Bob"
'.Introduction = "Please update my record to show the following changes:"
.Item.to = "SHA Leads"
.Item.Subject = "ORSA - Current reported position"
.Item.Display
End With

End Sub


