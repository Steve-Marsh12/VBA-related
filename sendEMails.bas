Attribute VB_Name = "sendEMails"
Sub Send_EMails()

    Dim DesignatedBodyName As String
    Dim DesignatedBodyCell As String
    Dim SHAName As String
    Dim SHANameCell As String
    Dim ClusterName As String
    Dim ClusterNameCell As String
    Dim EMailAdresseeString As String
    Dim EMailAdresseeCell As String
    Dim RecipientFirstName As String
    Dim RecipientFirstNameCell As String
    Dim FilePathString As String
       
    Dim olApp As Object
    Dim todaysEMail As Object
    Dim FSObj As FileSystemObject
    
    Worksheets("ORSA_DB").Activate
    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="DesignatedBody", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    DesignatedBodyCell = ActiveCell.Offset(1, 0).Address

    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="Area", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    SHANameCell = ActiveCell.Offset(1, 0).Address
        
    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="Cluster", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ClusterNameCell = ActiveCell.Offset(1, 0).Address
    
    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="ResponsibleOfficerEmail", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    EMailAdresseeCell = ActiveCell.Offset(1, 0).Address
    
    Range("A1").Activate
    Rows("1:1").Select
    Selection.Find(What:="ResponsibleOfficerFirstName", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    RecipientFirstNameCell = ActiveCell.Offset(1, 0).Address



    
    While Not (ActiveCell = "")
    
'Create an instance of Outlook (or use existing instance if it already exists
Set olApp = CreateObject("Outlook.Application")

'Create a mail item
Set todaysEMail = olApp.CreateItemFromTemplate("C:\Users\marshst\AppData\Roaming\Microsoft\Templates\ORSA covering email.oft")

        
        DesignatedBodyName = Range(DesignatedBodyCell).Value
        DesignatedBodyName = Replace(DesignatedBodyName, "/", " ")
        DesignatedBodyName = Replace(DesignatedBodyName, "\", " ")
        
        SHAName = Range(SHANameCell).Value
        ClusterName = Range(ClusterNameCell).Value
        EMailAdresseeString = Range(EMailAdresseeCell).Value
        RecipientFirstName = Range(RecipientFirstNameCell).Value
        
        FilePathString = ("C:\Users\marshst\Desktop\Mail Merge Docs\" & ClusterName & "\" & SHAName & "\" & DesignatedBodyName & ".doc")
        
'MsgBox ("DesignatedBodyName is " & DesignatedBodyName)
'MsgBox ("DesignatedBodyCell is " & DesignatedBodyCell)
'MsgBox ("SHAName is " & SHAName)
'MsgBox ("SHANameCell is " & SHANameCell)
'MsgBox ("ClusterName is " & ClusterName)
'MsgBox ("ClusterNameCell is " & ClusterNameCell)
'MsgBox ("EMailAdresseeString is " & EMailAdresseeString)
'MsgBox ("EMailAdresseeCell is " & EMailAdresseeCell)
'MsgBox ("RecipientFirstName is " & RecipientFirstName)
'MsgBox ("RecipientFirstNameCell is " & RecipientFirstNameCell)
'MsgBox ("FilePathString is " & FilePathString)
                
        'Open the HTML file using the FilesystemObject into a TextStream object
        Set FSObj = New Scripting.FileSystemObject
        Set TStream2 = FSObj.OpenTextFile("C:\CoveringEMailText.htm", ForReading)


'Now set the HTMLBody property of the message to the text contained in the TextStream object
strHTMLBody = TStream2.ReadAll

'By default the range will be centred. This line left aligns it and you can
'comment it out if you want the range centred.
strHTMLBody = Replace(strHTMLBody, "align=center", "align=left", , , vbTextCompare)
strHTMLBody = Replace(strHTMLBody, "Recipient", RecipientFirstName, , , vbTextCompare)


todaysEMail.HTMLBody = strHTMLBody
todaysEMail.Attachments.Add FilePathString
todaysEMail.To = EMailAdresseeString
todaysEMail.Subject = "ORSA 2011 - 2012 Final Results And Comparisons"
'todaysEMail.Display
todaysEMail.Save
todaysEMail.Close olPromtForSave


        DesignatedBodyCell = Range(DesignatedBodyCell).Offset(1, 0).Address
        SHANameCell = Range(SHANameCell).Offset(1, 0).Address
        ClusterNameCell = Range(ClusterNameCell).Offset(1, 0).Address
        EMailAdresseeCell = Range(EMailAdresseeCell).Offset(1, 0).Address
        RecipientFirstNameCell = Range(RecipientFirstNameCell).Offset(1, 0).Address
        Range(DesignatedBodyCell).Activate
        
'MsgBox ("DesignatedBodyCell is " & DesignatedBodyCell)
'MsgBox ("SHANameCell is " & SHANameCell)
'MsgBox ("ClusterNameCell is " & ClusterNameCell)
'MsgBox ("EMailAdresseeCell is " & EMailAdresseeCell)
'MsgBox ("RecipientFirstNameCell is " & RecipientFirstNameCell)


        Wend

Worksheets("ORSA_DB").Activate
Range("A1").Activate


        
End Sub
