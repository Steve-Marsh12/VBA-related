Attribute VB_Name = "saveIndividualDocuments"
Option Explicit
 
Sub Save_Individual_Documents()

     
    Dim x As Long
    Dim i As Long
    Dim DesignatedBodyName As String
    Dim SHAName As String
    Dim ClusterName As String
    
    With ActiveDocument.MailMerge
   
   .Destination = wdSendToNewDocument
   .SuppressBlankLines = True
   
   'get the record count of the datasource
   With .DataSource
     .ActiveRecord = wdLastRecord
     x = .ActiveRecord
     'set the activerecord back to the first
     .ActiveRecord = wdFirstRecord
    
   End With
   
   'loop the datasource count and merge one record at a time
   For i = 1 To x

     .DataSource.FirstRecord = i
     .DataSource.LastRecord = i
     

     
     DesignatedBodyName = ActiveDocument.MailMerge.DataSource.DataFields("DesignatedBody").Value
     SHAName = ActiveDocument.MailMerge.DataSource.DataFields("Area").Value
     ClusterName = ActiveDocument.MailMerge.DataSource.DataFields("Cluster").Value

'     MsgBox ("Designated Body is " & CStr(ActiveDocument.MailMerge.DataSource.DataFields("DesignatedBody").Value))


     DesignatedBodyName = Replace(DesignatedBodyName, "/", " ")
     DesignatedBodyName = Replace(DesignatedBodyName, "\", " ")
'     DesignatedBodyName = Replace(DesignatedBodyName, "'", "")
     



     .Execute Pause:=True
'     ActiveDocument.SaveAs ("C:\Users\marshst\Desktop\Mail Merge Docs\" & DesignatedBodyName & ".doc")
     ActiveDocument.SaveAs ("C:\Users\marshst\Desktop\Mail Merge Docs\" & ClusterName & "\" & SHAName & "\" & DesignatedBodyName & ".doc")
     ActiveDocument.Close wdSaveChanges

     ActiveDocument.MailMerge.DataSource.ActiveRecord = wdNextDataSourceRecord


   Next i
 End With
 
 
End Sub

