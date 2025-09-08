Attribute VB_Name = "FindGUIDreferences"
Option Explicit
 
Sub ListReferencePaths()
     'Macro purpose:  To determine full path and Globally Unique Identifier (GUID)
     'to each referenced library.  Select the reference in the Tools\References
     'window, then run this code to get the information on the reference's library
     'Care: ensure that there is an existing worksheet called "references" before running
     
    Worksheets.Add(After:=Sheets("Returns received to date")).Name = "References"
    
'    Dim referenceSheet As Worksheet
'    referenceSheet.Name ("References")
     
    On Error Resume Next
    Dim i As Long
    With ThisWorkbook.Sheets("References")
        .Cells.Clear
        .Range("A1") = "Reference name"
        .Range("B1") = "Full path to reference"
        .Range("C1") = "Reference GUID"
    End With
    For i = 1 To ThisWorkbook.VBProject.References.Count
        With ThisWorkbook.VBProject.References(i)
            ThisWorkbook.Sheets("References").Range("A65536").End(xlUp).Offset(1, 0) = .Name
            ThisWorkbook.Sheets("References").Range("A65536").End(xlUp).Offset(0, 1) = .FullPath
            ThisWorkbook.Sheets("References").Range("A65536").End(xlUp).Offset(0, 2) = .GUID
        End With
    Next i
    On Error GoTo 0
End Sub

