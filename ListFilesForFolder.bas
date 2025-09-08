Attribute VB_Name = "ListFilesForFolder"
Option Explicit

'List all of the files and details for a specified folder
'
'


Sub ListFilesInFolder()

    Dim folderString As String
    Dim fso As New FileSystemObject
    Dim fls As Files
    Dim strText As String
    Dim i As Integer
    Dim noOfFiles As Integer
    Dim rangeString As String
    Dim f As File
    
    folderString = InputBox("Please enter the path for the folder you want details for e.g. C:\Users\Steve\Videos\Fringe\Season 4\")
    
    Set fls = fso.GetFolder(folderString).Files
    
    noOfFiles = fls.Count
    MsgBox ("Number of files is " & CStr(noOfFiles))
    
    i = 2
    
'    With Worksheets("Sheet1")
'        .Cells(1, 1) = "File Name"
'        .Cells(1, 2) = "File Size"
'        .Cells(1, 3) = "Date"
'
'    End With
'
''    MsgBox ("Name is " & CStr(fls.Item(i).Name))
''    MsgBox ("Name is " & CStr(fls(i).Name))
''    MsgBox ("Name is " & fls.Item(i).Name)
''    MsgBox ("Name is " & fls(i).Name)
'
'
'
'
'    For i = 2 To (noOfFiles + 1)
'            rangeString = "A" & CStr(i)
'            MsgBox ("rangeString is " & rangeString)
'            Range(rangeString).Value = fls(i).Name
'
''            Cells(i, 1).Value = fls.Item(i).Name
''            Cells(i, 2).Text = fls.Item(i).Size
''            Cells(i, 3).Text = fls.Item(i).DateLastModified
'
'            i = i + 1
'    Next


    With Worksheets("Sheet1")
        .Cells(1, 1) = "File Name"
        .Cells(1, 2) = "File Size"
        .Cells(1, 3) = "Date"
        For Each f In fls
            .Cells(i, 1) = f.Name
            .Cells(i, 2) = f.Size
            .Cells(i, 3) = f.DateLastModified
            i = i + 1
        Next
    End With
    
'    For Each f In fls
''        .Cells(i, 1) = f.Name
''        .Cells(i, 2) = f.Size
''        .Cells(i, 3) = f.DateLastModified
'        MsgBox ("File name is " & CStr(f.Name))
'        MsgBox ("File name is " & CStr(f.Size))
'        MsgBox ("File name is " & CStr(f.DateLastModified))
'    Worksheets("Sheet1").Cells(i, 1) = f.Name
'
'
'
'        i = i + 1
'    Next


End Sub


