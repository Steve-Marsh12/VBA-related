Attribute VB_Name = "listFolders1"
Sub List_Folders1()
    Dim fso As New FileSystemObject
    Dim flds As Folders
    Dim subFdrs As Folders
    Dim strText As String
    Dim strSize As String
    Dim subFolderText As String
    Dim noOfSubFolders As Integer
    Dim noOfFiles As Integer
    Dim i As Integer
    
    Set flds = fso.GetFolder("E:\Media\Music\").SubFolders
    

    i = 1
    

    
    For Each f In flds
        strText = f.Path
        strSize = f.Size
        noOfSubFolders = f.SubFolders.Count

'            If (noOfSubFolders > 1) Then
                Set subFdrs = f.SubFolders
                    For Each sf In subFdrs
                                        
                    subFolderText = sf.Path
                    noOfFiles = sf.Files.Count
                    
'                    MsgBox ("i is " & Str(i))
'                    MsgBox ("Folder Path is " & strText)
'                    MsgBox ("No of subfolders is " & Str(noOfSubFolders))
'                    MsgBox ("Sub Folder Path is " & subFolderText)
'                    MsgBox ("No of files is " & Str(noOfFiles))

                    Worksheets("Drive 1").Cells(i, 1) = strText
                    Worksheets("Drive 1").Cells(i, 2) = strSize
                    Worksheets("Drive 1").Cells(i, 3) = noOfSubFolders
                    Worksheets("Drive 1").Cells(i, 4) = subFolderText
                    Worksheets("Drive 1").Cells(i, 5) = noOfFiles
                    i = i + 1
                    Next
'                    i = i + noOfSubFolders
'            Else
'                    subFolderText = sf.Path
'                    noOfFiles = sf.Files.Count
'                    Worksheets("Sheet4").Cells(i, 1) = strText
'                    Worksheets("Sheet4").Cells(i, 2) = strSize
'                    Worksheets("Sheet4").Cells(i, 3) = noOfSubFolders
'                    Worksheets("Sheet4").Cells(i, 4) = subFolderText
'                    Worksheets("Sheet4").Cells(i, 5) = noOfFiles
'                    i = i + 1
'            End If
    Next
    
    Columns("D:D").Select
    Selection.Replace What:= _
        "E:\Media\Music\" _
        , Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False

End Sub


