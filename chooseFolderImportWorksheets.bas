Attribute VB_Name = "chooseFolderImportWorksheets"
Option Explicit

Sub Choose_Folder_Import_Worksheets()

Dim fso As New FileSystemObject
Dim fls As Files
Dim strText As String
Dim noOfFiles As Integer
Dim f As File
Dim folderName As String
Dim fd As FileDialog
Dim pasteSheetCount As Integer
pasteSheetCount = 1
Dim pasteSheetName As String

'Set fd = Application.FileDialog(msoFileDialogFilePicker)
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
'fd.InitialFileName = "Z:\Resource Pool\Expatriate Services\ASSIGNEES\"
fd.InitialFileName = "C:\Users\marshst\Desktop\SHA Cluster Metrics\Cluster Returns"
Application.FileDialog(msoFileDialogFolderPicker).Show
folderName = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)

Set fls = fso.GetFolder(folderName).Files

noOfFiles = fls.Count
MsgBox ("Number of files is " & CStr(noOfFiles))

If (noOfFiles = 4) Then
    For Each f In fls
    Workbooks.Open (f)
    Select Case pasteSheetCount
        Case 1
            pasteSheetName = "Mapping"
        Case 2
            pasteSheetName = "North of England"
        Case 3
            pasteSheetName = "Midlands and East"
        Case 4
            pasteSheetName = "London"
    End Select
    
    ActiveWorkbook.Sheets(1).Copy After:=Workbooks("SHA Cluster metrics - Template.xls").Sheets(pasteSheetName)
    
'    Sheets("Sheet1").Select
'    Sheets("Sheet1").Copy After:=Workbooks("Filechooser.xlsm").Sheets(3)
'    Windows("London.xlsx").Activate
'    Sheets("Sheet1").Select
'    Sheets("Sheet1").Copy Before:=Workbooks("Filechooser.xlsm").Sheets(2)
    
    pasteSheetCount = pasteSheetCount + 1
    


    Select Case intLoop
        Case 1
            selectedSheet = "North of England"
        Case 2
            selectedSheet = "Midlands and East"
        Case 3
            selectedSheet = "London"
        Case 4
            selectedSheet = "South of England"
    End Select
    
    
    Next
Else
    MsgBox ("The folder does not contain the correct number of files.  Please check the folder and make sure that only the four clusetr returns are present before rerunning")
End If

End Sub
