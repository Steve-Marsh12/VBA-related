Attribute VB_Name = "convertToCorrectFormat"
Option Explicit

Sub Convert_To_Correct_Format()

Dim varItem As Variant
Dim strPath As String
Dim filePicker As FileDialog

Dim noSheets As Integer

Dim sheetNumberSelected As Integer
Dim sheetSelected As String
Dim ws As Worksheet

Dim startCell As String


Set filePicker = Application.FileDialog(msoFileDialogFilePicker)

With filePicker
'setup File Dialog'
.AllowMultiSelect = False
.ButtonName = "Select File"
.InitialView = msoFileDialogViewList
.Title = "Select File"
.InitialFileName = "Y:\MRSA Prov\Input\"

'add filter for all files'
With .Filters
.Clear
.Add "All Files", "*.*"
End With
.FilterIndex = 1

'display file dialog box
.Show

End With

If filePicker.SelectedItems.Count > 0 Then
    Dim selectedFile As String
    selectedFile = filePicker.SelectedItems(1)
End If

'MsgBox ("The selected filepath is " & selectedFile)

'Set the object variable to Nothing.
Set filePicker = Nothing

Workbooks.Open (selectedFile)

noSheets = ActiveWorkbook.Sheets.Count

If noSheets > 1 Then
    MsgBox ("There are " & noSheets & " sheets in this workbook." + vbCr + _
    "Only 1 sheet is allowed.  On the next dialog box please type in the sheet number that you wish to keep")
End If

sheetNumberSelected = InputBox("Please enter the sheet number to keep.", "Sheet Number To Keep")
sheetSelected = Sheets(sheetNumberSelected).Name
'MsgBox ("Sheet Name is " & sheetSelected)

Application.DisplayAlerts = False

For Each ws In Workbooks(ActiveWorkbook.Name).Sheets
    If ws.Name = sheetSelected Then
        'do nothing
    Else
        ws.Delete
    End If
Next

Application.DisplayAlerts = True

Call Delete_Empty_Rows
Call Delete_Obselete_Rows
Call Transpose_Data

Application.DisplayAlerts = False

ActiveWorkbook.Close

Application.DisplayAlerts = True


End Sub
