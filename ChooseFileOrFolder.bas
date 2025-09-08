Attribute VB_Name = "ChooseFileOrFolder"
Option Explicit
Public selectedFolder As String



Sub Choose_File_Or_Folder()

Dim varItem As Variant
Dim strPath As String
'Dim filePicker As FileDialog
Dim folderChoice As FileDialog
'Dim selectedFolder As String

'Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
Set folderChoice = Application.FileDialog(msoFileDialogFolderPicker)

'With filePicker
'
''setup File Dialog'
'.AllowMultiSelect = False
'.ButtonName = "Select File"
'.InitialView = msoFileDialogViewList
'.Title = "Select File"
'.InitialFileName = "C:\Users\Steve\Videos"

With folderChoice

'setup File Dialog'
.Title = "Select Folder 1"
.AllowMultiSelect = False
.InitialFileName = "C:\Users\Steve\Videos"
.InitialView = msoFileDialogViewList

''add filter for all files'
'With .Filters
'.Clear
'.Add "All Files", "*.*"
'End With
'.FilterIndex = 1

'display file dialog box'
.Show

End With

'If filePicker.SelectedItems.Count > 0 Then
'
'Dim selectedFile As String
'selectedFile = filePicker.SelectedItems(1)
'
''Me.PathToFile = selectedFile
'
'End If
'
'MsgBox ("The selected filepath is " & selectedFile)
'
'
''Set the object variable to Nothing.
'Set filePicker = Nothing


If folderChoice.SelectedItems.Count > 0 Then
selectedFolder = folderChoice.SelectedItems(1)
MsgBox ("The selected folder is " & selectedFolder)
End If

Set folderChoice = Nothing

End Sub





