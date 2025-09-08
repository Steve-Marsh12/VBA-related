VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   11505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12870
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCutAndPasteSpecial_Click()

Call cut_and_paste_special_values

End Sub

Private Sub cmdUpdateAll_Click()

    Call Change_Tipping_Point
    ThisWorkbook.RefreshAll
    Call Update_Headers_and_Footers

End Sub

Private Sub cmdUpdateTipping_Click()

Call Change_Tipping_Point

End Sub

Private Sub cmdRefresh_Click()

ThisWorkbook.RefreshAll

End Sub

Private Sub cmdUpdateHeaders_Click()

Call Update_Headers_and_Footers

End Sub

Private Sub CmdPrint_Click()

Call Print_Weekly_Sheets

End Sub

Private Sub cmdCancelForm_Click()

UserForm1.Hide

End Sub

Private Sub cmdQuitNoSave_Click()

On Error Resume Next
       Dim xObject As Object

       Set xObject = ThisWorkbook.VBProject.References.Item("VBIDE")
       ThisWorkbook.VBProject.References.Remove xObject

Application.Quit
ActiveWorkbook.Close False

End Sub

Private Sub cmdQuitwithSave_Click()

On Error Resume Next
       Dim xObject As Object

       Set xObject = ThisWorkbook.VBProject.References.Item("VBIDE")
       ThisWorkbook.VBProject.References.Remove xObject

Application.Quit
ActiveWorkbook.Close True

End Sub

Private Sub cmdListRefs_Click()

Call ListReferencePaths

End Sub

Private Sub cmdDelRefs_Click()

    Sheets("References").Delete
    Worksheets("Weekly Outstanding by mod").Activate
    Range("A1").Activate

End Sub

Private Sub cmdRemoveComponents_Click()

Call remove_components

End Sub

Private Sub cmdChartRefresh_Click()

Call Update_Range_Names

End Sub


Private Sub CommandButton1_Click()

Call Delete_Hidden_Sheets

End Sub
