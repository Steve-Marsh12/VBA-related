VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

       On Error Resume Next
       Dim xObject As Object

       Set xObject = ThisWorkbook.VBProject.References.Item("VBIDE")
       ThisWorkbook.VBProject.References.Remove xObject


End Sub


Private Sub Workbook_Open()


ThisWorkbook.VBProject.References.AddFromFile "C:\Program Files\Common Files\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"

UserForm1.Show

End Sub
