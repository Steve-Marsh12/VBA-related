Attribute VB_Name = "Delete_components"
Option Explicit

Public Sub remove_components()

       On Error Resume Next
       Dim xObject As Object

       Set xObject = ThisWorkbook.VBProject.References.Item("VBIDE")
       ThisWorkbook.VBProject.References.Remove xObject

Dim CodeMod As CodeModule
Dim vbComp As VBComponent

Sheets("References").Delete

Set CodeMod = ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
With CodeMod
   .DeleteLines 1, .CountOfLines
End With


Dim vbCom As VBComponents
Set vbCom = Application.VBE.ActiveVBProject.VBComponents
vbCom.Remove VBComponent:=vbCom.Item("UserForm1")
vbCom.Remove VBComponent:=vbCom.Item("Change_Tipping_Date")
vbCom.Remove VBComponent:=vbCom.Item("find_GUID_references")
vbCom.Remove VBComponent:=vbCom.Item("Update_Headers_Footers")
vbCom.Remove VBComponent:=vbCom.Item("Cut_Paste_Special_Vals")
vbCom.Remove VBComponent:=vbCom.Item("Update_Charts")
vbCom.Remove VBComponent:=vbCom.Item("Update_All_Range_Names")
vbCom.Remove VBComponent:=vbCom.Item("Update_Ranges_And_All_Charts")
vbCom.Remove VBComponent:=vbCom.Item("Print_Sheets")
vbCom.Remove VBComponent:=vbCom.Item("Delete_components")


End Sub



