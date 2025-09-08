Attribute VB_Name = "Change_Tipping_Date"
Option Explicit

Sub Change_Tipping_Point()
Attribute Change_Tipping_Point.VB_Description = "Change Tipping Point date"
Attribute Change_Tipping_Point.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Change Tipping Point Macro
'
'
    
    Dim tippingDate As String
    Dim newDateString As String
    Dim currentDate As String
    Dim dateStart As Integer
    Dim dateEnd As Integer
    
    Worksheets("Query Data").Activate
    Range("A1").Activate
    
    While ((Not (ActiveCell = "Tipping Point Grouping")) And (Not (ActiveCell = "")))
        ActiveCell.Offset(0, 1).Activate
    Wend
    
    
    If ActiveCell = "Tipping Point Grouping" Then
    ActiveCell.Offset(1, 0).Activate
    
    dateEnd = InStr(28, ActiveCell.Formula, ")", vbTextCompare)
    currentDate = Mid(ActiveCell.Formula, 28, ((dateEnd - 28) + 1))
    
    
    
    MsgBox ("The current tipping date is " & currentDate)
    tippingDate = InputBox(Prompt:="Please enter the new Tipping Date as yyyy,mm,dd", Title:="Tipping Date", Default:="")
    newDateString = "DATE( " & tippingDate & ")"
    MsgBox ("The new Tipping Date is: " & newDateString)
     
    Columns(ActiveCell.Column).Select
    Selection.Replace What:=currentDate, Replacement:=newDateString, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    End If
        
End Sub
