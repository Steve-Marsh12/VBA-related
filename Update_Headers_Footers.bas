Attribute VB_Name = "Update_Headers_Footers"
Option Explicit

Sub Update_Headers_and_Footers()

Dim newDate As String

Dim headerDateStart As Integer
Dim newHeader As String
Dim headerBase As String

Dim footerDateStart As Integer
Dim newFooter As String
Dim footerBase As String
Dim pathFooter As String

Dim countAll As Integer
Dim intLoop As Integer

newDate = InputBox(Prompt:="Please enter the new header/footer date as dd/mm/yy", Title:="Header Date", Default:="")
pathFooter = ActiveWorkbook.FullName
countAll = ActiveWorkbook.Sheets.Count

For intLoop = 1 To countAll                       'Loop for first to last sheet.
    
    
    Sheets(intLoop).Activate                'select the worksheet equivalent to the count of the
                                            'number of loops i.e. on the first loop select the first
                                            'worksheet
                                            
  
    If InStr(1, ActiveSheet.PageSetup.CenterHeader, "@", vbTextCompare) Then
    headerDateStart = InStr(1, ActiveSheet.PageSetup.CenterHeader, "@", vbTextCompare)
    headerBase = Mid(ActiveSheet.PageSetup.CenterHeader, 1, (headerDateStart + 1))
    newHeader = headerBase & newDate

    ActiveSheet.PageSetup.CenterHeader = newHeader
    End If

 
    If InStr(1, ActiveSheet.PageSetup.LeftFooter, "@", vbTextCompare) Then
    footerDateStart = InStr(1, ActiveSheet.PageSetup.LeftFooter, "@", vbTextCompare)
    footerBase = Mid(ActiveSheet.PageSetup.LeftFooter, 1, (footerDateStart + 1))
    newFooter = footerBase & newDate

    ActiveSheet.PageSetup.LeftFooter = newFooter + vbLf + pathFooter
    End If
    
'    MsgBox ("ActivSheet.Name is: " & ActiveSheet.Name)
'    MsgBox ("ActiveSheet.Type is: " & ActiveSheet.Type)
    
    If ActiveSheet.Type = -4167 Then
    ActiveSheet.Range("A1").Activate
    End If


Next intLoop                                  'ends the loop

Sheets(1).Activate


End Sub

