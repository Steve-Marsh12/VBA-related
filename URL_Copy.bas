Attribute VB_Name = "URL_Copy"
Sub URL()
Attribute URL.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copy URL Address
'
'    Dim firstCell As String
'    Dim secondCell As String
    
    Dim currentRow As Integer
    currentRow = 1
    
    Dim cellItem As String
    
    Range("A1").Activate
    
    
    Do Until (IsEmpty(ActiveCell))
    
           cellItem = ActiveCell.Cells.Hyperlinks(1).Address
           ActiveCell.Offset(0, 1).Activate
           Range(ActiveCell.Address).Value = cellItem
           ActiveCell.Offset(1, -1).Activate
        
    Loop
    
  
End Sub

