VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmClientSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:   This module illustrates how to create a search form, _
            where the user can enter as many or few criteria as they wish, _
            and results are shown one per line.
'Note:      Only records matching ALL of the criteria are returned.
'Author:    Allen Browne (allen@allenbrowne.com), June 2006.
Option Compare Database
Option Explicit

Private Sub cmdFilter_Click()
    'Purpose:   Build up the criteria string form the non-blank search boxes, and apply to the form's Filter.
    'Notes:     1. We tack " AND " on the end of each condition so you can easily add more search boxes; _
                        we remove the trailing " AND " at the end.
    '           2. The date range works like this: _
                        Both dates      = only dates between (both inclusive. _
                        Start date only = all dates from this one onwards; _
                        End date only   = all dates up to (and including this one).
    Dim strWhere As String                  'The criteria string.
    Dim lngLen As Long                      'Length of the criteria string to append to.
    Const conJetDate = "\#mm\/dd\/yyyy\#"   'The format expected for dates in a JET query string.
    
    '***********************************************************************
    'Look at each search box, and build up the criteria string from the non-blank ones.
    '***********************************************************************
    'Text field example. Use quotes around the value in the string.
    If Not IsNull(Me.txtFilterCity) Then
        strWhere = strWhere & "([City] = """ & Me.txtFilterCity & """) AND "
    End If
    
    'Another text field example. Use Like to find anywhere in the field.
    If Not IsNull(Me.txtFilterMainName) Then
        strWhere = strWhere & "([MainName] Like ""*" & Me.txtFilterMainName & "*"") AND "
    End If
    
    'Number field example. Do not add the extra quotes.
    If Not IsNull(Me.cboFilterLevel) Then
        strWhere = strWhere & "([LevelID] = " & Me.cboFilterLevel & ") AND "
    End If
    
    'Yes/No field and combo example. If combo is blank or contains "ALL", we do nothing.
    If Me.cboFilterIsCorporate = -1 Then
        strWhere = strWhere & "([IsCorporate] = True) AND "
    ElseIf Me.cboFilterIsCorporate = 0 Then
        strWhere = strWhere & "([IsCorporate] = False) AND "
    End If
    
    'Date field example. Use the format string to add the # delimiters and get the right international format.
    If Not IsNull(Me.txtStartDate) Then
        strWhere = strWhere & "([EnteredOn] >= " & Format(Me.txtStartDate, conJetDate) & ") AND "
    End If
    
    'Another date field example. Use "less than the next day" since this field has times as well as dates.
    If Not IsNull(Me.txtEndDate) Then   'Less than the next day.
        strWhere = strWhere & "([EnteredOn] < " & Format(Me.txtEndDate + 1, conJetDate) & ") AND "
    End If
    
    '***********************************************************************
    'Chop off the trailing " AND ", and use the string as the form's Filter.
    '***********************************************************************
    'See if the string has more than 5 characters (a trailng " AND ") to remove.
    lngLen = Len(strWhere) - 5
    If lngLen <= 0 Then     'Nah: there was nothing in the string.
        MsgBox "No criteria", vbInformation, "Nothing to do."
    Else                    'Yep: there is something there, so remove the " AND " at the end.
        strWhere = Left$(strWhere, lngLen)
        'For debugging, remove the leading quote on the next line. Prints to Immediate Window (Ctrl+G).
        'Debug.Print strWhere
        
        'Finally, apply the string as the form's Filter.
        Me.Filter = strWhere
        Me.FilterOn = True
    End If
End Sub

Private Sub cmdReset_Click()
    'Purpose:   Clear all the search boxes in the Form Header, and show all records again.
    Dim ctl As Control
    
    'Clear all the controls in the Form Header section.
    For Each ctl In Me.Section(acHeader).Controls
        Select Case ctl.ControlType
        Case acTextBox, acComboBox
            ctl.Value = Null
        Case acCheckBox
            ctl.Value = False
        End Select
    Next
    
    'Remove the form's filter.
    Me.FilterOn = False
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    'To avoid problems if the filter returns no records, we did not set its AllowAdditions to No.
    'We prevent new records by cancelling the form's BeforeInsert event instead.
    'The problems are explained at http://allenbrowne.com/bug-06.html
    Cancel = True
    MsgBox "You cannot add new clients to the search form.", vbInformation, "Permission denied."
End Sub

Private Sub Form_Open(Cancel As Integer)
    'Remove the single quote from these lines if you want to initially show no records.
    'Me.Filter = "(False)"
    'Me.FilterOn = True
End Sub
