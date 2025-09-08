Attribute VB_Name = "Update_All_Range_Names"

Sub Update_Range_Names()

Sheets("Weekly Outstanding by mod").Activate



Dim allModLabelStartString As String
Dim MRLabelStartString As String
Dim USLabelStartString As String
Dim FluoroLabelStartString As String
Dim CTLabelStartString As String
Dim InterLabelStartString As String

Dim allModApptStartString As String
Dim MRApptStartString As String
Dim USApptStartString As String
Dim FluoroApptStartString As String
Dim CTApptStartString As String
Dim InterApptStartString As String

Dim allModPendStartString As String
Dim MRPendStartString As String
Dim USPendStartString As String
Dim FluoroPendStartString As String
Dim CTPendStartString As String
Dim InterPendStartString As String

Dim allModCombinedStartString As String
Dim MRCombinedStartString As String
Dim USCombinedStartString As String
Dim FluoroCombinedStartString As String
Dim CTCombinedStartString As String
Dim InterCombinedStartString As String

Dim allModLabelEndString As String
Dim MRLabelEndString As String
Dim USLabelEndString As String
Dim FluoroLabelEndString As String
Dim CTLabelEndString As String
Dim InterLabelEndString As String

Dim allModApptEndString As String
Dim MRApptEndString As String
Dim USApptEndString As String
Dim FluoroApptEndString As String
Dim CTApptEndString As String
Dim InterApptEndString As String

Dim allModPendEndString As String
Dim MRPendEndString As String
Dim USPendEndString As String
Dim FluoroPendEndString As String
Dim CTPendEndString As String
Dim InterPendEndString As String

Dim allModCombinedEndString As String
Dim MRCombinedEndString As String
Dim USCombinedEndString As String
Dim FluoroCombinedEndString As String
Dim CTCombinedEndString As String
Dim InterCombinedEndString As String

Dim allModLabelRangeString As String
Dim MRLabelRangeString As String
Dim USLabelRangeString As String
Dim FluoroLabelRangeString As String
Dim CTLabelRangeString As String
Dim InterLabelRangeString As String

Dim allModApptRangeString As String
Dim MRApptRangeString As String
Dim USApptRangeString As String
Dim FluoroApptRangeString As String
Dim CTApptRangeString As String
Dim InterApptRangeString As String

Dim allModPendRangeString As String
Dim MRPendRangeString As String
Dim USPendRangeString As String
Dim FluoroPendRangeString As String
Dim CTPendRangeString As String
Dim InterPendRangeString As String

Dim allModCombinedRangeString As String
Dim MRCombinedRangeString As String
Dim USCombinedRangeString As String
Dim FluoroCombinedRangeString As String
Dim CTCombinedRangeString As String
Dim InterCombinedRangeString As String



Dim namedRangeRef As String


Range("A1").Activate

Cells.Find(What:="All Mods", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        
ActiveCell.Offset(2, 0).Activate


allModLabelStartString = ActiveCell.Address
allModApptStartString = ActiveCell.Offset(0, 1).Address
allModPendStartString = ActiveCell.Offset(0, 2).Address
allModCombinedStartString = ActiveCell.Offset(0, 3).Address

While (IsDate(ActiveCell.Value))
    ActiveCell.Offset(1, 0).Activate
Wend

ActiveCell.Offset(-1, 0).Activate
allModLabelEndString = ActiveCell.Address
allModApptEndString = ActiveCell.Offset(0, 1).Address
allModPendEndString = ActiveCell.Offset(0, 2).Address
allModCombinedEndString = ActiveCell.Offset(0, 3).Address

allModLabelRangeString = allModLabelStartString & ":" & allModLabelEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + allModLabelRangeString
ActiveWorkbook.Names.Add Name:="All_Mods_Label", RefersTo:=namedRangeRef

allModApptRangeString = allModApptStartString & ":" & allModApptEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + allModApptRangeString
ActiveWorkbook.Names.Add Name:="All_Mods_Appt", RefersTo:=namedRangeRef

allModPendRangeString = allModPendStartString & ":" & allModPendEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + allModPendRangeString
ActiveWorkbook.Names.Add Name:="All_Mods_Pend", RefersTo:=namedRangeRef

allModCombinedRangeString = allModCombinedStartString & ":" & allModCombinedEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + allModCombinedRangeString
ActiveWorkbook.Names.Add Name:="All_Mods_Combined", RefersTo:=namedRangeRef



Range("A1").Activate

Cells.Find(What:="MR", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        
ActiveCell.Offset(2, 0).Activate


MRLabelStartString = ActiveCell.Address
MRApptStartString = ActiveCell.Offset(0, 1).Address
MRPendStartString = ActiveCell.Offset(0, 2).Address
MRCombinedStartString = ActiveCell.Offset(0, 3).Address

While (IsDate(ActiveCell.Value))
    ActiveCell.Offset(1, 0).Activate
Wend

ActiveCell.Offset(-1, 0).Activate
MRLabelEndString = ActiveCell.Address
MRApptEndString = ActiveCell.Offset(0, 1).Address
MRPendEndString = ActiveCell.Offset(0, 2).Address
MRCombinedEndString = ActiveCell.Offset(0, 3).Address

MRLabelRangeString = MRLabelStartString & ":" & MRLabelEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + MRLabelRangeString
ActiveWorkbook.Names.Add Name:="MR_Label", RefersTo:=namedRangeRef

MRApptRangeString = MRApptStartString & ":" & MRApptEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + MRApptRangeString
ActiveWorkbook.Names.Add Name:="MR_Appt", RefersTo:=namedRangeRef

MRPendRangeString = MRPendStartString & ":" & MRPendEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + MRPendRangeString
ActiveWorkbook.Names.Add Name:="MR_Pend", RefersTo:=namedRangeRef

MRCombinedRangeString = MRCombinedStartString & ":" & MRCombinedEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + MRCombinedRangeString
ActiveWorkbook.Names.Add Name:="MR_Combined", RefersTo:=namedRangeRef




Range("A1").Activate

Cells.Find(What:="US", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        
ActiveCell.Offset(2, 0).Activate


USLabelStartString = ActiveCell.Address
USApptStartString = ActiveCell.Offset(0, 1).Address
USPendStartString = ActiveCell.Offset(0, 2).Address
USCombinedStartString = ActiveCell.Offset(0, 3).Address

While (IsDate(ActiveCell.Value))
    ActiveCell.Offset(1, 0).Activate
Wend

ActiveCell.Offset(-1, 0).Activate
USLabelEndString = ActiveCell.Address
USApptEndString = ActiveCell.Offset(0, 1).Address
USPendEndString = ActiveCell.Offset(0, 2).Address
USCombinedEndString = ActiveCell.Offset(0, 3).Address

USLabelRangeString = USLabelStartString & ":" & USLabelEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + USLabelRangeString
ActiveWorkbook.Names.Add Name:="US_Label", RefersTo:=namedRangeRef

USApptRangeString = USApptStartString & ":" & USApptEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + USApptRangeString
ActiveWorkbook.Names.Add Name:="US_Appt", RefersTo:=namedRangeRef

USPendRangeString = USPendStartString & ":" & USPendEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + USPendRangeString
ActiveWorkbook.Names.Add Name:="US_Pend", RefersTo:=namedRangeRef

USCombinedRangeString = USCombinedStartString & ":" & USCombinedEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + USCombinedRangeString
ActiveWorkbook.Names.Add Name:="US_Combined", RefersTo:=namedRangeRef




Range("A1").Activate

Cells.Find(What:="Fluoro", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        
ActiveCell.Offset(2, 0).Activate


FluoroLabelStartString = ActiveCell.Address
FluoroApptStartString = ActiveCell.Offset(0, 1).Address
FluoroPendStartString = ActiveCell.Offset(0, 2).Address
FluoroCombinedStartString = ActiveCell.Offset(0, 3).Address

While (IsDate(ActiveCell.Value))
    ActiveCell.Offset(1, 0).Activate
Wend

ActiveCell.Offset(-1, 0).Activate
FluoroLabelEndString = ActiveCell.Address
FluoroApptEndString = ActiveCell.Offset(0, 1).Address
FluoroPendEndString = ActiveCell.Offset(0, 2).Address
FluoroCombinedEndString = ActiveCell.Offset(0, 3).Address

FluoroLabelRangeString = FluoroLabelStartString & ":" & FluoroLabelEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + FluoroLabelRangeString
ActiveWorkbook.Names.Add Name:="Fluoro_Label", RefersTo:=namedRangeRef

FluoroApptRangeString = FluoroApptStartString & ":" & FluoroApptEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + FluoroApptRangeString
ActiveWorkbook.Names.Add Name:="Fluoro_Appt", RefersTo:=namedRangeRef

FluoroPendRangeString = FluoroPendStartString & ":" & FluoroPendEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + FluoroPendRangeString
ActiveWorkbook.Names.Add Name:="Fluoro_Pend", RefersTo:=namedRangeRef

FluoroCombinedRangeString = FluoroCombinedStartString & ":" & FluoroCombinedEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + FluoroCombinedRangeString
ActiveWorkbook.Names.Add Name:="Fluoro_Combined", RefersTo:=namedRangeRef




Range("A1").Activate

Cells.Find(What:="CT", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        
ActiveCell.Offset(2, 0).Activate


CTLabelStartString = ActiveCell.Address
CTApptStartString = ActiveCell.Offset(0, 1).Address
CTPendStartString = ActiveCell.Offset(0, 2).Address
CTCombinedStartString = ActiveCell.Offset(0, 3).Address

While (IsDate(ActiveCell.Value))
    ActiveCell.Offset(1, 0).Activate
Wend

ActiveCell.Offset(-1, 0).Activate
CTLabelEndString = ActiveCell.Address
CTApptEndString = ActiveCell.Offset(0, 1).Address
CTPendEndString = ActiveCell.Offset(0, 2).Address
CTCombinedEndString = ActiveCell.Offset(0, 3).Address

CTLabelRangeString = CTLabelStartString & ":" & CTLabelEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + CTLabelRangeString
ActiveWorkbook.Names.Add Name:="CT_Label", RefersTo:=namedRangeRef

CTApptRangeString = CTApptStartString & ":" & CTApptEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + CTApptRangeString
ActiveWorkbook.Names.Add Name:="CT_Appt", RefersTo:=namedRangeRef

CTPendRangeString = CTPendStartString & ":" & CTPendEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + CTPendRangeString
ActiveWorkbook.Names.Add Name:="CT_Pend", RefersTo:=namedRangeRef

CTCombinedRangeString = CTCombinedStartString & ":" & CTCombinedEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + CTCombinedRangeString
ActiveWorkbook.Names.Add Name:="CT_Combined", RefersTo:=namedRangeRef




Range("A1").Activate

Cells.Find(What:="Inter", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
        
ActiveCell.Offset(2, 0).Activate


InterLabelStartString = ActiveCell.Address
InterApptStartString = ActiveCell.Offset(0, 1).Address
InterPendStartString = ActiveCell.Offset(0, 2).Address
InterCombinedStartString = ActiveCell.Offset(0, 3).Address

While (IsDate(ActiveCell.Value))
    ActiveCell.Offset(1, 0).Activate
Wend

ActiveCell.Offset(-1, 0).Activate
InterLabelEndString = ActiveCell.Address
InterApptEndString = ActiveCell.Offset(0, 1).Address
InterPendEndString = ActiveCell.Offset(0, 2).Address
InterCombinedEndString = ActiveCell.Offset(0, 3).Address

InterLabelRangeString = InterLabelStartString & ":" & InterLabelEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + InterLabelRangeString
ActiveWorkbook.Names.Add Name:="Inter_Label", RefersTo:=namedRangeRef

InterApptRangeString = InterApptStartString & ":" & InterApptEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + InterApptRangeString
ActiveWorkbook.Names.Add Name:="Inter_Appt", RefersTo:=namedRangeRef

InterPendRangeString = InterPendStartString & ":" & InterPendEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + InterPendRangeString
ActiveWorkbook.Names.Add Name:="Inter_Pend", RefersTo:=namedRangeRef

InterCombinedRangeString = InterCombinedStartString & ":" & InterCombinedEndString
namedRangeRef = "=" + "'" + "Weekly Outstanding by mod" + "'" + "!" + InterCombinedRangeString
ActiveWorkbook.Names.Add Name:="Inter_Combined", RefersTo:=namedRangeRef



End Sub

