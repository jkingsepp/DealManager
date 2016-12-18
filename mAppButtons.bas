Attribute VB_Name = "mAppButtons"
Option Explicit
'Initialize Module Level Variables
Dim sh As Worksheet
Dim WSName As String
Dim DVRange As String
Dim KDICISheet As String
Dim MsgBoxQues As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Sub GoToSettings()
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Select Setting Worksheet
 Sheets("Settings").Select
   
 'Select Cell D1 In Worksheet
 Range("D1").Select
End Sub
Sub CopyDealData()
 'Initialize Variables
 Dim DealFile As String
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim connStr As String
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Display Standard Open File Dialog Box
 DealFile = Application.GetOpenFilename("Excel Files (*.xls), *.xls", , "Select Deal File To Copy Data From")
 If DealFile = "False" Then
  'No File Selected Message
  Beep
  MsgBoxQues = MsgBox("No file was selected.  Try again?", vbYesNo + vbCritical, gsAPP_NAME)
   
  'Repeat Procedure On "Yes" Answer, Exit Sub On "No"
  If MsgBoxQues = vbYes Then
   Call CopyDealData
  End If
 Else
  'Set Connection String
  connStr = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & _
  DealFile & ";DefaultDir=" & CurDir & ";"
 
  'Create & Open New Connection
  Set adoConn = New ADODB.Connection
  adoConn.Open connStr
   
  'Create & Open New Recordset
  Set adoRS = adoConn.OpenSchema(adSchemaTables)
 
  'Clear WSNames Listbox
  fWSNames.ComboBoxWSName.Clear
  
  'Check To See If Required Worksheet Exists
  Do Until adoRS.EOF
   If Right(Replace(adoRS.Fields("Table_Name"), "'", ""), 1) = "$" Then
    'Populate WSNames Listbox
    fWSNames.ComboBoxWSName.AddItem Replace(Replace(adoRS.Fields("Table_Name"), "'", ""), "$", "")
   End If
  
   'Move To Next Worksheet
   adoRS.MoveNext
  Loop
 
  'Close Recordset & Connection
  adoRS.Close
  adoConn.Close
  
  'Set Top & Left Position & Show UserForm
  fWSNames.ComboBoxWSName.ListRows = fWSNames.ComboBoxWSName.ListCount
  fWSNames.TextBoxWSFileName.Value = DealFile
  fWSNames.Top = Application.Height / 4
  fWSNames.Left = (Application.Width - fWSNames.Width) / 2
  fWSNames.Show
 End If

 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("CopyDealData", Err.Number, Err.Description)
End Sub
Sub FormatWorksheets()
 'Initialize Variable
 Dim Lastcol As String
 Dim i As Integer
 
 ' 20090501 this procedure sets up the data so that it can move correctly into the Report worksheet
 ' this procedure will check formating conditions and also move the appropriate concentration information
 ' the appropriate places.
 
 ' 20100505 updated procedure to add "UpdateFox" routine that aggregates the two fox relationships into one.
 
 ' 20120201 updated procedure for PNC Deal so that we insert the top 10 relationships onto the Servicer Report,
 '  not just the top rated relationships.
 
 '20140630 adding a step to include number of sales per month
 ' NOTE: may need to update again if there could be more than 2 sales per month
 
 
 ' run the poolID list procedure to update list of most recent pool ids
 Call FindTopPoolIDs
 
 ' reset the formula for current month sales
 ThisWorkbook.Sheets("Report").Activate
 Range("NetValueSold").Select
 Selection.Formula = "=DSUM(AssetData," & Chr(34) & "SecuritizedValue" & Chr(34) & ", Settings!b25:b26)"
 ' update Loan Amount formula for 16 c
 ThisWorkbook.Sheets("settings").Activate
 ThisWorkbook.Sheets("Settings").Range("NewExpressSecValue").Select
 Selection.Formula = "=DSUM(AssetData," & Chr(34) & "SecuritizedValue" & Chr(34) & ", Settings!b25:c26)"
  ThisWorkbook.Sheets("Report").Activate
 '
 'Confirm Format Process
 Beep
' MsgBoxQues = MsgBox("Are you processing" & vbCr & _
'"a mid-month sale?", vbYesNo + vbExclamation, gsAPP_NAME)
  
 'Exit Sub On "No"
' If MsgBoxQues = vbYes Then
  i = InputBox("How Many Pools Should be Included in this Report?", gsAPP_NAME, 2)
    ' allow user to enter number of pools and capture
    ' i = reply
    
    ' j = new row number - started b25 to b25+i
   j = i + 25
   ThisWorkbook.Sheets("Report").Range("NetValueSold").Select
 Selection.Formula = "=DSUM(AssetData," & Chr(34) & "SecuritizedValue" & Chr(34) & ", Settings!b25:b" & j & ")"
  ' update Loan Amount formula for 16 c
  ThisWorkbook.Sheets("Settings").Activate
 ThisWorkbook.Sheets("Settings").Range("NewExpressSecValue").Select
 Selection.Formula = "=DSUM(AssetData," & Chr(34) & "SecuritizedValue" & Chr(34) & ", Settings!b25:c" & j & ")"
  ThisWorkbook.Sheets("Report").Activate

 
  'Exit Sub
 'End If
 
 'Select First Cell in Report Worksheet
' Call ReportFirstCell
  
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 ' update Fox relationship (20100504)
 Call UpdateFox
 
 ' Check to make sure there are no existing Concentrations items.
 '  If so, delete the rows so that the report can run from the start
' Call mDeleteConcRows
 
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Format Certain Worksheets
' Call FormatWorksheetsCheck("Inputs")
' Call FormatWorksheetsCheck("Data")
' Call FormatWorksheetsCheck("Capital")
' Call FormatWorksheetsCheck("Closed End")
' Call FormatWorksheetsCheck("Fixed")
' Call FormatWorksheetsCheck("Government")
' Call FormatWorksheetsCheck("Obligors")
'' Call FormatWorksheetsCheck("Sales")
' Call FormatWorksheetsCheck("Rated")
' Call FormatWorksheetsCheck("NonRated")
 
 'Determine Last Column - Data Worksheet
 WSName = "Data"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(1, i).Address
 
 'Select All Rows of Obligors Worksheet
 WSName = "Obligors"
 Sheets(WSName).Select
 j = GetLastRow(1, WSName)
 
 'Set Concentration Percent by dividing the bookvalue for each Parent by the SecValue of the Deal
 Range("E1").Value = "ConcPct"
 For i = 2 To j
  Range("E" & i).Formula = "=D" & i & "/Data!" & _
  Cells(2, Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Sheets("Data").Range("A1:" & Lastcol), 0)).Address
 
 Next i
 
 'Reformat Worksheet
' Call FormatWorksheetsCheck(WSName)
 
 'Copy Entire Selected Worksheet
 'Cells.Select
 'Selection.Copy
 
 'Select Rated Worksheet
 'WSName = "Rated"
 'Sheets(WSName).Select
 
 'Paste Copied Worksheet
 'Cells.Select
 'ActiveSheet.Paste
 'Application.CutCopyMode = False
 
 'Delete NonRated Rows
 'Application.ScreenUpdating = False
 'j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 'For i = GetLastRow(1, WSName) To 2 Step -1
  'If Cells(i, j).Value = "NR" Then
  ' Rows(i).Delete
  'End If
 'Next i
 
 'Sort Selected Range
 i = GetLastRow(1, WSName)
 Range("A2:f" & i).Select
 i = Application.WorksheetFunction.Match("ParentName", Range("A1:D1"), 0)
 j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 k = Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Range("A1:D1"), 0)
 ' this sort was changed for PNC deal - sort on all obligations, not just rated 20120201
 Selection.sort Key1:=Cells(2, k), Order1:=xlDescending, Key2:=Cells(2, i), Order2:=xlAscending, _
 Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
 DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortTextAsNumbers
 
 'original
 'Selection.Sort Key1:=Cells(2, k), Order1:=xlDescending, Key2:=Cells(2, i), Order2:=xlAscending, Key3:=Cells(2, j), _
 'Order3:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
 'DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortTextAsNumbers
 
 'Reformat Worksheet
 'Call FormatWorksheetsCheck(WSName)
 
 'Select Obligors Worksheet
 Sheets("Obligors").Select
 
 'Copy Entire Selected Wowksheet
 'Cells.Select
 'Selection.Copy
 
 'Select NonRated Worksheet
' WSName = "NonRated"
 'Sheets(WSName).Select
 
 'Paste Copied Worksheet
 'Cells.Select
 'ActiveSheet.Paste
 'Application.CutCopyMode = False
 
 
 ' no longer need to delete the nonrated, b/c we are bringing top 10 only for pnc 20120201
 'Delete Rated Rows
j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 For i = GetLastRow(1, WSName) To 2 Step -1
'  If Cells(i, j).Value <> "NR" Then
'   Rows(i).Delete
'  End If
  ' check to see if concentration is less than 3%
'  If Cells(i, j + 2).Value < 0.03 Then
'    Rows(i).Delete
'   End If
   ' delete Navy
   If Cells(i, j - 2).Value = "81020" Then
        Rows(i).Delete
    End If
 Next i
 
 'Sort Selected Range
 'i = GetLastRow(1, WSName)
 'Range("A2:D" & i).Select
 'i = Application.WorksheetFunction.Match("ParentName", Range("A1:D1"), 0)
 'j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 'k = Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Range("A1:D1"), 0)
 'Selection.Sort Key1:=Cells(2, k), Order1:=xlDescending, Key2:=Cells(2, i), Order2:=xlAscending, Key3:=Cells(2, j), _
 'Order3:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
 'DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortTextAsNumbers
 
 'Reformat Worksheet NOT NEEDED AGAIN
 'Call FormatWorksheetsCheck(WSName) '
 
 'Select Obligors Worksheet
' Sheets("Obligors").Select
' Range("G1").Select
 
 'Insert Worksheet Data Links
 Call WSLinkInserts
 
 'Insert Delinquencies Data From Access File
' Call InsertDelinquenciesData
 
 'Insert ChargeOffs and Recoveries Data From DealManager Database
 'Call InsertChargeOffData
 
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
 
 
 'Completion Message
 If ActiveWorkbook.Sheets("Report").Range("j90") > 0 Then
    MsgBox "Formatting complete, but there more be more than 10 Obligors with 3% Concentration!" & _
        vbCrLf & vbCrLf & "Check the Obligors worksheet and manually add in additional rows!", vbCritical, gsAPP_NAME
 Else
    MsgBox "Formatting of the Custom Worksheets Has Been Completed.", vbExclamation, gsAPP_NAME
End If

 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("FormatWorksheets", Err.Number, Err.Description)
End Sub
Sub FormatWorksheetsCheck(AddSheetName As String)
 'Set Up New Worksheet, If Necessary
 For Each sh In Sheets
  If sh.Name = AddSheetName Then
   'Exit Loop
   Exit For
  ElseIf sh.Name = Sheets(Sheets.Count).Name Then
   'Add New Worksheet After Last Existing Worksheet
   Worksheets.Add(After:=Sheets(Sheets.Count)).Name = AddSheetName
   
   'Set Worksheet Color To Light Gray
   Cells.Interior.ColorIndex = 15
   
   'Select First Cell In Worksheet
   Range("A1").Select
  End If
 Next sh

 'Format Worksheets
 i = Sheets(AddSheetName).Range("A1").End(xlToRight).Column
 Select Case AddSheetName
  Case "Inputs": Call WSFormat(AddSheetName, IIf(i = 234, 3, i), GetLastRow(1, AddSheetName))
  Case "Data": Call WSFormat(AddSheetName, IIf(i = 234, 3, i), GetLastRow(1, AddSheetName))
  Case "Capital": Call WSFormat(AddSheetName, IIf(i = 234, 3, i), GetLastRow(1, AddSheetName))
  Case "Closed End": Call WSFormat(AddSheetName, IIf(i = 234, 7, i), GetLastRow(1, AddSheetName))
  Case "Fixed": Call WSFormat(AddSheetName, IIf(i = 234, 7, i), GetLastRow(1, AddSheetName))
  Case "Government": Call WSFormat(AddSheetName, IIf(i = 234, 7, i), GetLastRow(1, AddSheetName))
  Case "Obligors": Call WSFormat(AddSheetName, IIf(i = 234, 4, i), GetLastRow(1, AddSheetName))
  Case "Sales": Call WSFormat(AddSheetName, IIf(i = 234, 9, i), GetLastRow(1, AddSheetName))
  Case "Rated": Call WSFormat(AddSheetName, IIf(i = 234, 4, i), GetLastRow(1, AddSheetName))
  Case "NonRated": Call WSFormat(AddSheetName, IIf(i = 234, 4, i), GetLastRow(1, AddSheetName))
 End Select
End Sub
Sub ConfigureDealTests() 'This Procedure Calls the TestData UserForm For Deal Test Configuration.
 If fTestData.Visible = False Then
  'Missing Deal Data Message
  If ThisWorkbook.Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row = 1 Then
   MsgBox "No deal data has yet been entered in the 'Report' worksheet." & vbCr & vbCr & _
   "Such data is required before the deal tests can be configured.", vbCritical, gsAPP_NAME
 
   'Select First Cell in Report Worksheet
   Call ReportFirstCell
  
   'Exit Procedure
   Exit Sub
  ElseIf GetTNColNum = 234 Then
   'Insert Deal Test Columns
   Call InsertDealTestColumns
  End If
  
  'Select Report Worksheet, If Inactive
  If ActiveSheet.Name <> "Report" Then
   'Select First Cell in Report Worksheet
   Call ReportFirstCell
  End If
  
  'Set Screen Display Controls
  Call SetScreenControls
  
  'Set Up TestData UserForm
  Call TestDataFormSetup
 End If
End Sub
Sub ConfigureKDIs() 'This Procedure Calls the KDI UserForm For Mapping KDI Values/Cells
 If fKDI.Visible = False Then
  'Select Report Worksheet, If Inactive
  If ThisWorkbook.ActiveSheet.Name <> "Report" Then
   'Select First Cell in Report Worksheet
   Call ReportFirstCell
  End If
 
  'Set Screen Display Controls
  Call SetScreenControls
 
  'Determine Key Deal Indicators Worksheet
  If fKDI.TextBoxKDICISheet.Value <> "" Then
   KDICISheet = fKDI.TextBoxKDICISheet.Value
  Else
   KDICISheet = GetKDICIWSName
   fKDI.TextBoxKDICISheet.Value = KDICISheet
  End If
 
  If KDICISheet = "" Then
   'Missing KDI-CI Worksheet Message
   MsgBox "No Key Deal Indicators data could be found in this workbook file." & vbCr & vbCr & _
   "Such data is required in the KDI-CI or similarly-named worksheet" & vbCr & _
   "before the KDI data can be configured.", vbCritical, gsAPP_NAME
  Else
   'Set Up KDI UserForm
   Call KDIFormSetup(KDICISheet)
  End If
 End If
End Sub
Sub ConfigureCIs() 'This Procedure Calls the CalcInput UserForm For Calculated Input Configuration.
 If fCalcInput.Visible = False Then
  'Wrong Worksheet Message
  If ThisWorkbook.ActiveSheet.Name <> "Report" Then
   'Select First Cell in Report Worksheet
   Call ReportFirstCell
  End If
 
  'Set Screen Display Controls
  Call SetScreenControls
 
  'Determine Key Deal Indicators Worksheet
  If fCalcInput.TextBoxKDICISheet.Value <> "" Then
   KDICISheet = fCalcInput.TextBoxKDICISheet.Value
  Else
   KDICISheet = GetKDICIWSName
   fCalcInput.TextBoxKDICISheet.Value = KDICISheet
  End If
 
  'Reformat KDI-CI Table, If Necessary
  Call WSFormat(KDICISheet, 9, 2)
 
  'Select Report Worksheet
  Sheets("Report").Select
  
  If KDICISheet = "" Then
   'Missing KDI-CI Worksheet Message
   MsgBox "No Key Deal Indicators data could be found in this workbook file." & vbCr & vbCr & _
   "Such data is required in the KDI-CI or similarly-named worksheet" & vbCr & _
   "before the CI data can be configured.", vbCritical, gsAPP_NAME
  Else
   'Set Up KDI UserForm
   Call CalcInputFormSetup(KDICISheet)
  End If
 End If
End Sub
Sub ReviewFormulas() 'This Procedure Inserts Any Formula Errors In the Error Log Worksheet.
 'Initialize Variables
 Dim ErrorRange As Range
 Dim ErrorCell As Range

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Missing Deal Data Message
 If ThisWorkbook.Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row = 1 Then
  MsgBox "No deal data has yet been entered in the 'Report' worksheet." & vbCr & vbCr & _
  "Please enter this data first, and then define the deal tests.", vbCritical, gsAPP_NAME
 
  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
  'Exit Procedure
  Exit Sub
 ElseIf GetTNColNum = 234 Then
  'Insert Deal Test Columns
  Call InsertDealTestColumns
 End If
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Insert Temporary Formula Error Cell To Avoid MS Error Message
 Sheets("Report").Range("A1").SpecialCells(xlLastCell).Formula = "=A+1"
 
 'Determine Test Name Column Number, If Any
 j = GetTNColNum
 
 'Determine Last Row of Error Log
 k = GetLastRow(1, "Error Log")
 m = k
 
 'Check Report Worksheet For Error Cells & Loop Through If Any Found
 If Sheets("Report").UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors).Count > 1 Then
  'List Formula Errors In Error Log Worksheet
  Set ErrorRange = Sheets("Report").Cells.SpecialCells(xlFormulas, xlErrors)
  For Each ErrorCell In ErrorRange
   If ErrorCell.Address <> Sheets("Report").Range("A1").SpecialCells(xlLastCell).Address Then
    'Increment Error Log Row
    k = k + 1
     
    'Insert New Row, If Necessary
    WSName = "Error Log"
    Call InsertNewWSRow(k, WSName)
    
    'Insert Error Detail In Error Log Worksheet
    Sheets(WSName).Range("A" & k).Value = ErrorCell.Address '(A) Error Number
    If ErrorCell.Column = j + 1 Then
     Sheets(WSName).Range("B" & k).Value = _
     "Error in Test Result formula." '(B) Error Description
    ElseIf ErrorCell.Column = j + 2 Then
     Sheets(WSName).Range("B" & k).Value = _
     "Error in Difference formula." '(B) Error Description
    Else
     Sheets(WSName).Range("B" & k).Value = _
     "Other Formula Error." '(B) Error Description
    End If
    Sheets(WSName).Range("C" & k).Value = "Formula Error" '(C) VBA Procedure Error Occurred In
    Sheets(WSName).Range("D" & k).Value = Format(Now(), "m/d/yyyy hh:mm:ss") '(D) Error Time
    Sheets(WSName).Range("E" & k).Value = ThisWorkbook.Name '(E) Filename
   End If
  Next ErrorCell
 End If
 
 'Delete Temporary Formula Error Cell
 Sheets("Report").Range("A1").SpecialCells(xlLastCell).Formula = ""
 
 'Show Last Row of Error Log Worksheet
 Sheets("Error Log").Select
 Sheets("Error Log").Range("A" & GetLastRow(1, "Error Log")).Select
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 
 'Check How Many New Error Log Rows Were Added
 If k > m Then
  'Errors Message
  MsgBox "It appears that " & k - m & " formula cell" & IIf(k - m > 1, "s", "") & _
  " in the 'Report' worksheet " & IIf(k - m > 1, "have", "has") & " an error.", vbCritical, gsAPP_NAME
 Else
  'No Errors Message
  MsgBox "The 'Report' worksheet does not contain any formula errors.", vbExclamation, gsAPP_NAME
 End If
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ReviewFormulas", Err.Number, Err.Description)
End Sub
Sub ReviewDealViolations() 'This Procedure Inserts Any Failed Deal Tests In the Error Log Worksheet.
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Missing Deal Data Message
 If ThisWorkbook.Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row = 1 Then
  MsgBox "No deal data has yet been entered in the 'Report' worksheet." & vbCr & vbCr & _
  "Please enter this data first, and then define the deal tests.", vbCritical, gsAPP_NAME
 
  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
  'Exit Procedure
  Exit Sub
 ElseIf GetTNColNum = 234 Then
  'Insert Deal Test Columns
  Call InsertDealTestColumns
 End If
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Determine Test Name Column Number, If Any
 j = GetTNColNum
 
 'Determine Last Row of Error Log
 k = GetLastRow(1, "Error Log")
 m = k
 
 'Check Report Worksheet For Failed Deal Test Cells & Loop Through If Any Found
 DVRange = "A2:" & Cells(GetLastRow(j + 1, "Report"), j + 1).Address
 If WorksheetFunction.CountIf(Sheets("Report").Range(DVRange), "Fail") > 0 Then
  'List Failed Deal Tests In Error Log Worksheet
  For i = 2 To GetLastRow(j, "Report")
   If Sheets("Report").Cells(i, j).Value <> "" And Sheets("Report").Cells(i, j + 1).Value = "Fail" Then
    'Increment Error Log Row
    k = k + 1
     
    'Insert New Row, If Necessary
    WSName = "Error Log"
    Call InsertNewWSRow(k, WSName)
    
    'Insert Error Detail In Error Log Worksheet
    Sheets(WSName).Range("A" & k).Value = "Row " & Format(i, "#,##0") '(A) Error Number
    Sheets(WSName).Range("B" & k).Value = "Violation of deal test." '(B) Error Description
    Sheets(WSName).Range("C" & k).Value = "N/A" '(C) VBA Procedure Error Occurred In
    Sheets(WSName).Range("D" & k).Value = Format(Now(), "m/d/yyyy hh:mm:ss") '(D) Error Time
    Sheets(WSName).Range("E" & k).Value = ThisWorkbook.Name '(E) Filename
   End If
  Next i
 End If
 
 'Show Last Row of Error Log Worksheet
 Sheets("Error Log").Select
 Sheets("Error Log").Range("A" & GetLastRow(1, "Error Log")).Select
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 
 'Check How Many New Error Log Rows Were Added
 If k > m Then
  'Errors Message
  MsgBox "It appears that " & k - m & " deal test" & IIf(k - m > 1, "s", "") & _
  " in the 'Report' worksheet " & IIf(k - m > 1, "have", "has") & " been violated.", vbCritical, gsAPP_NAME
 Else
  'No Errors Message
  MsgBox "The 'Report' worksheet does not contain any deal test violations.", vbExclamation, gsAPP_NAME
 End If
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ReviewDealViolations", Err.Number, Err.Description)
End Sub
Sub UpdateTestsWS() 'This Procedure Copies All Deal Tests From the Report to the Tests Worksheet
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 If ThisWorkbook.Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row = 1 Then
  'Missing Deal Data Message
  MsgBox "No deal data has yet been entered in the 'Report' worksheet." & vbCr & vbCr & _
  "Please enter this data first, and then define the deal tests.", vbCritical, gsAPP_NAME
 
  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
  'Exit Procedure
  Exit Sub
 ElseIf GetTNColNum = 234 Then
  'Insert Deal Test Columns
  Call InsertDealTestColumns
 End If
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Check For Tests Worksheet
 WSName = "Tests"
 For Each sh In Sheets
  If sh.Name = WSName Then
   'Exit Loop
   Exit For
  ElseIf sh.Name = Sheets(Sheets.Count).Name Then
   'Add New Worksheet After Last Existing Worksheet
   Worksheets.Add(After:=Sheets(Sheets.Count)).Name = WSName
   
   'Set Worksheet Color To Light Gray
   Cells.Interior.ColorIndex = 15
   
   'Select First Cell In Worksheet
   Range("A1").Select
  
   'Set Up Settings Worksheet
   Call TestsWSSetup(WSName)
  End If
 Next sh
 
 'Clear Tests Worksheet Table
 Sheets(WSName).Rows("1:" & Sheets(WSName).Range("A1").SpecialCells(xlLastCell).Row).Delete
 Call TestsWSSetup(WSName)
 
 'Determine Test Name Column Number, If Any
 j = GetTNColNum
 
 'Copy Deal Tests To Tests Worksheet
 k = 1
 For i = 2 To GetLastRow(j, "Report")
  If Sheets("Report").Cells(i, j).Value <> "" Then
   'Increment Tests Row
   k = k + 1
     
   'Insert New Row, If Necessary
   Call InsertNewWSRow(k, WSName)
    
   'Insert Deal Test Detail In Tests Worksheet
   Sheets(WSName).Range("A" & k).Value = Sheets("Report").Cells(i, j) '(A) Test Name
   Sheets(WSName).Range("B" & k).Value = Sheets("Report").Cells(i, j + 1) '(B) Test Result
   If IsError(Sheets("Report").Cells(i, j + 2)) = True Then
    Sheets(WSName).Range("C" & k).Value = 0 '(C) Difference
   ElseIf Sheets("Report").Cells(i, j + 2) = "" Then
    Sheets(WSName).Range("C" & k).Value = 0 '(C) Difference
   Else
    Sheets(WSName).Range("C" & k).Value = Sheets("Report").Cells(i, j + 2) '(C) Difference
   End If
   Sheets(WSName).Range("D" & k).Value = Sheets("Report").Cells(i, j + 3) '(D) Test Type
   Sheets(WSName).Range("E" & k).Value = Sheets("Report").Cells(i, j).Address '(E) Cell Reference
  
   'Highlight Deal Test Row
   If Sheets(WSName).Range("B" & k).Value = "Pass" Then
    'Pass Test Result
    With Sheets("Report").Range("A" & i & ":" & Sheets("Report").Cells(i, j + 3).Address).Interior
     .ColorIndex = 4
     .Pattern = xlSolid
    End With
   Else
    'Fail Test Result
    With Sheets("Report").Range("A" & i & ":" & Sheets("Report").Cells(i, j + 3).Address)
     .Interior.ColorIndex = 3
     .Interior.Pattern = xlSolid
     .Font.ColorIndex = 2
    End With
   End If
  Else
   'Remove Highlight From Worksheet Row
   If Sheets("Report").Range(Sheets("Report").Cells(i, j).Address).Interior.ColorIndex = 3 Or _
   Sheets("Report").Range(Sheets("Report").Cells(i, j).Address).Interior.ColorIndex = 4 Then
    With Sheets("Report").Range("A" & i & ":" & Sheets("Report").Cells(i, j + 3).Address)
     .Interior.ColorIndex = xlNone
     .Font.ColorIndex = 0
    End With
   End If
  End If
 Next i
 
 'Complete Procedure Unless Called From Other CommandBar Button
 If Right(CommandBars.ActionControl.Caption, 15) = "Tests Worksheet" Then
  'Show First Row of Tests Worksheet
  Sheets(WSName).Select
  Sheets(WSName).Range("G1").Select
 
  'Clear Status Message & Turn On Screen Updating
  Call ClearScreenControls
  
  'Check How Many Deal Tests Rows Were Copied From the Report Worksheet
  If k > 1 Then
   'Count of Deal Tests Message
   MsgBox k - 1 & " deal tests have been copied from the 'Report' to the 'Tests' worksheet.", _
   vbExclamation, gsAPP_NAME
  Else
   'No Deal Tests Message
   MsgBox "The 'Report' worksheet does not contain any deal tests.", vbExclamation, gsAPP_NAME
  End If
 End If
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("UpdateTestsWS", Err.Number, Err.Description)
End Sub
Sub ExportDataToDM() 'This Procedure Exports the Deal Test, KDI & CI Data to the DealManager Database
 'Initialize Variables
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMCreationDate As Date
 Dim DMEffectiveDate As Date
 Dim connStr As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim WhereClause As String
 Dim SQLStmt As String
 Dim fso As Object
 Dim DealID As Integer
 Dim DMid As Integer
 Dim sAnswer As String
 Dim sValue As String
 Dim vSame As Variant


' procedure uses globabl variable vDMValue, which is set by the "checkforDMValue" routine
' for situations where updates are required to existing dmvalue.

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 

  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Determine Data Worksheet Name
 WSName = "KDI-CI"
 ' Set the table
  DBTable = "DealMetricValues"
 'Set Data Creation & Effective Date Values (Effective Date Is Last Inputs Effective Date)
 DMCreationDate = Now
 ' Get DealID
 DealID = ThisWorkbook.Sheets("Settings").Range("Dealid").Value
  ' determine connection string
connStr = GetConnectionString


' check the number of inputs
  i = Sheets("Inputs").Range("A1").End(xlToRight).Column
  i = Application.WorksheetFunction.Match("Effective Date", Sheets("Inputs").Range("A1:" & Sheets("Inputs").Cells(1, i).Address), 0)
' set the most recent effective date
  DMEffectiveDate = Sheets("Inputs").Cells(2, i)
 
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Determine Last Row of Worksheet Data
 j = GetLastRow(1, WSName)
 
 For i = 2 To j
          'Data Export Status Message
          Application.StatusBar = "Exporting " & WSName & " data from row " & i
          
          'Check to make sure a value exists
          If Sheets(WSName).Range("F" & i).Value = "" Then
            GoTo Skip
          End If
        
        ' set the dmid
        DMid = Sheets(WSName).Range("A" & i).Value
        ' call the procedure that will determine whether a value exists
        ' for each item.
        sValue = Sheets(WSName).Range("f" & i).Value
        sAnswer = CheckforDealMetricValue(DealID, DMEffectiveDate, DMid)
        ' remove the $ signs, if any
        
        sAnswer = Replace(sAnswer, "$", "", 1)
        sValue = Replace(sValue, "$", "", 1)
        
        vSame = StrComp(sAnswer, sValue, vbBinaryCompare)
        If vSame = "0" Then
            ' The values are equal, only update the CreateDt"
            '
                    
                    InsertClause = "UPDATE " & DBTable & " SET "
                    InsertClause = InsertClause & " CreateDt = '" & DMCreationDate & "', "
                    InsertClause = InsertClause & " Comments = 'Updated CreateDt via DealManager'"
                    
                    WhereClause = " WHERE DealMetricValueID = " & vDMValueID
                    
                    SQLStmt = InsertClause & WhereClause & ";"
                     'Create & Open New Recordset
                  Set adoRS = New ADODB.Recordset
                  adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
            GoTo Skip
        End If
        
        Select Case sAnswer
            Case "InsertNew"
                ' No effective date - insert new"
               'Set Basic INSERT INTO Clause
                  InsertClause = "INSERT INTO " & DBTable & "("
                       
                  'Set Values Portion of INSERT INTO Clause
                  
                   InsertClause = InsertClause & "DealMetricID, DealID, CreateDt, Comments, EffectiveDt, Value, Source) VALUES ("
                   InsertClause = InsertClause & "'" & Sheets(WSName).Range("A" & i) & "', '" & DealID & "', '" & DMCreationDate & "', " & "'Inserted via DealManager'" & _
                   ", '" & DMEffectiveDate & "', '" & Sheets(WSName).Range("F" & i) & "', '" & Sheets(WSName).Range("G" & i) & "')"
                  
                  
                  'SQL Query String
                  SQLStmt = InsertClause & ";"
        
                  'Create & Open New Recordset
                '''  Set adoRS = New ADODB.Recordset
                ' adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
             
            Case "UpdateEmpty"
                    ' the value is null, update all"
                    
                    InsertClause = "UPDATE " & DBTable & " SET "
                    InsertClause = InsertClause & " CreateDt = '" & DMCreationDate & "', "
                    InsertClause = InsertClause & " EffectiveDt = '" & DMEffectiveDate & "', "
                    InsertClause = InsertClause & " Value = '" & Sheets(WSName).Range("F" & i) & "', "
                    InsertClause = InsertClause & " Source = '" & Sheets(WSName).Range("g" & i) & "', "
                    InsertClause = InsertClause & " Comments = 'Updated Dates and Value via DealManager'"
                    
                    WhereClause = " WHERE DealMetricValueID = " & vDMValueID
                    
                    SQLStmt = InsertClause & WhereClause & ";"
                    'Debug.Print SQLStmt
             Case Else
                    ' The values are not equal, insert a row in DMHist table and update record"
                  'Set Basic INSERT INTO Clause
                   ' get teh values incase we need to insert into Hist table


                  InsertClause = "INSERT INTO " & gsDMHistTable & "("
                       
                  'Set Values Portion of INSERT INTO Clause
                  
                   InsertClause = InsertClause & "DealMetricID, DealID, Insertdt, CreateDt, Comments, EffectiveDt, Value, Source) VALUES ("
                   InsertClause = InsertClause & "'" & vDMid & "', '" & vDealID & "', '" & Now() & "', '" & vCreateDt & "', '" & vComments & _
                   "', '" & vEffectiveDt & "', '" & vValue & "', '" & vSource & "')"
                  
                  
                  'SQL Query String
                  SQLStmt = InsertClause & ";"
        Debug.Print SQLStmt
          'Create & Open New Recordset
                  Set adoRS = New ADODB.Recordset
                  adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
                  
                ' now run the update to update records in DealMetricValues table
                
               InsertClause = "UPDATE " & DBTable & " SET "
                    InsertClause = InsertClause & " CreateDt = '" & DMCreationDate & "', "
                    InsertClause = InsertClause & " EffectiveDt = '" & DMEffectiveDate & "', "
                    InsertClause = InsertClause & " Value = '" & Sheets(WSName).Range("F" & i) & "', "
                    InsertClause = InsertClause & " Source = '" & Sheets(WSName).Range("g" & i) & "', "
                    InsertClause = InsertClause & " Comments = 'Updated via DealManager'"
                    
                    WhereClause = " WHERE DealMetricValueID = " & vDMValueID
                    
                    SQLStmt = InsertClause & WhereClause & ";"
                    'Debug.Print SQLStmt
                    
        End Select

           'Create & Open New Recordset
                  Set adoRS = New ADODB.Recordset
                  adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
        
Skip:
 Next i
 
 'Close Connection
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 

 'Completion Message
 
  MsgBox "Export of the KDI & CI Data Has Been Completed.", vbExclamation, gsAPP_NAME

 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ExportDataToDM", Err.Number, Err.Description)
End Sub

Sub ExportReport() 'This Procedure Creates the Trial & Final Reports
 'Initialize Variables
 Dim ReportType As String
 Dim ReportName As String
 Dim MsgBoxQues As String
 Dim ReportZoom As Integer
 Dim FileName As String
 Dim fso As Object
 Dim i As Integer
 Dim ws As Worksheet
 Dim wbTemplate As Workbook
 Dim wbReport As Workbook
 Dim nm As Name
 
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 ThisWorkbook.Sheets("Report").Activate
 
 Set wbTemplate = ActiveWorkbook
 Application.ScreenUpdating = False
 
 
 'Missing Deal Data Message

 'Set Report Type To Create
 ReportType = "Final Report"
 'Right(CommandBars.ActionControl.Caption, 12)
 
i = MsgBox("If this is the final report, remember to export data to DealManager at the prompt!", vbInformation, gsAPP_NAME)
 
 'Delete Temporary Formula Error Cell
 Sheets("Report").Range("A1").SpecialCells(xlLastCell).Formula = ""
 
 'Determine Test Name Column Number, If Any
 'j = GetTNColNum
 

 'Set Screen Display Controls
 Call SetScreenControls
 
 'Upload KDI & CI Data to DealManager
 i = MsgBox("Would you like to export to DealManager before creating the Report?", vbYesNo, gsAPP_NAME)
 If i = vbYes Then
     Call ExportDataToDM
     
Else
'If i = vbNo Then
 '   i = MsgBox("You MUST export values to DealManager when generating the Final Report - Would you like to export data now?", vbYesNo, gsAPP_NAME)
 End If
 
 
 'Set Report Filename
 FileName = CStr(Range("DealName").Value & "-" & _
 ReportType & "-" & Format(ThisWorkbook.Sheets("Report").Range("DealDate").Value, "yyyymmdd")) & ".xls"
    ' Filename = ThisWorkbook.path & "\" & Range("DealName").Value & "-" & _
    '    ReportType & "-" & Format(ThisWorkbook.Sheets("Report").Range("DealDate").Value, "yyyymmdd") & ".xls"
  
 'prompt user for the location
  ReportName = Application.GetSaveAsFilename(FileName) ', "Excel files (*.xl*),*.xls")
  
    If ReportName = "False" Then
        MsgBox "Please Enter a File Name!", vbInformation, gsAPP_NAME
        Exit Sub
    End If
    'Application.ScreenUpdating = True
   ' ReportName = ReportName
    'MsgBox ReportName
   ' ActiveWorkbook.SaveAs FileName:=FileName
 'Check For Existing File of Same Name & Delete If Exists
 Set fso = CreateObject("Scripting.FileSystemObject")
 If fso.FileExists(ReportName) = True Then
  fso.DeleteFile ReportName
 End If
  
 'Set Current Worksheet Zoom Value
 ReportZoom = ActiveWindow.Zoom
  
 'Create New Workbook
' Workbooks.Add
' Set wbReport = ActiveWorkbook
 'Set Worksheet Color To Light Gray
' Cells.Interior.ColorIndex = 15
  ' Debug.Print ReportName
 'Save New Workbook Under New Name
 'ActiveWorkbook.SaveAs ReportName
 
 'Reactivate This Workbook & Copy Report Worksheet
 wbTemplate.Activate
 
 Sheets(Array("Report", "Schedule 1", "Schedule 2", "Hedge", "Hedge Schedule")).Select
 Sheets(Array("Report", "Schedule 1", "Schedule 2", "Hedge", "Hedge Schedule")).Copy
 'Cells.Select
' Selection.Copy
 
 'Disable Alerts
 Application.DisplayAlerts = False
 Set wbReport = ActiveWorkbook
  'Set Worksheet Color To Light Gray
 'Cells.Interior.ColorIndex = 15
  ' Debug.Print ReportName
 'Save New Workbook Under New Name
 ActiveWorkbook.SaveAs ReportName
 'Reactivate New Workbook & Paste Report Worksheet
 wbReport.Activate

 'Cells.Select
 'ActiveSheet.Paste
 'Application.CutCopyMode = False
 
 'Enable Alerts
 Application.DisplayAlerts = True
 
 'Rename Worksheet To "Department Report"
 ActiveSheet.Name = "Report"
 
 'Convert All Formula & Linked Cells To Fixed Values
 If ReportType = "Final Report" Then
  Cells.Select
  Selection.Copy
  Selection.PasteSpecial Paste:=xlPasteValues
  
 Range("L:AF").EntireColumn.Delete
 
  'Delete Deal Test Columns, If They Exist
 ' If GetTNColNum < 234 Then
 '  Range(Cells(1, GetTNColNum).Address & ":" & _
 'Cells(1, Cells(1, GetTNColNum).SpecialCells(xlLastCell).Row).Address).EntireColumn.Delete
 ' End If
 End If
 
 'Rename Worksheet
 ActiveSheet.Name = ReportType
 
 'Select First Cell In Worksheet
 Range("A1").Select
  
  ' remove the named range
    On Error Resume Next
    For Each nm In wbReport.Names
        nm.Delete
    Next
    On Error GoTo 0
 
  
 'Set New Worksheet Zoom Value
 ActiveWindow.Zoom = ReportZoom
 'Save & Exit New Workbook
 'ActiveWorkbook.Save
 wbReport.Save
 ActiveWorkbook.Saved = True
 'ActiveWorkbook.Close

' activate the Template workbook
wbTemplate.Activate
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
  
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls

 'Completion Message
 MsgBox "The " & ReportType & " Has Been Created and Saved." & vbCrLf & vbCrLf & _
    "The file is located at " & ReportName & ".", vbExclamation, gsAPP_NAME
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ExportReport", Err.Number, Err.Description)
End Sub
Sub HelpScreen()
 'Insert Help Screen Content
 Call HelpScreenContents
 
 'Set Top & Left Position & Show UserForm
 fHelp.Top = 2
 fHelp.Left = (Application.Width - fHelp.Width) / 2
 fHelp.Show
End Sub
Sub AboutNBS()
 'Set Version Number
 fAboutNBS.LabelVersion.Caption = "(" & gsVERSION & ")"

 'Set Top & Left Position & Show UserForm
 fAboutNBS.Top = Application.Height / 4
 fAboutNBS.Left = (Application.Width - fAboutNBS.Width) / 2
 fAboutNBS.Show
End Sub
Sub GoToSheet() 'This Procedure Activates the Selected Worksheet
 'Initialize Variable
 Dim i As Integer
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Determine Selected Worksheet
 i = CommandBars("DealManager").Controls("Go To Sheet").ListIndex
 
 'Go To Selected Worksheet
 Sheets(CommandBars("DealManager").Controls("Go To Sheet").List(i)).Select
   
 'Select First Cell In Worksheet
 Select Case ActiveSheet.Name
  Case "Settings": Range("D1").Select
  Case "Error Log", "Tests": Range("G1").Select
  Case Else: Range("A1").Select
 End Select
 ActiveWindow.ScrollRow = 1
 ActiveWindow.ScrollColumn = 1
End Sub

Sub oldExportDataToDM() 'This Procedure Exports the Deal Test, KDI & CI Data to the DealManager Database
' replaced to handle new db structure with 2.5.19 20111210
 'Initialize Variables
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim DMCreationDate As Date
 Dim DMEffectiveDate As Date
 Dim connStr As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim SQLStmt As String
 Dim fso As Object
 Dim sAnswer As String

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 

  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
 'Set Screen Display Controls
 Call SetScreenControls
 
 'First update the Test Worksheet
 'If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
 ' Call UpdateTestsWS
 'End If
 
 'Determine Data Worksheet Name
 WSName = "KDI-CI"
' For Each sh In Sheets
'  If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
'   If sh.Range("A1") = "Test Name" And Left(sh.Range("B1"), 4) = "Test" And _
'   sh.Range("C1") = "Difference" And sh.Range("D1") = "Test Type" Then
    'Assign Tests Sheet Variable Value
'    WSName = sh.Name
    
    'Exit Loop
'    Exit For
'   End If
'  Else
'   If sh.Range("A1") = "ID" And sh.Range("B1") = "Source" And _
'   sh.Range("C1") = "Name" And sh.Range("D1") = "Type" Then
    'Assign KDI-CI Sheet Variable Value
'    WSName = sh.Name
     
    'Exit Loop
'    Exit For
'   End If
'  End If
' Next sh
'
 
 'Set DealManager .ini File Name, Server Name & Database Name
 Set fso = CreateObject("Scripting.FileSystemObject")
 DMIniFileName = "C:\Documents and Settings\" & UserNameWindows & "\Application Data\NorthBound Solutions\DealManager Suite\dealmanager.ini"
 If fso.FileExists(DMIniFileName) = False Then
  DMIniFileName = "c:\Program Files\NorthBound Solutions\DealManager Suite\dealmanager.ini"

  ' Left(ThisWorkbook.path, 3) & _
 '"c:\Program Files\NorthBound Solutions\DealManager Suite 2.3.12\dealmanager.ini"
  If fso.FileExists(DMIniFileName) = False Then
  MsgBox DMIniFileName
   'No DealManager.ini File Message
   MsgBox "No dealmanager.ini file could be found either in this file's folder area" & vbCr & _
   "or with the path name " & DMIniFileName & "." & vbCr & vbCr & _
   "Check to ensure that this file exists in one of those two locations.  If it doesn't," & vbCr & _
   "contact NorthBound Solutions, Inc.", vbCritical, gsAPP_NAME
  End If
 End If
 DMServerName = ExtractServerName(GetPrivateProfileString32(DMIniFileName, "Database", "ConnectionString"))
 If DMServerName = "localhost" Then
  DMServerName = "(local)"
 End If
 DMDBName = ExtractDBName(GetPrivateProfileString32(DMIniFileName, "Database", "ConnectionString"))
 
 'Set Connection String
 UserName = ExtractUserName(GetPrivateProfileString32(DMIniFileName, "Database", "ConnectionString"))
 PwdName = ExtractPassword(GetPrivateProfileString32(DMIniFileName, "Database", "ConnectionString"))
 If UserName <> "" Then
  connStr = "DRIVER={SQL SERVER};Server=" & DMServerName & ";Database=" & DMDBName & _
  ";Uid=" & UserName & ";Pwd=" & PwdName & ";ReadOnly=False;"
 Else
  connStr = "DRIVER={SQL SERVER};Server=" & DMServerName & ";Database=" & DMDBName & ";ReadOnly=False;"
 End If
 
 'Set DB Table Name
' If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
'  DBTable = "DealTests"
' Else
  DBTable = "DealMetricValues"
' End If
 
 'Set Data Creation & Effective Date Values (Effective Date Is Last Inputs Effective Date)
 DMCreationDate = Now
 
  i = Sheets("Inputs").Range("A1").End(xlToRight).Column
  i = Application.WorksheetFunction.Match("Effective Date", Sheets("Inputs").Range("A1:" & Sheets("Inputs").Cells(1, i).Address), 0)
  DMEffectiveDate = Sheets("Inputs").Cells(2, i)
 
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Determine Last Row of Worksheet Data
 j = GetLastRow(1, WSName)
 
 For i = 2 To j
  'Data Export Status Message
  Application.StatusBar = "Exporting " & WSName & " data from row " & i
  
  'Check to make sure a value exists
  If Sheets(WSName).Range("F" & i).Value = "" Then
    GoTo Skip
  End If


' call the procedure that will determine whether a value exists
' for each item.

sAnswer = CheckforDealMetricValue("1", "11/30/2011", "45")



If sAnswer = "1" Then
    MsgBox "the value is null, update all"
    
Else
    sAnswer = "2"
        If sAnswer = "1" Then
            MsgBox "the values are equal, only update the CreateDt"
        Else
            MsgBox "the values are not equal, insert a row in DMHist table and update record"
            
        End If
End If

  'Set Basic INSERT INTO Clause
  InsertClause = "INSERT INTO " & DBTable & "("
  
 
  'Set Values Portion of INSERT INTO Clause
  
   InsertClause = InsertClause & "DealMetricID, DealID, CreateDt, Comments, EffectiveDt, Value, Source) VALUES ("
   InsertClause = InsertClause & "'" & Sheets(WSName).Range("A" & i) & "', '" & "1" & "', '" & DMCreationDate & "', " & "'Loaded via DealManager'" & _
   ", '" & DMEffectiveDate & "', '" & Sheets(WSName).Range("F" & i) & "', '" & Sheets(WSName).Range("G" & i) & "')"
  
  
  'SQL Query String
  SQLStmt = InsertClause & ";"
 
 
' MsgBox "You are about to update the database with the following information! Continue????" & vbCrLf & vbCrLf & SQLStmt, vbYesNoCancel + vbInformation, gsAPP_NAME
    
 ' Debug.Print SQLStmt
  'Create & Open New Recordset
  Set adoRS = New ADODB.Recordset
  adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
Skip:
 Next i
 
 'Close Connection
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 

 'Completion Message
 
  MsgBox "Export of the KDI & CI Data Has Been Completed.", vbExclamation, gsAPP_NAME

 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ExportDataToDM", Err.Number, Err.Description)
End Sub

Sub FormatWorksheetsOLD()
 'Initialize Variable
 Dim Lastcol As String
 
 ' 20090501 this procedure sets up the data so that it can move correctly into the Report worksheet
 ' this procedure will check formating conditions and also move the appropriate concentration information
 ' the appropriate places.
 
 ' 20100505 updated procedure to add "UpdateFox" routine that aggregates the two fox relationships into one.
 
 'Confirm Format Process
 Beep
 MsgBoxQues = MsgBox("This format process is currently set" & vbCr & _
 "up only for limited Deals.  Continue?", vbYesNo + vbExclamation, gsAPP_NAME)
  
 'Exit Sub On "No"
 If MsgBoxQues = vbNo Then
  Exit Sub
 End If
 
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
  
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 ' update Fox relationship (20100504)
 Call UpdateFox
 
 ' Check to make sure there are no existing Concentrations items.
 '  If so, delete the rows so that the report can run from the start
 Call mDeleteConcRows
 
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Format Certain Worksheets
 Call FormatWorksheetsCheck("Inputs")
 Call FormatWorksheetsCheck("Data")
 Call FormatWorksheetsCheck("Capital")
 Call FormatWorksheetsCheck("Closed End")
 Call FormatWorksheetsCheck("Fixed")
 Call FormatWorksheetsCheck("Government")
 Call FormatWorksheetsCheck("Obligors")
' Call FormatWorksheetsCheck("Sales")
 Call FormatWorksheetsCheck("Rated")
 Call FormatWorksheetsCheck("NonRated")
 
 'Determine Last Column - Data Worksheet
 WSName = "Data"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(1, i).Address
 
 'Select All Rows of Obligors Worksheet
 WSName = "Obligors"
 Sheets(WSName).Select
 j = GetLastRow(1, WSName)
 
 'Set Concentration Percent by dividing the bookvalue for each Parent by the SecValue of the Deal
 Range("E1").Value = "ConcPct"
 For i = 2 To j
  Range("E" & i).Formula = "=D" & i & "/Data!" & _
  Cells(2, Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Sheets("Data").Range("A1:" & Lastcol), 0)).Address
 
 Next i
 
 'Reformat Worksheet
 Call FormatWorksheetsCheck(WSName)
 
 'Copy Entire Selected Worksheet
 Cells.Select
 Selection.Copy
 
 'Select Rated Worksheet
 WSName = "Rated"
 Sheets(WSName).Select
 
 'Paste Copied Worksheet
 Cells.Select
 ActiveSheet.Paste
 Application.CutCopyMode = False
 
 'Delete NonRated Rows
 Application.ScreenUpdating = False
 j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 For i = GetLastRow(1, WSName) To 2 Step -1
  If Cells(i, j).Value = "NR" Then
   Rows(i).Delete
  End If
 Next i
 
 'Sort Selected Range
 i = GetLastRow(1, WSName)
 Range("A2:D" & i).Select
 i = Application.WorksheetFunction.Match("ParentName", Range("A1:D1"), 0)
 j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 k = Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Range("A1:D1"), 0)
 Selection.sort Key1:=Cells(2, k), Order1:=xlDescending, Key2:=Cells(2, i), Order2:=xlAscending, Key3:=Cells(2, j), _
 Order3:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
 DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortTextAsNumbers
 
 'Reformat Worksheet
 'Call FormatWorksheetsCheck(WSName)
 
 'Select Obligors Worksheet
 Sheets("Obligors").Select
 
 'Copy Entire Selected Wowksheet
 Cells.Select
 Selection.Copy
 
 'Select NonRated Worksheet
 WSName = "NonRated"
 Sheets(WSName).Select
 
 'Paste Copied Worksheet
 Cells.Select
 ActiveSheet.Paste
 Application.CutCopyMode = False
 
 'Delete Rated Rows
 j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 For i = GetLastRow(1, WSName) To 2 Step -1
  If Cells(i, j).Value <> "NR" Then
   Rows(i).Delete
  End If
  ' check to see if concentration is less than 3%
  If Cells(i, j + 2).Value < 0.03 Then
    Rows(i).Delete
   End If
   ' delete Navy
   If Cells(i, j - 2).Value = "81020" Then
        Rows(i).Delete
    End If
 Next i
 
 'Sort Selected Range
 i = GetLastRow(1, WSName)
 Range("A2:D" & i).Select
 i = Application.WorksheetFunction.Match("ParentName", Range("A1:D1"), 0)
 j = Application.WorksheetFunction.Match("SP_Rating", Range("A1:D1"), 0)
 k = Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Range("A1:D1"), 0)
 Selection.sort Key1:=Cells(2, k), Order1:=xlDescending, Key2:=Cells(2, i), Order2:=xlAscending, Key3:=Cells(2, j), _
 Order3:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
 DataOption1:=xlSortNormal, DataOption2:=xlSortTextAsNumbers, DataOption3:=xlSortTextAsNumbers
 
 'Reformat Worksheet NOT NEEDED AGAIN
 'Call FormatWorksheetsCheck(WSName) '
 
 'Select Obligors Worksheet
 Sheets("Obligors").Select
 Range("G1").Select
 
 'Insert Worksheet Data Links
 Call WSLinkInserts
 
 'Insert Delinquencies Data From Access File
' Call InsertDelinquenciesData
 
 'Insert ChargeOffs and Recoveries Data From DealManager Database
 'Call InsertChargeOffData
 
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
 
 
 'Completion Message
 MsgBox "Formatting of the Custom Worksheets Has Been Completed.", vbExclamation, gsAPP_NAME

 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("FormatWorksheets", Err.Number, Err.Description)
End Sub

