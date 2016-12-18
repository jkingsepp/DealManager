Attribute VB_Name = "mUserFormSetup"
Option Explicit
'Initialize Module Level Variables
Dim PFString As String
Dim CIRange As String
Dim CICell As String
Dim ListBoxArray()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Sub TestDataFormSetup()
 'Initialize Variable
 Dim OpType As String
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Check To See If UserForm Test Type ComboBox Needs To Be Initialized
 If fTestData.ComboBoxTestType.ListCount = 0 Then
  'Set List of Test Types
  fTestData.ComboBoxTestType.Clear
  fTestData.ComboBoxTestType.AddItem "Trigger"
  fTestData.ComboBoxTestType.AddItem "Balance"
  fTestData.ComboBoxTestType.AddItem "Cash"
  fTestData.ComboBoxTestType.AddItem "Other"
 End If
 
 'Check To See If UserForm Operator ComboBox Needs To Be Initialized
 If fTestData.ComboBoxOperator.ListCount = 0 Then
  'Set List of Test Data Operators
  fTestData.ComboBoxOperator.Clear
  fTestData.ComboBoxOperator.AddItem "equals"
  fTestData.ComboBoxOperator.AddItem "does not equal"
  fTestData.ComboBoxOperator.AddItem "is greater than"
  fTestData.ComboBoxOperator.AddItem "is greater than or equal to"
  fTestData.ComboBoxOperator.AddItem "is less than"
  fTestData.ComboBoxOperator.AddItem "is less than or equal to"
  fTestData.ComboBoxOperator.AddItem "begins with"
  fTestData.ComboBoxOperator.AddItem "does not begin with"
  fTestData.ComboBoxOperator.AddItem "ends with"
  fTestData.ComboBoxOperator.AddItem "does not end with"
  fTestData.ComboBoxOperator.AddItem "contains"
  fTestData.ComboBoxOperator.AddItem "does not contain"
 End If
 
 'Set TestData UserForm Current Test Name, If Any, & Test Name Row
 j = GetTNColNum
 fTestData.TextBoxTestName.Value = Cells(ActiveCell.Row, j).Value
 If ActiveCell.Row < 2 Then
  fTestData.SpinRow.Value = 2
  Cells(2, j).Select
  ActiveWindow.ScrollRow = 2
 Else
  fTestData.SpinRow.Value = ActiveCell.Row
 End If
 fTestData.TextBoxTestRow.Value = Format(fTestData.SpinRow.Value, "#,##0")
 
 'Set TestData UserForm Current Test Type, If Any
 fTestData.ComboBoxTestType.Value = Cells(ActiveCell.Row, j + 3).Value
 
 'Set TestData UserForm Current Operator, If Any
 PFString = Mid(Cells(ActiveCell.Row, j + 1).Formula, 5)
 If PFString <> "" Then
  If InStr(PFString, ">=") > 0 Then
   OpType = ">="
   fTestData.ComboBoxOperator.Value = "is greater than or equal to"
  ElseIf InStr(PFString, "<=") > 0 Then
   OpType = "<="
   fTestData.ComboBoxOperator.Value = "is less than or equal to"
  ElseIf InStr(PFString, "=") > 0 Then
   OpType = "="
   If InStr(PFString, "LEFT(") > 0 Then
    fTestData.ComboBoxOperator.Value = "begins with"
   ElseIf InStr(PFString, "RIGHT(") > 0 Then
    fTestData.ComboBoxOperator.Value = "ends with"
   ElseIf InStr(PFString, "SEARCH(") > 0 Then
    fTestData.ComboBoxOperator.Value = "contains"
   ElseIf InStr(PFString, "ISERROR(") > 0 Then
    fTestData.ComboBoxOperator.Value = "does not contain"
   Else
    fTestData.ComboBoxOperator.Value = "equals"
   End If
  ElseIf InStr(PFString, "<>") > 0 Then
   OpType = "<>"
   If InStr(PFString, "LEFT(") > 0 Then
    fTestData.ComboBoxOperator.Value = "does not begin with"
   ElseIf InStr(PFString, "RIGHT(") > 0 Then
    fTestData.ComboBoxOperator.Value = "does not end with"
   Else
    fTestData.ComboBoxOperator.Value = "does not equal"
   End If
  ElseIf InStr(PFString, ">") > 0 Then
   OpType = ">"
   fTestData.ComboBoxOperator.Value = "is greater than"
  ElseIf InStr(PFString, "<") > 0 Then
   OpType = "<"
   fTestData.ComboBoxOperator.Value = "is less than"
  Else
   OpType = ""
   fTestData.ComboBoxOperator.Value = ""
  End If
 Else
  OpType = ""
  fTestData.ComboBoxOperator.Value = ""
 End If
 
 'Set TestData UserForm Current Left & Right Operands, If Any
 If OpType <> "" Then
  fTestData.TextBoxLeftOperand.Value = Mid(PFString, 1, InStr(PFString, OpType) - 1)
  fTestData.TextBoxRightOperand.Value = Mid(PFString, InStr(PFString, OpType) + Len(OpType), InStr(PFString, ",") - Len(OpType) - InStr(PFString, OpType))
  fTestData.TextBoxResultingFormula.Value = Cells(ActiveCell.Row, j + 1).Formula
  
  'Set TestData UserForm Current Pass/Fail Indicator, If Any
  If InStr(PFString, "Pass") < InStr(PFString, "Fail") Then
   fTestData.OptionPass.Value = True
  Else
   fTestData.OptionFail.Value = True
  End If
 
  'Clean Up Selected Data By Deleting Quotation Marks - Left Operand
  fTestData.TextBoxLeftOperand.Value = Replace(fTestData.TextBoxLeftOperand.Value, """", "")
  
  'Clean Up Selected Data By Deleting Quotation Marks - Right Operand
  fTestData.TextBoxRightOperand.Value = Replace(fTestData.TextBoxRightOperand.Value, """", "")
 Else
  'Clear UserForm Blanks If No Operator Selected
  fTestData.TextBoxLeftOperand.Value = ""
  fTestData.TextBoxRightOperand.Value = ""
  fTestData.TextBoxResultingFormula.Value = ""
  fTestData.OptionPass.Value = False
  fTestData.OptionFail.Value = False
 End If

 'Set TestData UserForm Current Left & Right Difference Values, If Any
 If Cells(ActiveCell.Row, j + 2).Formula <> "" Then
  PFString = Mid(Cells(ActiveCell.Row, j + 2).Formula, 2)
  fTestData.TextBoxLeftValue.Value = Mid(PFString, 1, InStr(PFString, "-") - 1)
  fTestData.TextBoxRightValue.Value = Mid(PFString, InStr(PFString, "-") + 1, Len(PFString) - InStr(PFString, "-"))
  fTestData.TextBoxResultingFormulaDiff.Value = Cells(ActiveCell.Row, j + 2).Formula
 Else
  'Clear UserForm Blanks If Difference Cell Is Blank
  fTestData.TextBoxLeftValue.Value = ""
  fTestData.TextBoxRightValue.Value = ""
  fTestData.TextBoxResultingFormulaDiff.Value = ""
 End If

 'Hide Completion Message Label
 fTestData.LabelComplete.Visible = False
 
 'Hide KDI & CalcInput UserForms If Currently Visible
 If fKDI.Visible = True Then
  fKDI.Hide
 End If
 If fCalcInput.Visible = True Then
  fCalcInput.Hide
 End If
 
 'Reset Toolbar Position
 Call ResetToolbarPosition
 
 'Set Top & Left Position & Show UserForm
 If fTestData.Visible = False Then
  fTestData.Top = Application.Height / 4
  fTestData.Left = (Application.Width - fTestData.Width) / 2
  fTestData.Show
 End If
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("TestDataFormSetup", Err.Number, Err.Description)
End Sub
Sub KDIFormSetup(KDICISheet As String)
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Sort KDI-CI Table
 Call SortKDICITable
 
 'Reinitialize Array Variable For Key Indicators Actual Size
 j = GetCIFirstRow - 2
 ReDim ListBoxArray(j, 3)
 
 If fKDI.Visible = False Then
  'Set Cell Values & References In KDI-CI Worksheet
  For i = 1 To j
   If Sheets(KDICISheet).Range("G" & i + 1).Value = "" Then
    Sheets(KDICISheet).Range("G" & i + 1).Value = "????"
   End If
   Sheets(KDICISheet).Range("H" & i + 1).Formula = _
   "=VLOOKUP(" & "A" & i + 1 & ",$A$2:$G$" & j + 1 & ",6,FALSE)"
   Sheets(KDICISheet).Range("I" & i + 1).Formula = _
   "=VLOOKUP(" & "A" & i + 1 & ",$A$2:$G$" & j + 1 & ",7,FALSE)"
  Next i
  
  'Set UserForm Current Test Cell Address
  fKDI.TextBoxKDICell.Value = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)

  'Populate Listbox List
  fKDI.ListBoxIndicators.Clear
  For i = 0 To j - 1
   ListBoxArray(i, 0) = i + 1 & ". " & Sheets(KDICISheet).Range("C" & i + 2)
   ListBoxArray(i, 1) = IIf(Sheets(KDICISheet).Range("H" & i + 2) = 0, "", Sheets(KDICISheet).Range("H" & i + 2))
   ListBoxArray(i, 2) = Sheets(KDICISheet).Range("I" & i + 2)
  Next i
  fKDI.ListBoxIndicators.List() = ListBoxArray
  fKDI.ListBoxIndicators.ListIndex = 0
  fKDI.ListBoxIndicators.ControlTipText = Sheets(KDICISheet).Range("E2")
  Application.StatusBar = Sheets(KDICISheet).Range("E2")
   
  'Hide TestData & CalcInput UserForms If Currently Visible
  If fTestData.Visible = True Then
   fTestData.Hide
  End If
  If fCalcInput.Visible = True Then
   fCalcInput.Hide
  End If
  
  'Reset Toolbar Position
  Call ResetToolbarPosition
 
  If fKDI.Visible = False Then
   'Set Top & Left Position & Show UserForm
   fKDI.Top = Application.Height / 4
   fKDI.Left = (Application.Width - fKDI.Width) / 2
   fKDI.Show
  End If
 Else
  'Set Focus Back To KDI UserForm Operand
  PFString = fKDI.FrameKeyDealIndicators.ActiveControl.Name
  fKDI.Controls(PFString).SetFocus
 End If

 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("KDIFormSetup", Err.Number, Err.Description)
End Sub
Sub CalcInputFormSetup(KDICISheet As String)
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Sort KDI-CI Table
 Call SortKDICITable
 
 'Check To See If UserForm Data Type ComboBox Needs To Be Initialized
 If fCalcInput.ComboBoxCIType.ListCount = 0 Then
  'Set List of Data Types
  fCalcInput.ComboBoxCIType.Clear
  fCalcInput.ComboBoxCIType.AddItem "Integer"
  fCalcInput.ComboBoxCIType.AddItem "Boolean"
  fCalcInput.ComboBoxCIType.AddItem "Decimal"
  fCalcInput.ComboBoxCIType.AddItem "Date"
  fCalcInput.ComboBoxCIType.AddItem "String"
  fCalcInput.ComboBoxCIType.AddItem "Money"
  fCalcInput.ComboBoxCIType.AddItem "Rate"
 End If
 
 'Find First & Last Calculated Input Rows In KDI-CI Table
 j = GetCIFirstRow
 k = GetLastRow(1, KDICISheet)
 
 If fCalcInput.Visible = False Then
  'Set Cell Values & References In KDI-CI Worksheet
  For i = j To k
   If Sheets(KDICISheet).Range("G" & i).Value = "" Then
    Sheets(KDICISheet).Range("G" & i).Value = "????"
   End If
   Sheets(KDICISheet).Range("H" & i).Formula = _
   "=VLOOKUP(" & "A" & i & ",$A$" & j & ":$G$" & k & ",6,FALSE)"
   Sheets(KDICISheet).Range("I" & i).Formula = _
   "=VLOOKUP(" & "A" & i & ",$A$" & j & ":$G$" & k & ",7,FALSE)"
  Next i
 
  'Find Selected Cell In Report Worksheet
  CIRange = "G" & j & ":G" & k
  CICell = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
  If WorksheetFunction.CountIf(Sheets(KDICISheet).Range(CIRange), CICell) > 0 Then
   i = Application.WorksheetFunction.Match(CICell, Sheets(KDICISheet).Range(CIRange), 0) + j - 1
   fCalcInput.TextBoxCIName.Value = Sheets(KDICISheet).Range("C" & i).Value
   fCalcInput.ComboBoxCIType.Value = Sheets(KDICISheet).Range("D" & i).Value
   fCalcInput.TextBoxCIDescr.Value = Sheets(KDICISheet).Range("E" & i).Value
   fCalcInput.TextBoxCIValue.Value = Sheets(KDICISheet).Range("F" & i).Value
   fCalcInput.TextBoxCICell.Value = Sheets(KDICISheet).Range("G" & i).Value
  Else
   'Clear UserForm Blanks If Cell Reference Not Found
   fCalcInput.TextBoxCIName.Value = ""
   If IsError(ActiveCell.Value) Then
    fCalcInput.ComboBoxCIType.Value = "String"
   ElseIf IsDate(ActiveCell.Value) Then
    fCalcInput.ComboBoxCIType.Value = "Date"
   ElseIf IsNumeric(Right(ActiveCell.Value, 1)) Then
    fCalcInput.ComboBoxCIType.Value = "Money"
   ElseIf InStr(ActiveCell.Value, "%") Then
    fCalcInput.ComboBoxCIType.Value = "Rate"
   Else
    fCalcInput.ComboBoxCIType.Value = "String"
   End If
   fCalcInput.TextBoxCIDescr.Value = ""
   fCalcInput.TextBoxCIValue.Value = IIf(IsError(ActiveCell.Value), "", ActiveCell.Value)
   fCalcInput.TextBoxCICell.Value = CICell
  End If
 Else
  'Set Focus Back To CI UserForm Operand
  PFString = fCalcInput.FrameCIData.ActiveControl.Name
  fCalcInput.Controls(PFString).SetFocus
 End If

 'Hide Completion Message Label
 fCalcInput.LabelComplete.Visible = False
 
 'Hide TestData & KDI UserForms If Currently Visible
 If fTestData.Visible = True Then
  fTestData.Hide
 End If
 If fKDI.Visible = True Then
  fKDI.Hide
 End If
 
 'Reset Toolbar Position
 Call ResetToolbarPosition
 
 'Set Top & Left Position & Show UserForm
 If fCalcInput.Visible = False Then
  fCalcInput.Top = Application.Height / 4
  fCalcInput.Left = (Application.Width - fCalcInput.Width) / 2
  fCalcInput.Show
 End If
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("CalcInputFormSetup", Err.Number, Err.Description)
End Sub
Sub SelectionInput(Target As Range)
 'Initialize Variable
 Dim MsgBoxQues As String
 Dim KDICISheet As String
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Insert Selected Value & Cell Reference In ListBox Popup
 If fTestData.Visible = True Then
  If Left(fTestData.FrameTestData.ActiveControl.Name, 11) <> "TextBoxLeft" _
  And Left(fTestData.FrameTestData.ActiveControl.Name, 12) <> "TextBoxRight" Then
   'Initialize Test Data UserForm, If Necessary
   Call TestDataFormSetup
  ElseIf Left(fTestData.FrameTestData.ActiveControl.Name, 11) = "TextBoxLeft" _
  Or Left(fTestData.FrameTestData.ActiveControl.Name, 12) = "TextBoxRight" Then
   'Insert Cell Address/Value Question
   Beep
   MsgBoxQues = MsgBox("Do you want to insert the selected cell's address?" & vbCr & _
   "(Answer 'No' to insert the selected cell's value.)", vbQuestion + vbYesNoCancel, gsAPP_NAME)
   
   'Insert Selected Cell In Active Operand
   If MsgBoxQues = vbYes Then
    fTestData.Controls(fTestData.FrameTestData.ActiveControl.Name).Value = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
   ElseIf MsgBoxQues = vbNo Then
    fTestData.Controls(fTestData.FrameTestData.ActiveControl.Name).Value = ActiveCell.Value
   End If
  
   'Set Focus Back To Test Data UserForm Operand
   PFString = fTestData.FrameTestData.ActiveControl.Name
   fTestData.ComboBoxOperator.SetFocus
   fTestData.Controls(PFString).SetFocus
  End If
 
  'Hide Completion Message Label
  fTestData.LabelComplete.Visible = False
 ElseIf fKDI.Visible = True Then
  i = fKDI.ListBoxIndicators.ListIndex
  j = fKDI.ListBoxIndicators.TopIndex
  ListBoxArray(i, 1) = Trim(Target.Text)
  If InStr(Target.Address(RowAbsolute:=False, ColumnAbsolute:=False), ":") > 0 Then
   ListBoxArray(i, 2) = _
   Mid(Target.Address(RowAbsolute:=False, ColumnAbsolute:=False), 1, _
   InStr(Target.Address(RowAbsolute:=False, ColumnAbsolute:=False), ":") - 1)
  Else
   ListBoxArray(i, 2) = _
   Target.Address(RowAbsolute:=False, ColumnAbsolute:=False)
  End If
  fKDI.ListBoxIndicators.List() = ListBoxArray
  fKDI.ListBoxIndicators.ListIndex = i
  fKDI.ListBoxIndicators.TopIndex = j
 Else
  'Find First & Last Calculated Input Rows In KDI-CI Table
  KDICISheet = GetKDICIWSName
  j = GetCIFirstRow
  k = GetLastRow(1, KDICISheet)
  
  'Find Selected Cell In Report Worksheet
  CIRange = "G" & j & ":G" & k
  CICell = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
  If WorksheetFunction.CountIf(Sheets(KDICISheet).Range(CIRange), CICell) > 0 Then
   i = Application.WorksheetFunction.Match(CICell, Sheets(KDICISheet).Range(CIRange), 0) + j - 1
   fCalcInput.TextBoxCIName.Value = Sheets(KDICISheet).Range("C" & i).Value
   fCalcInput.ComboBoxCIType.Value = Sheets(KDICISheet).Range("D" & i).Value
   fCalcInput.TextBoxCIDescr.Value = Sheets(KDICISheet).Range("E" & i).Value
   fCalcInput.TextBoxCIValue.Value = Sheets(KDICISheet).Range("F" & i).Value
   fCalcInput.TextBoxCICell.Value = Sheets(KDICISheet).Range("G" & i).Value
  Else
   fCalcInput.TextBoxCIName.Value = ""
   If IsError(ActiveCell.Value) Then
    fCalcInput.ComboBoxCIType.Value = "String"
   ElseIf IsDate(ActiveCell.Value) Then
    fCalcInput.ComboBoxCIType.Value = "Date"
   ElseIf IsNumeric(Right(ActiveCell.Value, 1)) Then
    fCalcInput.ComboBoxCIType.Value = "Money"
   ElseIf InStr(ActiveCell.Value, "%") Then
    fCalcInput.ComboBoxCIType.Value = "Rate"
   Else
    fCalcInput.ComboBoxCIType.Value = "String"
   End If
   fCalcInput.TextBoxCIDescr.Value = ""
   fCalcInput.TextBoxCIValue.Value = IIf(IsError(ActiveCell.Value), "", ActiveCell.Value)
   fCalcInput.TextBoxCICell.Value = CICell
  End If
 End If
 
 'Hide Completion Message Label
 fCalcInput.LabelComplete.Visible = False
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("SelectionInput", Err.Number, Err.Description)
End Sub
Sub HelpScreenContents()
 'Initialize Variable
 Dim HelpHeader As String
 Dim HelpPar As String
 Dim HelpParHeight As Integer
 Dim HelpHeaderTop As Integer
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Number of Help Sections
 j = 15
 
 'Set Cell Values & References In Key Deal Indicators Worksheet
 HelpHeaderTop = 96
 For i = 1 To j
  Select Case i
   Case 1: HelpHeader = "INSERT DATA INTO REPORT WORKSHEET"
   HelpPar = "The simplest way to insert your structured deal data into the Report worksheet " & _
   "is to cut and paste it from another spreadsheet file.  This process is automated for you " & _
   "when you click the 'Copy In Deal Data' button.  Doing this will open the file dialog box " & _
   "where you can select the file from which to copy the data worksheet." & vbCr & vbCr & _
   "You can also manually cut and paste deal data from a Word document or other form of text file.  " & _
   "If not already in any other digital file format, you'll need to manually type in your " & _
   "structured deal data."
   HelpParHeight = 96
   
   Case 2: HelpHeader = "DEAL TEST COLUMNS IN REPORT WORKSHEET"
   HelpPar = "After your structured deal data is inserted into the 'Report' worksheet, four deal test columns " & _
   "(Test Name, Test Result, Difference & Test Type) are automatically inserted to the right of your " & _
   "right-most data column in the 'Report' worksheet.  This occurs when you first click one of the " & _
   "program buttons.  These columns are used to store the deal test formulas and results."
   HelpParHeight = 60

   Case 3: HelpHeader = "FORMAT CUSTOM WORKSHEETS"
   HelpPar = "In certain cases, Data Reports created in DealManager will require custom formatting after " & _
   "the new Report opens up in Excel.  Click 'Format Custom Worksheets' to accomplish this process."
   HelpParHeight = 36

   Case 4: HelpHeader = "REPORT SETTINGS"
   HelpPar = "The Report Settings worksheet stores various values that control how the deal data is organized."
   HelpParHeight = 24
   
   Case 5: HelpHeader = "CONFIGURE DEAL TESTS"
   HelpPar = "Select any cell on the 'Report' worksheet, and then click 'Define/Modify Deal Tests' " & _
   "to add or modify the deal test in that worksheet row." & vbCr & vbCr & _
   "'Pass/Fail' Deal Tests are created/modified by first entering/editing a Test Name and then selecting " & _
   "a Test Type.  Next, the Left & Right Operands are entered/edited by selecting the blank beneath each " & _
   "and then selecting the appropriate worksheet cell.  After this, select an Operator from the drop-down list " & _
   "and then indicate Pass/Fail on the right-hand side of the form." & vbCr & vbCr & _
   "'Difference' Deal Tests are similarly created/modified by selecting Left & Right Values in the same manner " & _
   "as the Left & Right Operands are selected as described above."
   HelpParHeight = 132
   
   Case 6: HelpHeader = "CONFIGURE KEY DEAL INDICATORS"
   HelpPar = "Select any cell on the 'Report' worksheet, and then click 'Define Key Deal Indicators' " & _
   "to map deal data to the corresponding key deal indicators.  KDI values are selected by selecting " & _
   "a KDI from the form list and then selecting the appropriate worksheet cell."
   HelpParHeight = 36
   
   Case 7: HelpHeader = "CONFIGURE CALCULATED INPUTS"
   HelpPar = "Select any cell on the 'Report' worksheet, and then click 'Define Calculated Inputs' " & _
   "to map deal data to be uploaded to the DealManager database.  CI values are selected by selecting " & _
   "a cell from the 'Report' worksheet."
   HelpParHeight = 36
   
   Case 8: HelpHeader = "CHECK FOR DEAL TEST FORMULA ERRORS"
   HelpPar = "Click the 'Check For Formula Errors' button to review the 'Error Log' worksheet for " & _
   "deal test formula errors in the 'Report' worksheet.  These will need to be corrected in the " & _
   "'Report' worksheet before the deal test data can be uploaded into the DealManager database."
   HelpParHeight = 36
   
   Case 9: HelpHeader = "CHECK FOR DEAL TEST VIOLATIONS"
   HelpPar = "Click the 'Check For Deal Violations' button to review the 'Error Log' worksheet for " & _
   "deal test result violations in the 'Report' worksheet.  These will need to be corrected in the " & _
   "'Report' worksheet before the deal test data can be uploaded into the DealManager database."
   HelpParHeight = 48
   
   Case 10: HelpHeader = "UPDATE 'TESTS' WORKSHEET"
   HelpPar = "Click the 'Update Tests Worksheet' button to copy all deal tests data to the 'Tests' " & _
   "worksheet in preparation for uploading to the DealManager database."
   HelpParHeight = 24
   
   Case 11: HelpHeader = "UPLOAD DEAL TEST DATA TO DEALMANAGER DATABASE"
   HelpPar = "Click the 'Upload Test Data to DM' button to upload all deal tests data to the DealManager database."
   HelpParHeight = 24
   
   Case 12: HelpHeader = "UPLOAD KDI & CI DATA TO DEALMANAGER DATABASE"
   HelpPar = "Click the 'Upload KDI & CI Data to DM' button to upload all Key Deal Indicator and Calculated " & _
   "Input data to the DealManager database."
   HelpParHeight = 24
   
   Case 13: HelpHeader = "CREATE TRIAL REPORT"
   HelpPar = "This button creates the Trial Report by copying the deal information on the 'Report' worksheet " & _
   "into a new workbook file, retaining all formulas and links."
   HelpParHeight = 24
   
   Case 14: HelpHeader = "CREATE FINAL REPORT"
   HelpPar = "This button creates the Final Report by copying the deal information on the 'Report' worksheet " & _
   "into a new workbook file, but retaining just the values and not any formulas or links."
   HelpParHeight = 36
   
   Case 15: HelpHeader = "WORKSHEET NAVIGATION"
   HelpPar = "To move among the worksheets in this workbook file, select the worksheet name from the drop-down " & _
   "list at the bottom of the floating toolbar.  You may also click the worksheet tab as you would with " & _
   "any other Excel file."
   HelpParHeight = 36
  End Select
  
  'Insert & Format Help Screen Header & Paragraph
  fHelp.Controls("LabelParHeader" & i).Caption = i & ". " & HelpHeader
  fHelp.Controls("LabelParHeader" & i).Top = HelpHeaderTop
  fHelp.Controls("LabelPar" & i).Caption = HelpPar
  fHelp.Controls("LabelPar" & i).Height = HelpParHeight
  fHelp.Controls("LabelPar" & i).Top = HelpHeaderTop + 14
  HelpHeaderTop = HelpHeaderTop + 14 + 22 + HelpParHeight - 12
 Next i

 'Reset Help Screen Size Based On Size of All Content
 fHelp.ScrollHeight = HelpHeaderTop + 12
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("HelpScreenContents", Err.Number, Err.Description)
End Sub
