Attribute VB_Name = "mWorksheetSetup"
Option Explicit
'Initialize Module Level Variables
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Long
Dim o As Long
Sub WorksheetsSetup()
 'Initialize Variables
 Dim WSName() As String
 Dim wsCount As Integer
 
 'Set wsName Array Size
 wsCount = 6
 ReDim WSName(wsCount)
 
 'Set Worksheet Names
 k = 0
 WSName(k) = "Settings": k = k + 1
 WSName(k) = "Error Log": k = k + 1
 WSName(k) = "Tests": k = k + 1
 WSName(k) = "Inputs": k = k + 1
 WSName(k) = "KDI-CI": k = k + 1
 WSName(k) = "Data": k = k + 1
 
 'Check For & Create Worksheets, As Necessary
 For k = 0 To wsCount - 1
  'Set Up Worksheet, If Necessary
  Call AddSheet(WSName(k))
 Next k
End Sub
Sub AddSheet(AddSheetName As String)
 'Initialize Variable
 Dim sh As Worksheet

 'Set Up New Worksheet, If Necessary
 For Each sh In Sheets
  If sh.Name = AddSheetName Then
   'Set Template Version Number
   If sh.Name = "Settings" Then
    If NRCheck("TVersion") = True Then
     Range("TVersion").Value = gsVERSION
     Range("A1").Select
    End If
   End If
   
   'Exit Loop
   Exit For
  ElseIf sh.Name = Sheets(Sheets.Count).Name Then
   'Add New Worksheet After Last Existing Worksheet
   Worksheets.Add(After:=Sheets(Sheets.Count)).Name = AddSheetName
   
   'Set Worksheet Color To Light Gray
   Cells.Interior.ColorIndex = 15
   
   'Select First Cell In Worksheet
   Range("A1").Select
  
   'Format Certain Worksheets
   Select Case AddSheetName
    Case "Settings": Call SettingsWSSetup(AddSheetName)
    Case "Error Log": Call ErrorLogWSSetup(AddSheetName)
    Case "Tests": Call TestsWSSetup(AddSheetName)
    Case "KDI-CI": Call KDICIWSSetup(AddSheetName)
   End Select
  End If
 Next sh
End Sub
Sub SettingsWSSetup(AddSheetName As String)
 'Initialize Variables
 Dim SettingName() As String
 Dim SettingCount As Integer
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set SettingName Array Size
 SettingCount = 10
 ReDim SettingName(SettingCount)
 
 'Set Setting Names
 i = 0
 SettingName(i) = "Deal ID": i = i + 1
 SettingName(i) = "Deal Name": i = i + 1
 SettingName(i) = "Worksheet Lock": i = i + 1
 SettingName(i) = "Load Test Indicator (1-=Final)": i = i + 1
 SettingName(i) = "Test Ind ( 1 = Fail for any DealTest)": i = i + 1
 SettingName(i) = "Toolbar Top Position": i = i + 1
 SettingName(i) = "Toolbar Left Position": i = i + 1
 SettingName(i) = "Template Version": i = i + 1
 SettingName(i) = "Temporary Value #1": i = i + 1
 SettingName(i) = "Temporary Value #2": i = i + 1

 'Set Up Header Row
 With Sheets(AddSheetName).Range("A1")
  .Value = AddSheetName
  .Font.Bold = True
  .HorizontalAlignment = xlCenter
 End With
 With Sheets(AddSheetName).Range("B1")
  .Value = "Value"
  .Font.Bold = True
 End With
 
 'Check For & Create Settings List
 For i = 0 To SettingCount - 1
  'Set Up Setting Row
  Sheets(AddSheetName).Range("A" & i + 2).Value = SettingName(i)
  
  'Set Up Some Value Rows
  Select Case SettingName(i)
   Case "Deal ID"
   ThisWorkbook.Names.Add Name:="DealID", RefersTo:="=Settings!$B$" & i + 2
   Range("DealID").Value = 1
   Case "Deal Name"
   ThisWorkbook.Names.Add Name:="DealName", RefersTo:="=Settings!$B$" & i + 2
   Range("DealName").Value = "New Deal"
   Case "Worksheet Lock"
   ThisWorkbook.Names.Add Name:="WSLock", RefersTo:="=Settings!$B$" & i + 2
   Range("WSLock").Value = 0
   Case "Load Test Indicator (1-=Final)"
   ThisWorkbook.Names.Add Name:="FinalInd", RefersTo:="=Settings!$B$" & i + 2
   Range("FinalInd").Value = 0
   Case "Test Ind ( 1 = Fail for any DealTest)"
   ThisWorkbook.Names.Add Name:="TestInd", RefersTo:="=Settings!$B$" & i + 2
   Range("TestInd").Value = 0
   Case "Toolbar Top Position"
   ThisWorkbook.Names.Add Name:="TBTop", RefersTo:="=Settings!$B$" & i + 2
   Range("TBTop").Value = 0
   Case "Toolbar Left Position"
   ThisWorkbook.Names.Add Name:="TBLeft", RefersTo:="=Settings!$B$" & i + 2
   Range("TBLeft").Value = 0
   Case "Template Version"
   ThisWorkbook.Names.Add Name:="TVersion", RefersTo:="=Settings!$B$" & i + 2
   Range("TVersion").Value = gsVERSION
   Case "Temporary Value #1"
   ThisWorkbook.Names.Add Name:="TValue1", RefersTo:="=Settings!$B$" & i + 2
   Range("TValue1").Value = ""
   Case "Temporary Value #2"
   ThisWorkbook.Names.Add Name:="TValue2", RefersTo:="=Settings!$B$" & i + 2
   Range("TValue2").Value = ""
  End Select
 Next i

 'Format Setting Names Column - Bold, Column Width & Horizontal Alignment
 With Sheets(AddSheetName).Range("A2:A" & SettingCount + 1)
  .Font.Bold = True
  .Columns.AutoFit
  .HorizontalAlignment = xlLeft
 End With

 'Format Setting Value Column - Column Width & Horizontal Alignment
 With Sheets(AddSheetName).Range("B1:B" & SettingCount + 1)
  .Font.Size = 8
  .ColumnWidth = 35
  .HorizontalAlignment = xlCenter
 End With

 'Format Settings Group Borders & Shading
 With Sheets(AddSheetName).Range("A1:B" & SettingCount + 1)
  .Borders(xlEdgeLeft).Weight = xlThick
  .Borders(xlEdgeTop).Weight = xlThick
  .Borders(xlEdgeBottom).Weight = xlThick
  .Borders(xlEdgeRight).Weight = xlThick
  .Borders(xlInsideVertical).Weight = xlThin
  .Borders(xlInsideHorizontal).Weight = xlThin
  .Interior.ColorIndex = xlNone
 End With
 Sheets(AddSheetName).Range("A1:B1").Borders(xlEdgeBottom).Weight = xlThick

 'Select Cell D1 In Worksheet
 Range("D1").Select
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("SettingsWSSetup", Err.Number, Err.Description)
End Sub
Sub ErrorLogWSSetup(AddSheetName As String)
 'Initialize Variables
 Dim ErrorColName() As String
 Dim ErrorColCount As Integer
 Dim ErrorRows As Integer
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set ErrorColName Array Size
 ErrorColCount = 5
 ErrorRows = 2
 ReDim ErrorColName(ErrorColCount)
 
 'Set Error Log Column Names
 i = 0
 ErrorColName(i) = "Error" & vbLf & "Number": i = i + 1
 ErrorColName(i) = "Error Description": i = i + 1
 ErrorColName(i) = "VBA Procedure" & vbLf & "Error Occurred In": i = i + 1
 ErrorColName(i) = "Error Time": i = i + 1
 ErrorColName(i) = "Filename": i = i + 1

 'Insert Error Log Table Column Names
 For i = 0 To ErrorColCount - 1
  'Set Up Error Log Column Name
  Sheets(AddSheetName).Cells(1, i + 1).Value = ErrorColName(i)
 Next i
 
 'Format Table
 Call WSFormat(AddSheetName, ErrorColCount, ErrorRows)
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ErrorLogWSSetup", Err.Number, Err.Description)
End Sub
Sub TestsWSSetup(AddSheetName As String)
 'Initialize Variables
 Dim TestColName() As String
 Dim TestColCount As Integer
 Dim TestRows As Integer
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set TestColName Array Size
 TestColCount = 5
 TestRows = 2
 ReDim TestColName(TestColCount)
 
 'Set Test Log Column Names
 i = 0
 TestColName(i) = "Test Name": i = i + 1
 TestColName(i) = "Test" & vbLf & "Result": i = i + 1
 TestColName(i) = "Difference": i = i + 1
 TestColName(i) = "Test Type": i = i + 1
 TestColName(i) = "Cell" & vbLf & "Reference": i = i + 1

 'Insert Test Log Table Column Names
 For i = 0 To TestColCount - 1
  'Set Up Test Log Column Name
  Sheets(AddSheetName).Cells(1, i + 1).Value = TestColName(i)
 Next i
 
 'Format Table
 Call WSFormat(AddSheetName, TestColCount, TestRows)
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("TestsWSSetup", Err.Number, Err.Description)
End Sub
Sub KDICIWSSetup(AddSheetName As String)
 'Initialize Variables
 Dim InputColName() As String
 Dim InputColCount As Integer
 Dim InputRows As Integer
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set InputColName Array Size
 InputColCount = 9
 InputRows = 2
 ReDim InputColName(InputColCount)
 
 'Set Input Log Column Names
 i = 0
 InputColName(i) = "ID": i = i + 1 'Automatically Generated
 InputColName(i) = "Source": i = i + 1
 InputColName(i) = "Name": i = i + 1
 InputColName(i) = "Type": i = i + 1
 InputColName(i) = "Description": i = i + 1
 InputColName(i) = "Cell Value": i = i + 1
 InputColName(i) = "Cell" & vbLf & "Address": i = i + 1
 InputColName(i) = "Vlookup Value": i = i + 1
 InputColName(i) = "Vlookup" & vbLf & "Address": i = i + 1

 'Insert Input Log Table Column Names
 For i = 0 To InputColCount - 1
  'Set Up Input Log Column Name
  Sheets(AddSheetName).Cells(1, i + 1).Value = InputColName(i)
 Next i
 
 'Format Table
 Call WSFormat(AddSheetName, InputColCount, InputRows)
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("KDICIWSSetup", Err.Number, Err.Description)
End Sub
Sub WSFormat(AddSheetName As String, ColCount As Integer, RowCount As Integer)
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Format Header Row
 With Sheets(AddSheetName).Range("A1:" & Cells(1, ColCount).Address)
  .Font.Bold = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlBottom
  .Borders(xlEdgeLeft).Weight = xlThick
  .Borders(xlEdgeTop).Weight = xlThick
  .Borders(xlEdgeBottom).Weight = xlThick
  .Borders(xlEdgeRight).Weight = xlThick
  .Borders(xlInsideVertical).Weight = xlThin
  .Interior.ColorIndex = xlNone
  .RowHeight = 28
  .WrapText = True
 End With
 
 'Format Data Rows
 With Sheets(AddSheetName).Range("A2:" & _
 Cells(IIf(GetLastRow(1, AddSheetName) < 3, 3, GetLastRow(1, AddSheetName)), ColCount).Address)
  .Font.Size = 8
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlTop
  .WrapText = True
  .Borders(xlEdgeLeft).Weight = xlThick
  .Borders(xlEdgeBottom).Weight = xlThick
  .Borders(xlEdgeRight).Weight = xlThick
  .Borders(xlInsideVertical).Weight = xlThin
  .Borders(xlInsideHorizontal).Weight = xlThin
  .Interior.ColorIndex = xlNone
 End With
 
 'Format Column Widths
 Select Case AddSheetName
  Case "Settings"
  Case "Error Log"
  Sheets(AddSheetName).Columns("A").ColumnWidth = 10
  Sheets(AddSheetName).Columns("B").ColumnWidth = 25
  Sheets(AddSheetName).Columns("C").ColumnWidth = 20
  Sheets(AddSheetName).Columns("D").ColumnWidth = 15
  Sheets(AddSheetName).Columns("E").ColumnWidth = 40
  'Set Time Format Error Time Column
  Sheets(AddSheetName).Range("D2:D" & RowCount + 1).NumberFormat = "m/d/yyyy hh:mm:ss"
  Case "Tests"
  Sheets(AddSheetName).Columns("A").ColumnWidth = 25
  Sheets(AddSheetName).Columns("B").ColumnWidth = 8
  Sheets(AddSheetName).Columns("C").ColumnWidth = 14
  Sheets(AddSheetName).Columns("D").ColumnWidth = 10
  Sheets(AddSheetName).Columns("E").ColumnWidth = 10
  'Format Difference Column As Currency
  Sheets(AddSheetName).Range("C2:C" & RowCount + 1).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
  Case "KDI-CI"
  Sheets(AddSheetName).Columns("A").ColumnWidth = 3
  Sheets(AddSheetName).Columns("B").ColumnWidth = 15
  Sheets(AddSheetName).Columns("C").ColumnWidth = 40
  Sheets(AddSheetName).Columns("D").ColumnWidth = 7
  Sheets(AddSheetName).Columns("E").ColumnWidth = 25
  Sheets(AddSheetName).Columns("F").ColumnWidth = 20
  Sheets(AddSheetName).Columns("G").ColumnWidth = 10
  Sheets(AddSheetName).Columns("H").ColumnWidth = 20
  Sheets(AddSheetName).Columns("I").ColumnWidth = 10
  Sheets(AddSheetName).Range("F1").Value = "Cell Value"
  Sheets(AddSheetName).Range("G1").Value = "Cell Address"
  Sheets(AddSheetName).Range("H1").Value = "Vlookup Value"
  Sheets(AddSheetName).Range("I1").Value = "Vlookup Address"
  Case Else
  Sheets(AddSheetName).Range("A1:" & Sheets(AddSheetName).Range("A1").End(xlToRight).Address).Columns.AutoFit
  n = Sheets(AddSheetName).Range("A1").End(xlDown).Row
  For m = 1 To Sheets(AddSheetName).Range("A1").End(xlToRight).Column
   If IsNumeric(Sheets(AddSheetName).Cells(2, m)) And _
   Len(Sheets(AddSheetName).Cells(2, m)) > 4 And _
   Right(Sheets(AddSheetName).Cells(1, m), 2) <> "Id" Then
    Sheets(AddSheetName).Range(Sheets(AddSheetName).Cells(1, m).Address & ":" & _
    Sheets(AddSheetName).Cells(IIf(n > 65000, 5, n), m).Address).NumberFormat = _
    "$#,##0.00_);[Red]($#,##0.00)"
   ElseIf IsDate(Sheets(AddSheetName).Cells(2, m)) Then
    Sheets(AddSheetName).Columns(m).AutoFit
   ElseIf Right(Sheets(AddSheetName).Cells(1, m), 2) = "Dt" Then
    Sheets(AddSheetName).Columns(m).ColumnWidth = 8
   End If
  Next m
  Select Case AddSheetName
   Case "Obligors", "Rated", "NonRated"
   'Format Percent Column As Percentage
   n = Sheets(AddSheetName).Range("E1").End(xlDown).Row
   Sheets(AddSheetName).Range("E2:E" & IIf(n > 65000, 5, n)).NumberFormat = "0.00%;[Red](0.00%)"
  End Select
 End Select
 
 'Select Cell Just Right of Table In Worksheet
 Sheets(AddSheetName).Select
 Sheets(AddSheetName).Cells(1, ColCount + 2).Select
 Application.ScreenUpdating = True
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("WSFormat", Err.Number, Err.Description)
End Sub
Sub WSLinkInserts()
 'Initialize Variables
 Dim WSName As String
 Dim NRName As String
 Dim Lastcol As String
 Dim LookUpCol As String
 Dim ValueCol As String
    
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Select First Cell in Report Worksheet
 'Call ReportFirstCell
  
 'Disable Events
 Application.EnableEvents = False
 
 'Determine Preceding Month
 WSName = "Inputs"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 j = Application.WorksheetFunction.Match("Effective Date", Sheets(WSName).Range("A1:" & Sheets(WSName).Cells(1, i).Address), 0)
 k = Month(Sheets(WSName).Cells(2, j)) - 1
 If k = 0 Then
  k = 12
 End If
 For m = 2 To Sheets(WSName).Cells(2, j).End(xlDown).Row + 1
  If Month(Sheets(WSName).Cells(m, j)) = k Then
   'Exit Loop
   Exit For
  End If
 Next m
 
 'Determine Next Preceding Month
 k = Month(Sheets(WSName).Cells(2, j)) - 2
 If k = 0 Then
  k = 12
 ElseIf k = -1 Then
  k = 11
 End If
 For n = 2 To Sheets(WSName).Cells(2, j).End(xlDown).Row + 1
  If Month(Sheets(WSName).Cells(n, j)) = k Then
   'Exit Loop
   Exit For
  End If
 Next n
 
 'Determine Last Column - Government Worksheet
 WSName = "Government"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(5, i).Address
 
 'Insert Link to Government Worksheet - Deal Name (NEED TO GET SOMEWHERE ELSE!)
 NRName = "DealName"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=UPPER(HLOOKUP(""Deal Name""," & WSName & _
  "!$A$1:" & Lastcol & ",2,FALSE))"
 End If
 
 'Determine Last Column - Inputs Worksheet
 WSName = "Inputs"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(3, i).Address
 
 'Insert Link to Inputs Worksheet - Deal Date (this is most recent Effective Date of data!)
 ' removed DATEVALUE b/c error on bill malloy's machine
 
 NRName = "DealDate"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=(HLOOKUP(""Effective Date""," & WSName & _
  "!$A$1:" & Lastcol & ",2,FALSE))"
 End If
 
 

 


 'Determine Last Column - Inputs Worksheet
 WSName = "Inputs"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(2, i).Address
 
 'Insert Link to Inputs Worksheet - Portfolio Receivables Beginning Balance (4a)
 NRName = "AgingBegBal"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=IF(HLOOKUP(""Included Balance""," & WSName & _
  "!$A$1:" & Sheets(WSName).Cells(m, i).Address & "," & m & ",FALSE)=""""," & IIf(Range("TValue3") = "", 0, "TValue3") & _
  ",VALUE(HLOOKUP(""Included Balance""," & WSName & "!$A$1:" & _
  Sheets(WSName).Cells(m, i).Address & "," & m & ",FALSE)))"
'  Range(NRName).Formula = "=IF(HLOOKUP(""Included Balance""," & WSName & _
  "!$A$1:" & Sheets(WSName).Cells(m, i).Address & "," & m & ",FALSE)="""",0," & _
  "VALUE(HLOOKUP(""Included Balance""," & WSName & "!$A$1:" & _
  Sheets(WSName).Cells(m, i).Address & "," & m & ",FALSE)))"
 End If
 
 
 

 'Determine Last Column - Fixed Worksheet
 WSName = "Fixed"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(1, i).Address
 
 'Insert Link to Fixed Worksheet - Fixed Rate Leases (9a)
NRName = "PoolFRNV"
 If NRCheck(NRName) = True Then
  'Determine FixedFloatInd Column & Lookup Range
  i = Application.WorksheetFunction.Match("FixedFloatInd", Sheets(WSName).Range("A1:" & Lastcol), 0)
  LookUpCol = Sheets(WSName).Cells(2, i).Address & ":" & _
  Sheets(WSName).Cells(GetLastRow(1, WSName), i).Address
  
  'Determine SecuritizedValue Column & Lookup Range
  i = Application.WorksheetFunction.Match("SecuritizedValue", Sheets(WSName).Range("A1:" & Lastcol), 0)
  ValueCol = Sheets(WSName).Cells(2, i).Address & ":" & _
  Sheets(WSName).Cells(GetLastRow(1, WSName), i).Address
  
  'Set Formula - SumIf For FixedFloatInd = Fixed
  Range(NRName).Formula = "=SUMIF(" & WSName & "!" & LookUpCol & ",""Fixed""," & WSName & "!" & ValueCol & ")"
 End If

 
 
 'Determine Last Column - Government Worksheet
 WSName = "Government"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(1, i).Address
 
 'Insert Link to Government Worksheet - Governmental Authority Leases (10a)
 NRName = "PoolGANV"
 If NRCheck(NRName) = True Then
  'Determine ParentId Column & Lookup Range NOTE Change to ABSGovt ID 4/11
  i = Application.WorksheetFunction.Match("ABSGovernment_Ind", Sheets(WSName).Range("A1:" & Lastcol), 0)
  LookUpCol = Sheets(WSName).Cells(2, i).Address & ":" & _
  Sheets(WSName).Cells(GetLastRow(1, WSName), i).Address
  
 
  
  'Determine SecuritizedValue Column & Lookup Range
  i = Application.WorksheetFunction.Match("SecuritizedValue", Sheets(WSName).Range("A1:" & Lastcol), 0)
  ValueCol = Sheets(WSName).Cells(2, i).Address & ":" & _
  Sheets(WSName).Cells(GetLastRow(1, WSName), i).Address
  
  'Set Formula - SumIf For ParentId = 81020  NOTE CHANGE TO 1 4/11
  Range(NRName).Formula = "=SUMIF(" & WSName & "!" & LookUpCol & _
  ",""1""," & WSName & "!" & ValueCol & ")"
  
   Range(NRName).Formula = "=SUMIF(" & WSName & "!" & LookUpCol & _
  ",""81020""," & WSName & "!" & ValueCol & ")"
 End If

 
 
 'Determine Last Column - Closed End Worksheet
 WSName = "Closed End"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(1, i).Address
 
 'Insert Link to Closed End Worksheet - Closed End Leases (11a)
 NRName = "PoolCENV"
 If NRCheck(NRName) = True Then
  'Determine ClosedOpenInd Column & Lookup Range
  i = Application.WorksheetFunction.Match("ClosedOpenInd", Sheets(WSName).Range("A1:" & Lastcol), 0)
  LookUpCol = Sheets(WSName).Cells(2, i).Address & ":" & _
  Sheets(WSName).Cells(GetLastRow(1, WSName), i).Address
  
  'Determine SecuritizedValue Column & Lookup Range
  i = Application.WorksheetFunction.Match("SecuritizedValue", Sheets(WSName).Range("A1:" & Lastcol), 0)
  ValueCol = Sheets(WSName).Cells(2, i).Address & ":" & _
  Sheets(WSName).Cells(GetLastRow(1, WSName), i).Address
  
  'Set Formula - SumIf For ClosedOpenInd = Closed
  Range(NRName).Formula = "=SUMIF('" & WSName & "'!" & LookUpCol & _
  ",""Closed"",'" & WSName & "'!" & ValueCol & ")"
 End If

 
 
 'Determine Last Column - Inputs Worksheet
 WSName = "Inputs"
 i = Sheets(WSName).Range("A1").End(xlToRight).Column
 Lastcol = Sheets(WSName).Cells(4, i).Address
 
 
 
 'Insert Link to Inputs Worksheet - RentalProceeds (18a)
 NRName = "iRentalProceeds"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=VALUE(HLOOKUP(""RentalProceeds""," & WSName & _
  "!$A$1:" & Lastcol & ",2,FALSE))"
 End If
 
 'Insert Link to Inputs Worksheet - SalesProceeds (18b)
 NRName = "iSalesProceeds"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=VALUE(HLOOKUP(""SalesProceeds""," & WSName & _
  "!$A$1:" & Lastcol & ",2,FALSE))"
 End If
 
 'Insert Link to Inputs Worksheet - HedgeReceipts (18g)
 'NRName = "iHedgeRec"
 'If NRCheck(NRName) = True Then
  'Range(NRName).Formula = "=VALUE(HLOOKUP(""HedgeReceipts""," & WSName & _
 ' "!$A$1:" & LastCol & ",2,FALSE))"
 'End If
 
 'Insert Link to Inputs Worksheet - TrusteeInterest (18i)
 'NRName = "iTrusteeInt"
 'If NRCheck(NRName) = True Then
 ' Range(NRName).Formula = "=VALUE(HLOOKUP(""TrusteeInterest""," & WSName & _
 ' "!$A$1:" & LastCol & ",2,FALSE))"
' End If
 
 'Insert Link to Inputs Worksheet - ProgFinCharges (19a)
 'NRName = "iProgFinChg"
 'I'f NRCheck(NRName) = True Then
 ' Range(NRName).Formula = "=VALUE(HLOOKUP(""ProgFinCharges""," & WSName & _
  '"!$A$1:" & LastCol & ",2,FALSE))"
 'End If

 'Insert Link to Inputs Worksheet - LiqCommitFee (19c)
 'NRName = "iLiqComFee"
 'If NRCheck(NRName) = True Then
 ' Range(NRName).Formula = "=VALUE(HLOOKUP(""LiqCommitFee""," & WSName & _
 ' "!$A$1:" & LastCol & ",2,FALSE))"
' End If
 
 'Insert Link to Inputs Worksheet - ProgramFee (19e)
  ' ****** Never been used - removed from SR 201403
' NRName = "iProgFee"
' If NRCheck(NRName) = True Then
'  Range(NRName).Formula = "=VALUE(HLOOKUP(""ProgramFee""," & WSName & _
'  "!$A$1:" & LastCol & ",2,FALSE))"
'End If
 
 'Insert Link to Inputs Worksheet - HedgePayments (19g)
' NRName = "iHedgePmts"
' If NRCheck(NRName) = True Then
'  Range(NRName).Formula = "=VALUE(HLOOKUP(""HedgePayments""," & WSName & _
'  "!$A$1:" & LastCol & ",2,FALSE))"
' End If
 
 'Insert Link to Inputs Worksheet - EarlyPaymentCosts (19n)
 ' ****** Never been used - removed from SR 201403
' NRName = "iEPC"
' If NRCheck(NRName) = True Then
'  Range(NRName).Formula = "=VALUE(HLOOKUP(""EarlyPaymentCosts""," & WSName & _
'  "!$A$1:" & LastCol & ",2,FALSE))"
' End If
 
 'Insert Link to Inputs Worksheet - OtherAmounts (19o)
' NRName = "iOther"
' If NRCheck(NRName) = True Then
'  Range(NRName).Formula = "=VALUE(HLOOKUP(""OtherAmounts""," & WSName & _
'  "!$A$1:" & LastCol & ",2,FALSE))"
' End If
 
 
 
 'Insert Link to Inputs Worksheet - Delinquent Units Beginning Balance - Prior Month (5d)
 ' 20090501 - need to change this from Del bukcets to the Delinquent Balance field.
 
 
 NRName = "DelBegBalPM"
 'If NRCheck(NRName) = True Then
 ' Range(NRName).Formula = "=VALUE(HLOOKUP(""Delinquent Units Beginning Balance""," & WSName & _
 ' "!$A$1:" & Lastcol & ",3,FALSE))"
 
 'End If
 
 'Insert Link to Inputs Worksheet - Delinquent Units Beginning Balance - 2nd Prior Month (5g)
 NRName = "DelBegBalPM2"
 'If NRCheck(NRName) = True Then
 ' Range(NRName).Formula = "=VALUE(HLOOKUP(""Delinquent Units Beginning Balance""," & WSName & _
 ' "!$A$1:" & Lastcol & ",4,FALSE))"

 'End If
 

 
 'Insert Link to Inputs Worksheet - Loan Beginning Balance (13a)
 NRName = "LoanBegBal"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=VALUE(HLOOKUP(""Ending Loan Balance""," & WSName & _
  "!$A$1:" & Lastcol & ",3,FALSE))"
' End If
 
 'Insert Link to Inputs Worksheet - Loan Payments (13b)
 NRName = "LoanPmts"
 If NRCheck(NRName) = True Then
  Range(NRName).Formula = "=-VALUE(HLOOKUP(""Loan Payments""," & WSName & _
  "!$A$1:" & Lastcol & ",3,FALSE))"
 End If
 
 
 
 'Insert Rated Clients
 WSName = "Obligors" 'changed on 20120201 for pnc"Rated"
 NRName = "RatedHeader"
 If NRCheck(NRName) = True Then
  'Set First Row of Rated Clients Section
  j = Range(NRName).Row
  k = j
  m = Application.WorksheetFunction.Match("ParentName", Sheets(WSName).Range("A1:D1"), 0)
  n = Application.WorksheetFunction.Match("SP_Rating", Sheets(WSName).Range("A1:D1"), 0)
  o = Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Sheets(WSName).Range("A1:D1"), 0)
 
  'Loop Through Rated Worksheet To Select All Clients
  For i = 2 To 11 ' only need top 10 GetLastRow(1, WSName)
   If Sheets(WSName).Cells(i, o).Value <> "" Then
    k = k + 1
   ' Sheets("Report").Rows(k).EntireRow.Insert no longer need to insert
    Sheets("Report").Range("c" & k).Value = Sheets(WSName).Cells(i, n).Value ' s&p rating
    Sheets("Report").Range("g" & k).Value = Sheets(WSName).Cells(i, m).Value ' name
   ' Sheets("Report").Range("G" & k).Value = 0.04 now a formula =INDEX($AC$82:$AD$85,MATCH(T81,$AC$82:$AC$85,0),2)
    Sheets("Report").Range("I" & k).Value = Sheets(WSName).Range("E" & i).Value ' acctual pct
    Sheets("Report").Range("j" & k).Value = Sheets(WSName).Cells(i, o).Value ' actual value
    Sheets("Report").Range("K" & k).Formula = "=IF(i" & k & "<h" & k & ",0,(i" & k & "-h" & k & ")*$j$12)"
   ' Sheets("Report").Range("J" & k).Value = "=IF(((H" & k & "-G" & k & ")*I" & k & ")>0," & _
   ' "(H" & k & "-G" & k & ")*I" & k & ",0)" now a formula =IF(H83<G83,0,(H83-G83)*$I$12)
   End If
  Next i
 
  'Format Percentage & Amount Cells
  If k > j Then
   Sheets("Report").Range("H" & j + 1 & ":I" & k).NumberFormat = "0.000%;[Red](0.000%)"
   Sheets("Report").Range("J" & j + 1 & ":K" & k).NumberFormat = "_($* #,##0_);[Red]_($* (#,##0);_($* "" - ""??_);_(@_)"
   Sheets("Report").Rows(j + 1 & ":" & k).RowHeight = 12.75
   Sheets("Report").Rows(j + 1 & ":" & k).VerticalAlignment = xlBottom
  End If
 
  'Set Rated Total Excess Amount
   ' no longer need to set total, b/c fixed number of rows 20120201
   '     Sheets("Report").Range("J" & k + 1).Value = "=ROUND(SUM(J" & j + 1 & ":J" & k & "),0)"
    
 End If
 
 'Insert NonRated Clients
 'WSName = "NonRated"
 'NRName = "NonRatedHeader"
 'If NRCheck(NRName) = True Then
  'Set First Row of NonRated Clients Section
 ' j = Range(NRName).Row
 ' k = j
 ' m = Application.WorksheetFunction.Match("ParentName", Sheets(WSName).Range("A1:D1"), 0)
 ' n = Application.WorksheetFunction.Match("SP_Rating", Sheets(WSName).Range("A1:D1"), 0)
 ' o = Application.WorksheetFunction.Match("SUM(SecuritizedValue)", Sheets(WSName).Range("A1:D1"), 0)
 
  'Loop Through NonRated Worksheet To Select Clients > 3% Concentration
 ' For i = 2 To GetLastRow(1, WSName)
  
 '  If Sheets(WSName).Cells(i, o).Value <> "" And _
 '  Sheets(WSName).Range("A" & i).Value <> "81020" And _
 '  Sheets(WSName).Range("E" & i).Value > 0.03 Then
 '   k = k + 1
 '   Sheets("Report").Rows(k).EntireRow.Insert
 '   Sheets("Report").Range("C" & k).Value = Sheets(WSName).Cells(i, n).Value
 '   Sheets("Report").Range("F" & k).Value = Sheets(WSName).Cells(i, m).Value
 '   Sheets("Report").Range("G" & k).Value = 0.03
 '   Sheets("Report").Range("H" & k).Value = Sheets(WSName).Range("E" & i).Value
 '   Sheets("Report").Range("I" & k).Value = Sheets(WSName).Cells(i, o).Value
 '   Sheets("Report").Range("J" & k).Value = "=IF((H" & k & "-G" & k & ")>0, (H" & _
 '  k & "-G" & k & ")*AgingBal,0)"
 '  End If
 ' Next i
'
'  'Format Percentage & Amount Cells
'  If k > j Then
'   Sheets("Report").Range("G" & j + 1 & ":H" & k).NumberFormat = "0.000%;[Red](0.000%)"
'   Sheets("Report").Range("I" & j + 1 & ":J" & k).NumberFormat = "_($* #,##0_);[Red]_($* (#,##0);_($* "" - ""??_);_(@_)"
'   Sheets("Report").Rows(j + 1 & ":" & k).RowHeight = 42
'   Sheets("Report").Rows(j + 1 & ":" & k).VerticalAlignment = xlBottom
'
'   ' reformat cells 20100803
'Sheets("Report").Range("c" & j + 1 & ":J" & k).Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
'    Selection.Borders(xlEdgeTop).LineStyle = xlNone
'    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
'    Selection.Borders(xlEdgeRight).LineStyle = xlNone
'    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'    Selection.Font.Bold = False
'    With Selection.Font
'        .Name = "Arial"
'        .Size = 10
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
'        .Bold = False
'    End With
'    With Selection
'
 '   End With
 '   Else
 '   Sheets("Report").Range("J" & k + 1).Value = "0"
 ' End If
 
  'Set NonRated Total Excess Amount

 '      If k > j + 1 Then
 '       Sheets("Report").Range("J" & k + 1).Value = "=ROUND(SUM(J" & j + 1 & ":J" & k & "),0)"
 '   Else
 '       Sheets("Report").Range("J" & k + 1).Value = "0"
 '   End If
 
    
End If
 
 'Enable Events
 Application.EnableEvents = True
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("WSLinkInserts", Err.Number, Err.Description)
End Sub
Sub InsertDelinquenciesData()
 'Initialize Variables
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim WSName As String
 Dim connStr As String
 Dim DBTable As String
 Dim SelClause As String
 Dim FromClause As String
 Dim WhereClause As String
 Dim HavingClause As String
 Dim SQLStmt As String

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Disable Events
 Application.EnableEvents = False
 
 'Set Connection String
 connStr = "DRIVER={Microsoft Access Driver (*.mdb)};Dbq=" & Range("AccessFile") & ";Uid="";Pwd="";"
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Set SELECT Clause
 SelClause = "SELECT Sum(qTrimLeaseID.SecuritizedValue) AS SumOfSecuritizedValue"
 
 'Set FROM Clause Database
 FromClause = " FROM qTrimLeaseID LEFT JOIN qxwrkLeaseDetail " & _
 "ON qTrimLeaseID.TrimLeaseID=qxwrkLeaseDetail.LeaseNbr"
   
 'Set WHERE Clause Criteria
 WhereClause = " WHERE qxwrkLeaseDetail.LeaseNbr Is Not Null"
 
 For i = 1 To 6
  'Set HAVING Clause Criteria
  Select Case i
   Case 1: HavingClause = " HAVING qxwrkLeaseDetail.SumOfFinalCurrent>0"
   Case 2: HavingClause = " HAVING qxwrkLeaseDetail.SumOfFinalAge01>0"
   Case 3: HavingClause = " HAVING qxwrkLeaseDetail.SumOfFinalAge02>0"
   Case 4: HavingClause = " HAVING qxwrkLeaseDetail.SumOfFinalAge03>0"
   Case 5: HavingClause = " HAVING qxwrkLeaseDetail.SumOfFinalAge04>0"
   Case 6: HavingClause = " HAVING qxwrkLeaseDetail.SumOfFinalAge05>0"
  End Select
 
  'SQL Query String
  SQLStmt = SelClause & FromClause & WhereClause & HavingClause & ";"
 
  'Create & Open New Recordset
  Set adoRS = New ADODB.Recordset
  adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
  'Insert Receivables Aging Values
  Select Case i
   'Insert Value - Portfolio Aging - Current (2a)
   Case 1: Range("AgingCur").Value = IIf(IsNull(adoRS.Fields("SumOfSecuritizedValue")), 0, Format(adoRS.Fields("SumOfSecuritizedValue"), "$#,##0.00;[Red]($#,##0.00)"))
   Case 2: Range("AgingCur").Value = Range("AgingCur") + IIf(IsNull(adoRS.Fields("SumOfSecuritizedValue")), 0, Format(adoRS.Fields("SumOfSecuritizedValue"), "$#,##0.00;[Red]($#,##0.00)"))
   'Insert Value - Portfolio Aging - 31_60 (2b)
   Case 3: Range("Aging31_60").Value = IIf(IsNull(adoRS.Fields("SumOfSecuritizedValue")), 0, Format(adoRS.Fields("SumOfSecuritizedValue"), "$#,##0.00;[Red]($#,##0.00)"))
   'Insert Value - Portfolio Aging - 61_90 (2c)
   Case 4: Range("Aging61_90").Value = IIf(IsNull(adoRS.Fields("SumOfSecuritizedValue")), 0, Format(adoRS.Fields("SumOfSecuritizedValue"), "$#,##0.00;[Red]($#,##0.00)"))
   'Insert Value - Portfolio Aging - 91_120 (2d)
   Case 5: Range("Aging91_120").Value = IIf(IsNull(adoRS.Fields("SumOfSecuritizedValue")), 0, Format(adoRS.Fields("SumOfSecuritizedValue"), "$#,##0.00;[Red]($#,##0.00)"))
   'Insert Value - Portfolio Aging - 120+ (2e)
   Case 6: Range("Aging120Up").Value = IIf(IsNull(adoRS.Fields("SumOfSecuritizedValue")), 0, Format(adoRS.Fields("SumOfSecuritizedValue"), "$#,##0.00;[Red]($#,##0.00)"))
  End Select
  
  'Close Recordset
  adoRS.Close
 Next i
  
 'Close Connection
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls

 'Enable Events
 Application.EnableEvents = True
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("InsertDelinquenciesData", Err.Number, Err.Description)
End Sub
Sub InsertChargeOffData()
 'Initialize Variables
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim WSName As String
 Dim connStr As String
 Dim DBTable As String
 Dim SelClause As String
 Dim FromClause As String
 Dim WhereClause As String
 Dim GroupByClause As String
 Dim HavingClause As String
 Dim SQLStmt As String
 Dim fso As Object

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Disable Events
 Application.EnableEvents = False
 
 'Set DealManager .ini File Name, Server Name & Database Name
 Set fso = CreateObject("Scripting.FileSystemObject")
 DMIniFileName = "C:\Users\" & UserNameWindows & "\AppData\Roaming\NorthBound Solutions\DealManager Suite\dealmanager.ini" 'Mid(ThisWorkbook.path, 1, InStrRev(ThisWorkbook.path, "\")) & "dealmanager.ini"
 If fso.FileExists(DMIniFileName) = False Then
  DMIniFileName = Left(ThisWorkbook.path, 3) & _
  "Program Files\NorthBound Solutions\DealManager Suite\dealmanager.ini"
  If fso.FileExists(DMIniFileName) = False Then
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
 DBTable = "ChargeOff"
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Set SELECT Clause
 SelClause = "SELECT Sum(ChargeOff.BookValue) AS SumOfBookValue"
 
 'Set FROM Clause Database
 FromClause = " FROM ChargeOff"
   
 'Set WHERE Clause Criteria
 WhereClause = " WHERE ChargeOff.DealID = 1"
 
 'Set GROUP BY Clause Criteria
 GroupByClause = " GROUP BY ChargeOff.AsOfDt"
 
 'Set HAVING Clause Criteria
 i = Month(DateAdd("d", -1, DateValue(Month(Now) & "-1-" & Year(Now))))
 HavingClause = " HAVING Month(ChargeOff.AsOfDt)=" & i
' HavingClause = " HAVING Month(ChargeOff.AsOfDt)=Month(DateAdd(""d"",-1,DateValue(Month(Now)&""-1-""&Year(Now))))"
 
 'SQL Query String
 SQLStmt = SelClause & FromClause & WhereClause & GroupByClause & HavingClause & ";"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
 'Insert Receivables Aging Values
 Do Until adoRS.EOF
  Range("DefUnits").Value = IIf(IsNull(adoRS.Fields("SumOfBookValue")), 0, Format(adoRS.Fields("SumOfBookValue"), "$#,##0.00;[Red]($#,##0.00)"))
  Exit Do
 Loop
  
 'Close Recordset
 adoRS.Close
 
 'Set SELECT Clause
 SelClause = "SELECT Sum(ChargeOff.RecoveryAmt) AS SumOfRecoveryAmt"
 
 'SQL Query String
 SQLStmt = SelClause & FromClause & WhereClause & GroupByClause & HavingClause & ";"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
 'Insert Receivables Aging Values
 Do Until adoRS.EOF
  Range("RecovAmt").Value = IIf(IsNull(adoRS.Fields("SumOfRecoveryAmt")), 0, Format(adoRS.Fields("SumOfRecoveryAmt"), "$#,##0.00;[Red]($#,##0.00)"))
  Exit Do
 Loop
  
 'Close Recordset & Connection
 adoRS.Close
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls

 'Enable Events
 Application.EnableEvents = True
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("InsertChargeOffData", Err.Number, Err.Description)
End Sub
Sub InsertDealTestColumns() 'This Procedure Sets Up the 4 Deal Test Columns
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
  
 'Missing Deal Data Message
 If ThisWorkbook.Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row = 1 Then
  MsgBox "No deal data has yet been entered in the 'Report' worksheet." & vbCr & vbCr & _
  "Please enter that data first.", vbCritical, gsAPP_NAME
 
  'Exit Procedure
  Exit Sub
 End If
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Determine Last Column & Row User By Inserted Deal Data
 i = Sheets("Report").Range("A1").SpecialCells(xlLastCell).Column + 5
 j = Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row
 If j = 1 Then
  j = 100
 End If
 
 'Insert Column Headers For Deal Test Data
 Sheets("Report").Cells(1, i).Value = "Test Name"
 Sheets("Report").Cells(1, i + 1).Value = "Test Result"
 Sheets("Report").Cells(1, i + 2).Value = "Difference"
 Sheets("Report").Cells(1, i + 3).Value = "Test Type"

 'Format Column Headers
 With Sheets("Report").Range(Cells(1, i).Address & ":" & Cells(1, i + 3).Address)
  .Font.Size = 8
  .Font.Bold = True
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlBottom
  .Borders(xlEdgeBottom).Weight = xlMedium
 End With
 
 'Set Column Width
 Sheets("Report").Columns(i).ColumnWidth = 20
 Sheets("Report").Columns(i + 1).ColumnWidth = 12
 Sheets("Report").Columns(i + 2).ColumnWidth = 15
 Sheets("Report").Columns(i + 3).ColumnWidth = 15
 
 'Format Column Data
 With Sheets("Report").Range(Cells(2, i).Address & ":" & Cells(j, i + 3).Address)
  .Font.Size = 8
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlTop
  .WrapText = True
 End With
End Sub
