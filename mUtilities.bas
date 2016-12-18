Attribute VB_Name = "mUtilities"
Option Explicit
'Initialize Module Level Variables
Dim sh As Worksheet
Dim WSName As String
Dim i As Long
Dim j As Integer
Public Sub FindTopPoolIDs()
' used to generate an ordered list of pool ids by creating unique list of Pool IDs from the
' assetdata worksheet and placing the unique list in descending order on the settings worksheet b25

Dim rge As Range
Dim rge2 As Range
Dim sh2 As Worksheet
Dim col As Integer
Dim Lastcol As String
Dim wsBegin As Worksheet


Set sh = ThisWorkbook.Sheets("assetdata")
Set sh2 = ThisWorkbook.Sheets("settings")
Set wsBegin = ActiveWorkbook.ActiveSheet


' get the last column to serve as holding pattern for unique list of pool ids (j+2)
j = sh.Range("A1").End(xlToRight).Column
j = j + 2
 Lastcol = sh.Cells(1, j).Address
 
' get the column number
col = Sheets("AssetData").Rows(1).Find(what:="Pool ID", LookIn:=xlValues, lookat:=xlWhole, MatchCase:=False).Column
i = GetLastRow(col, "AssetData")

' activite the assetdata worksheet and create unique list
sh.Activate
sh.Cells(1, col).Select
 Range(Selection, Selection.End(xlDown)).Select
    Selection.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range(Lastcol), unique:=True 'Range("Ad2"), Unique:=True

' Move the list to cell 'b25' on the settings worksheet
 sh.Cells(1, j).Select
 Range(Selection, Selection.End(xlDown)).Cut sh2.Range("b25")
 'Selection.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("AB2"), Unique:=True 'sh2.Range("b25"), Unique:=True
'End With

' order in descending order
sh2.Activate
 Range("B25").Select
  Range(Selection, Selection.End(xlDown)).Select
  With Selection
    .sort Key1:=Range("b25"), Order1:=xlDescending
    End With

' activate the original worksheet
wsBegin.Activate

End Sub
Sub InsertNewWSRow(NewRowNum As Integer, WSName As String)
 'Insert New Row, If Necessary
 If NewRowNum > 2 Then
  Sheets(WSName).Rows(NewRowNum).EntireRow.Insert
  Sheets(WSName).Range("A" & NewRowNum & ":E" & NewRowNum).Borders(xlEdgeTop).Weight = xlThin
 End If
End Sub
Sub ReportFirstCell()
 'Select First Cell in Report Worksheet
 Sheets("Report").Select
 Sheets("Report").Range("A2").Select
 ActiveWindow.ScrollRow = 1
 ActiveWindow.ScrollColumn = 1
End Sub
Sub SortKDICITable()
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set KDI-CI Worksheet Name
 WSName = "KDI-CI"
 
 'Determine Last Row of KDI-CI Worksheet Data
 j = GetLastRow(1, WSName)
 
 'Select KDI-CI Worksheet
 Sheets(WSName).Select
 
 'Select All Data, Excluding Vlookup Columns
 Range("A2:G" & j).Select
 
 'Sort Selected Range
 Selection.sort Key1:=Range("B2"), Order1:=xlDescending, Key2:=Range("A2"), Order2:=xlAscending, Header:=xlNo, _
 OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
 DataOption2:=xlSortTextAsNumbers
 
 'Select First Cell
 Range("A2").Select
 
 'Select Report Worksheet
 Sheets("Report").Select
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("SortKDICITable", Err.Number, Err.Description)
End Sub
Sub ErrorLogRecord(ProcName As String, ErrorNum As Integer, ErrorDescr As String)
 If Sheets.Count > 1 Then
  'Set Screen Display Controls
  Call SetScreenControls
 
  'Determine Last Row of Error Log
  i = GetLastRow(1, "Error Log") + 1
 
  'Insert New Row, If Necessary
  WSName = "Error Log"
  Call InsertNewWSRow(i, WSName)
    
  'Insert Error Detail In Error Log Worksheet
  Sheets(WSName).Range("A" & i).Value = "MS Run-time error " & ErrorNum '(A) Error Number
  Sheets(WSName).Range("B" & i).Value = ErrorDescr '(B) Error Description
  Sheets(WSName).Range("C" & i).Value = ProcName '(C) VBA Procedure Error Occurred In
  Sheets(WSName).Range("D" & i).Value = Now() 'Format(Now(), "m/d/yyyy hh:mm:ss") '(D) Error Time
  Sheets(WSName).Range("E" & i).Value = ThisWorkbook.Name '(E) Filename
 
  'Show Last Row of Error Log Worksheet
  Sheets("Error Log").Select
  Sheets("Error Log").Range("A" & i).Select
 
  'Clear Status Message & Turn On Screen Updating
  Call ClearScreenControls
  
  'Error Message
  MsgBox "A Microsoft error has just been generated." & vbCr & vbCr & _
  "Review the 'Error Log' worksheet for more details.", vbCritical, gsAPP_NAME
 End If
End Sub
Sub DisableLinks()
 'Initialize Variable
 Dim SRLinks As Variant
 
 'Check For Linked Files
 SRLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
 If Not IsEmpty(SRLinks) Then
  'ActiveWorkbook.UpdateLinks = xlUpdateLinksNever
 End If
End Sub
Sub Temp_DeleteNewSheets()
 'Delete All Worksheets But Report
 Application.DisplayAlerts = False
 For Each sh In ThisWorkbook.Sheets
  If sh.Name <> "Report" Then
   'Delete Worksheet
   sh.Delete
  End If
 Next sh
 Application.DisplayAlerts = True
End Sub
Sub TestCode()
MsgBox ActiveCell.Name.Name
End
 'Initialize Variable
 Dim SRLinks As Variant
 
 'Check For Linked Files
 SRLinks = ActiveWorkbook.LinkSources(xlExcelLinks)
 If Not IsEmpty(SRLinks) Then
  MsgBox "Links"
 Else
  MsgBox "No Links"
 End If
 
 End
End Sub

Sub RemNamedRanges()
     
    Dim nm As Name
     
    On Error Resume Next
    For Each nm In ActiveWorkbook.Names
        nm.Delete
    Next
    On Error GoTo 0
     
End Sub

Sub backgroundcolor()

ThisWorkbook.Sheets("Report").Activate

Cells.Interior.ColorIndex = 2
End Sub


