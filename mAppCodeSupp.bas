Attribute VB_Name = "mAppCodeSupp"
' 1. added "Rated" range and brought in the code from SRA that inserts appropriate number
'     of rows.



Sub GenerateReport()
 
 
'his procedure will generate the Servicer Report in MS Excel, using
' data from the current period.
' If the user clicks the "Trial Report" button on the MAINSRA form, then a "1"
'  will pass to the procedure and the Report will print as links.
' If the user clicks the "Final Report" button on the MAINSRA form, then a "2"
'  will pass to the procedure and the Report will print as values, only, with all
'  sheets other than the Report removed.

' If the user clicks the "Preliminary Report" button on the MAINSRA form, then a "3"
' will pass to the procedure and the report will print the report in the same way as the
' Trial report, except that the values will include all Active Pools and the Pool selected
' on the

'20160510 changed concentration allowable from 3% to 4%

 ' This procedure exports data to Excel
 
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWBNew As Excel.Workbook
    Dim wsData As Excel.Worksheet
    Dim wsRated As Excel.Worksheet
     Dim wsNonRated As Excel.Worksheet
    Dim wsReport As Excel.Worksheet
    Dim wsTests As Excel.Worksheet
    Dim wsInputs As Excel.Worksheet
    Dim wsCash As Excel.Worksheet
    Dim wsGovt As Excel.Worksheet
     Dim iCol As Integer
    Dim iRow As Integer
    Dim lRow As Long
    Dim path As String
    Dim iCt As Integer
    Dim j As Integer
    Dim k As Integer
   ' Dim i As Integer
    Dim l As Integer
    Dim cell As String
    Dim iFail As Integer
    Dim iPool As Integer
    Dim sSql As String
    Dim dMEdt As Date
    Dim dPriorMEDt As Date
    Dim dMaxDtDealTest As Date
    Dim sFolderName As String
    Dim newpath As String
    Dim cBalance As Variant
    
    
 
        Set xlWB = ThisWorkbook ' .Workbooks.Open(path & "\SR.xls", True, False)
        Set wsData = xlWB.Sheets("Data")
        Set wsRated = xlWB.Sheets("Rated")
        Set wsReport = xlWB.Sheets("Report")
        Set wsTests = xlWB.Sheets("Tests")
        Set wsInputs = xlWB.Sheets("Inputs")
        Set wsNonRated = xlWB.Sheets("NonRated")
      
        
     ' the path of the Securitization folder
    path = "c:\" '"N:\NorthboundSolutions\Securitization Reports"
    
    dMEdt = ThisWorkbook.Sheets("Settings").Range("medt").Value
    
    sFolderName = Year(dMEdt)
    
    If Month(dMEdt) < 10 Then
        sFolderName = sFolderName & "0" & Month(dMEdt)
    Else
        sFolderName = sFolderName & Month(dMEdt)
    End If
    
    
    sFolderName = sFolderName & Day(dMEdt)
    
        
    
    
    newpath = path & "\" & sFolderName
    ' check to see if the directly exists; if not, create
    If Dir(path, vbDirectory) = "" Then
        MkDir sourcelocation
    End If
    
     
            

    ' set the count of the Rated clients.  This determines the
    ' number of rows to insert.
    
    iRow = wsRated.Range("b65000").End(xlUp).Row
    iCt = iRow - 1
' check to see if any data exists
    If iCt < 2 Then
        MsgBox "There doesn't appear to be any data for Rated Clients! " & vbCrLf & vbCrLf & _
            "Please check the logic and run again!", vbOKOnly + vbInformation, "SRA Help"
        GoTo Exit_Handled:
    End If
          
 ' Replace the formulas for values
     wsRated.Range("D2:e" & iRow).Copy
        wsRated.Range("D2").PasteSpecial xlValues
          
 wsReport.Activate
 wsReport.Range("Rated").Select
 'With Selection
 k = wsReport.Range("rated").Row + 2

' INsert a row for each of the items
For j = k To k + iCt - 1

    wsReport.Range("a" & j).Offset(0).EntireRow.Insert
 
Next j
wsRated.Activate

' copy the information from the Rated worksheet to the Report worksheet
wsRated.Range("a1").Select
wsRated.Range("c2:c" & iRow).Copy wsReport.Range("rated").Offset(2, 1) 'S&P Rating OK
wsRated.Range("b2:b" & iRow).Copy wsReport.Range("rated").Offset(2, 4) 'Parent Name
wsRated.Range("e2:e" & iRow).Copy wsReport.Range("rated").Offset(2, 6) 'Actual concentration
wsRated.Range("d2:d" & iRow).Copy wsReport.Range("rated").Offset(2, 7) 'Actual Securitized Value
wsReport.Range("G" & k & ":G" & k + iCt - 1).Value = ".04000" ' Copy wsReport.Range("rated").Offset(2, 5) 'Allowable Limit
'End With

wsReport.Activate


wsReport.Range("g94:g" & k + iCt).Style = "Percent"
wsReport.Range("g94:g" & k + iCt).NumberFormat = "0.000%" 'Select 'rated").Offset(2, 5).Select

wsReport.Range("h94:h" & k + iCt).Style = "Percent"
wsReport.Range("h94:h" & k + iCt).NumberFormat = "0.000%" 'Select 'rated").Offset(2, 5).Select

wsReport.Range("i94:i" & k + iCt).Style = "Currency"
wsReport.Range("i94:i" & k + iCt).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"


k = wsReport.Range("rated").Row + 2
For j = k To k + iCt - 1
   ' wsReport.Range("rated").Offset(j, 8).Select
    'wsReport.Cells(j, 10).Select
    'Selection.Formula = "=IF(((H" & j & "-G" & j & ")*I" & j & ")>0, (H" & j & "-G" & j & ")*I" & j & ", 0)"
wsReport.Cells(j, 10).Formula = "=IF(((H" & j & "-G" & j & ")*I" & j & ")>0, (H" & j & "-G" & j & ")*I" & j & ", 0)"
Next

If iCt > 1 Then
    wsReport.Range("rExcessConc").Formula = "=round(sum(J" & k & ":J" & k + iCt - 1 & "),0)"
    wsReport.Range("J" & k & ":J" & k + iCt - 1).Style = "CURRENCY"
    wsReport.Range("J" & k & ":J" & k + iCt - 1).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    wsData.Visible = xlSheetVisible
  
wsReport.Range("c" & k & ":J" & k + iCt - 1).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
 With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
Else
       Sheets("Report").Range("J" & k + 1).Value = "0"
    End If
' next, want to check all tests and list on the tests page.
lRow = wsReport.Range("l65536").End(xlUp).Row

j = 7
l = 2
For k = j To j + lRow

    If wsReport.Cells(k, 12).Value = "1" Then
        wsReport.Range("L" & k & ":" & "o" & k).Copy
        wsTests.Cells(l, 1).PasteSpecial xlValues
        
    '    wsReport.Cells(k, 10).Copy
    '    wsTests.Cells(l, 3).PasteSpecial xlValues
        cell = wsReport.Cells(k, 10).Address 'ActiveCell.Address
        wsTests.Cells(l, 5).Value = cell
        
        l = l + 1
        
    End If
    
Next k
' check to see if there is a fail indicator value of 1, which
' means a potential violation exists.

' ******************************   Non Rated Exposure

' ------> now check to see if there are any nonrated excess exposures  20080117

iRow = wsNonRated.Range("a65000").End(xlUp).Row

 If iRow > 1 Then
    iCt = iRow - 1 'determine how many nonrateds are on the list
    
    
     ' Replace the formulas for values
     wsNonRated.Range("D2:e" & iRow).Copy
        wsNonRated.Range("D2").PasteSpecial xlValues
          
     wsReport.Range("NonRated").Select 'select the row that starts the nonrated exposure
 'With Selection
    k = wsReport.Range("Nonrated").Row + 2 ' determine the row to begin inserting rows.

    ' Insert a row for each of the items
    For j = k To k + iCt - 2

    wsReport.Range("a" & j).Offset(0).EntireRow.Insert
    Next j

        wsNonRated.Activate
        
 ' ----
        wsNonRated.Range("a1").Select
        wsNonRated.Range("b2:b" & iRow - 1).Copy wsReport.Range("nonrated").Offset(2, 4) 'Parent Name
        wsNonRated.Range("e2:e" & iRow - 1).Copy wsReport.Range("nonrated").Offset(2, 6) 'Actual concentration
        wsNonRated.Range("d2:d" & iRow - 1).Copy wsReport.Range("nonrated").Offset(2, 7) 'Actual Securitized Value
        wsReport.Range("G" & k & ":G" & k + iCt - 1).Value = ".04000"
        
        
        wsReport.Activate
               
        wsReport.Range("g" & k & ":g" & k + iCt - 1).Style = "Percent"
        wsReport.Range("g" & k & ":g" & k + iCt - 1).NumberFormat = "0.0000%" 'Select 'rated").Offset(2, 5).Select
        
        wsReport.Range("h" & k & ":h" & k + iCt - 1).Style = "Percent"
        wsReport.Range("h" & k & ":h" & k + iCt - 1).NumberFormat = "0.0000%" 'Select 'rated").Offset(2, 5).Select
        
        wsReport.Range("i" & k & ":i" & k + iCt - 1).Style = "Currency"
        wsReport.Range("i" & k & ":i" & k + iCt - 1).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        

        k = wsReport.Range("nonrated").Row + 2
        For j = k To k + iCt - 1
           ' wsReport.Range("rated").Offset(j, 8).Select
            'wsReport.Cells(j, 10).Select
            'Selection.Formula = "=IF(((H" & j & "-G" & j & ")*I" & j & ")>0, (H" & j & "-G" & j & ")*I" & j & ", 0)"
         
        wsReport.Cells(j, 10).Formula = "=ROUND(((H" & j & "-G" & j & ")*I12),2)"
        wsReport.Cells(j, 10).Style = "CURRENCY"
        Next
        
        If iCt <= 1 Then
           ' wsReport.Range("EC_NonSpecial").Formula = "=round(sum(J" & k & ":J" & k & "),0)"
           wsReport.Range("EC_NonSpecial").Formula = "0"
           
         Else
            wsReport.Range("EC_NonSpecial").Formula = "=round(sum(J" & k & ":J" & k + iCt - 2 & "),0)"
        End If
        wsReport.Range("j" & k & ":j" & k + iCt - 1).Style = "CURRENCY"
        wsReport.Range("j" & k & ":j" & k + iCt - 1).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
            wsData.Visible = xlSheetVisible
        

wsReport.Range("c" & k & ":J" & k + iCt - 1).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Else
'this means that there was no excess exposure
 wsReport.Range("EC_NonSpecial").Formula = "=round(sum(J" & k & ":J" & k & "),0)"
wsReport.Range("EC_NonSpecial").Offset(1, 0).Formula = "0"
End If

Exit_Handled: ' Label to resume after error.
    
   
    Exit Sub 'Function

 End Sub
 
 Sub printreport()
  
   
   newpath = newpath & "\ULI 2007-1 Servicer Report " & Year(dMEdt)
   If Month(dMEdt) < 10 Then
        newpath = newpath & "0" & Month(dMEdt) & Day(dMEdt) & ".xls"
    Else
        newpath = newpath & Month(dMEdt) & Day(dMEdt) & ".xls"
    End If
    
    
    xlWBNew.SaveAs FileName:=newpath
    
    With xlWBNew.Sheets("Report")
        .Cells.Copy 'columns("A:J").Copy
       ' .Cells.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   ' xlWBNew.Sheets("Report").Columns("A:j").Copy ' ion.Copy
   ' xlWBNew.Sheets("Report").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
   ' xlWBNew.Sheets("Report").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
   '     SkipBlanks:=False, Transpose:=False
    
  '  xlWBNew.Sheets("Report").Activate
        .Cells.Range("a1").Activate
     '  .Cells.Find(What:="#REF", After:=ActiveCell, LookIn:=xlFormulas, _
     '   LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
      '  MatchCase:=False, SearchFormat:=False).Activate
   ' ActiveCell.Replace What:="#REF", Replacement:="0", LookAt:=xlPart, _
   '     SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
   '     ReplaceFormat:=False
   ' Selection.FindNext(After:=ActiveCell).Activate
   'ActiveCell.Replace What:="#REF", Replacement:="0", LookAt:=xlPart, _
    '    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    '    ReplaceFormat:=False
    'Selection.FindNext(After:=ActiveCell).Activate
    'Selection.FindNext(After:=ActiveCell).Activate
     .Columns("L:U").Delete Shift:=xlToLeft
    End With
    
    
     
 ' Delete all remaining worksheets (not named Report)
 
  
   

    
    ' Save the workbook
    ActiveWorkbook.Save
    
    ' set up the workbook for printing
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$2:$7"
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = "$B$1:$J$210"
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
    End With
  '  ActiveWindow.SelectedSheets.PrintPreview
    
    
 
 ' copy the entire worksheet
 
 ' paste special as values the entire worksheet
 
 ' trim values to shown (?)
'  Call GenerateFundRequest("6/20/2007")
  

 ' save the file as ULI 2007-1 Report & MM & - & YYYY
 
 
' End If
 
 
 
 

Exit_Handled: ' Label to resume after error.
     DoCmd.Hourglass (False)
     xlApp.ScreenUpdating = True
    Exit Sub 'Function
Err_Handler:
    DoCmd.Hourglass (False)
    xlApp.Visible = True
    MsgBox Err.Number & Err.Description
   Select Case Err.Number
      Case 9999                        ' Whatever number you anticipate.
          Resume Next                  ' Use this to just ignore the line.
      Case 999
          Resume Exit_Handled       ' Use this to give up on the proc.
      Case Else                        ' Any unexpected error.
         ' Call LogError(Err.Number, Err.Description, sproc)
         ' Resume Exit_Handled
      End Select

End Sub


Sub MoveObligors()


Dim wsObligors As Worksheet
Dim wsRated As Worksheet
Dim wsNonRated As Worksheet
Dim iO As Integer
Dim iNR As Integer
Dim iR As Integer
Dim j As Integer
Dim sSV As String








Set wsObligors = ThisWorkbook.Sheets("Obligors")
Set wsRated = ThisWorkbook.Sheets("Rated")
Set wsNonRated = ThisWorkbook.Sheets("NonRated")


' select the Obligors worksheet
wsObligors.Activate

' find the last row
'Set iO = wsObligors.Range("b50000").End(xlUp).Row
'Set iNR = wsNonRated.Range("b50000").End(xlUp).Row




' move rated to the Rated worksheet



End Sub



Sub mDeleteConcRows()
'
' mDeleteConcRows Macro
' MThis procedure checks to makes ure that no concentration rows
'    already exist within the Report worksheet.  If they do exist, then they are removed.

Dim i As Integer
Dim j As Integer
Dim wsR As Worksheet


Set wsR = ThisWorkbook.Sheets("Report")

i = wsR.Range("Rated").Row + 2 ' this is the first concentration item
j = wsR.Range("rExcessConc").Row - 1 ' this is the last concentration
' check to make sure the rows haven't already been deleted
If i > j Then
    GoTo NextOne
Else
    Rows(i & ":" & j).Select
    Selection.Delete Shift:=xlUp
   
   
NextOne:
'this checks the NonRated Concentrations
i = wsR.Range("NonRated").Row + 2
j = wsR.Range("EC_NonSpecial").Row - 1

If i > j Then
    Exit Sub
End If
      Rows(i & ":" & j).Select
    Selection.Delete Shift:=xlUp
End If

End Sub


Sub UpdateFox()
'
' UpdateFox Macro
' Macro recorded 5/4/2010 by landfall
'
' modifiedd 11/5/12 to skip over errors if they occur

'
On Error Resume Next
    Sheets("Obligors").Select
    ' look for 80227
    Columns("A:A").Select
    Selection.Find(what:="80227", After:=ActiveCell, LookIn:=xlFormulas, _
        lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    Cells(ActiveCell.Row, 6) = "1"
    ' now look for 80901
    Columns("A:A").Select
    Selection.Find(what:="80901", After:=ActiveCell, LookIn:=xlFormulas, _
        lookat:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Select
    Cells(ActiveCell.Row, 6) = "1"
    
    Cells(1, 6).Formula = "=sum(F2:F20000)"
    If Cells(1, 6).Value > 1 Then
    
            'Selection.End(xlUp).Select
            Range("f2").Select
            Selection.End(xlDown).Select
            
           
            Cells(ActiveCell.Row, 4).Select 'ange("E96").Select
             Selection.Copy
            Cells(ActiveCell.Row, 6).Select
            Selection.End(xlDown).Select
            
            ' find the next one and paste value by adding to existing value
            Cells(ActiveCell.Row, 4).Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
                False, Transpose:=False
            
            Range("f2").Select
            Selection.End(xlDown).Select
            Cells(ActiveCell.Row, 4) = "0"
            
            Sheets("Report").Select
    Else
    End If
    
On Error GoTo 0

End Sub
