Attribute VB_Name = "mABSappFunctions"
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
Sub sheetnm()

Call OnImportToDeal("cont", "9/30/2008")

End Sub

Function GetCurrentCell(sAddress As String, n As Integer, Separator As String)
' use this function to parse the relative string location and compare to
' the original string location captured at original configuration
' 1. what happens if the values change over time - should the static location
'    be updated each time?
' 2. also adjusting for servicer report types the increase from right to left each month,
'    which is defined as rSRtype in the Settings worksheet (=2)
'

Dim All As Variant
Dim sCol As String
Dim sRow As Long
Dim lSRType As Long
Dim lLastCol As Long

lSRType = ThisWorkbook.Sheets("settings").Range("rSRType").Value
lLastCol = GetMaxColumn("Report", 3)

All = Split(sAddress, Separator)

GetCurrentCell = All(n - 1)


If lSRType = "2" Then

'Determine if the column and row - checking to see if single or double columns)
    If IsNumeric(Mid(GetCurrentCell, 2, 1)) Then
        sCol = lLastCol
        sRow = Mid(GetCurrentCell, 2)
       '
    
        GetCurrentCell = ConvertToA1(sCol, sRow) 'sCol & sRow
     '
    Else
         sCol = lLastCol
        sRow = Mid(GetCurrentCell, 3)
        '
    
        GetCurrentCell = ConvertToA1(sCol, sRow) ' sCol & sRow
        '
    End If
Else
    All = Split(sAddress, Separator)
    GetCurrentCell = All(n - 1)
End If

leave:
End Function
Function UpdateFormula(sWS As String, sCell As String)
' use this function to update the formula on KDI-CI worksheet F




End Function


Function GetMaxColumn(sWS As String, lRow As Long)
' determine the maximum column to obtain current values

Dim ws As Worksheet
Dim l As Long

Set ws = ThisWorkbook.Sheets(sWS)

l = ws.Cells(lRow, 255).End(xlToLeft).Column

GetMaxColumn = l

End Function
Function ConvertToA1(col, Row)
ConvertToA1 = IIf(col > 26, Chr(64 + Int((col - 1) / 26)), "") & Chr(65 + (col - 1) Mod 26) & Row

End Function

Sub callfunction()

' this function checks the dynamic reference to the cell and is used
' to compare to the CI form to detect the current value of CI's if applicable


Dim ws As Worksheet
Dim sAddress As Range
Dim i As Integer
Dim m As Integer
Dim cell As String
Dim r As Long
Dim iSRType As Integer


Set ws = ThisWorkbook.Sheets("KDI-CI")
iSRType = ThisWorkbook.Sheets("Settings").Range("rSRtype").Value

r = ws.Range("A60000").End(xlUp).Row

Set sAddress = ws.Range("F2:F" & r)


For i = 2 To r
If ws.Cells(i, 6).Value = "" Then ' change from 6 column to get the cell reference, not value 20090430
    GoTo NextOne
Else

cell = GetCurrentCell(ws.Cells(i, 6).Formula, 2, "!")
ws.Cells(i, 7).Value = cell
End If

If iSRType = "2" Then

    ws.Cells(i, 6).Formula = "=+Report!" & cell
Else
End If


NextOne:

Next i




End Sub

Sub oldOnImportToDeal(connStr As String, effDate As Date)

Dim l As Long
Dim r As Worksheet
Dim MEDt As Date
Dim lCol As Long
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim DMCreationDate As Date
 Dim DMEffectiveDate As Date
' Dim cString As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim SQLStmt As String
 Dim fso As Object
 Dim j As Integer
 Dim i As Integer
 

Set r = ThisWorkbook.Sheets("Report")
r.Activate
' step1 get the date
l = r.Range("d3").End(xlToRight).Column
MEDt = r.Cells(3, l).Value
MEDt = CDate(MEDt)

MEDt = effDate


WSName = "Report" 'rSRSheetNm



' step 2 compare the dates
If MEDt <> effDate Then
    MsgBox "The dates do not match - please try again", vbCritical, gsAPP_NAME
    Exit Sub
End If

' step 3 get the most recent values and move to worksheet
l = ThisWorkbook.Sheets("Report").Range("rCIind").Column 'End(xlUp).row
j = r.Range("d3").End(xlToRight).Column



' For i = 6 To 200 ' GetLastRow(j, "Report")
 ' figure out of the CI Ind column has a '1', in which case there is a CI
 
'    If r.Cells(i, l).Value = "1" Then
'        r.Cells(i, l + 6).Value = r.Cells(i, 13).Value ' not j
'        r.Cells(i, l + 7).Value = r.Cells(i, 13).Address 'not j for Auto
'    End If
        

'Next i


' step 4 update to db
'Call ExportCIDataToDM(connStr, effDate)
r.Activate
' check for the values of 1 in column ia




 'Set Screen Display Controls
 Call SetScreenControls
 
' set connection string and table name
 
  DBTable = "DealInputValues"
 
 'Set Data Creation & Effective Date Values (Effective Date Is Last Inputs Effective Date)
 DMCreationDate = Now
 DMEffectiveDate = effDate

  'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 connStr = "DRIVER={SQL SERVER};Server=LENOVO-B02A0E29;initial catalog=DealManager;Integrated Security=SSPI"
connStr = "DRIVER={SQL SERVER};Server=LENOVO-B02A0E29;Database=DealManager;ReadOnly=False;"
  
 adoConn.Open connStr
   


 k = 1
 For i = 3 To 200 ' GetLastRow(j, "Report")
 ' figure out of the CI Ind column has a '1', in which case there is a CI
 
    If r.Cells(i, l).Value = "1" Then
    
        Application.StatusBar = "Exporting " & WSName & " data from row " & i

         'Set Basic INSERT INTO Clause
        InsertClause = "INSERT INTO " & DBTable & "("
 
        'Set Values Portion of INSERT INTO Clause

        InsertClause = InsertClause & "DealInputID, CreateDt, Comments, EffectiveDt, Value, CellReference) VALUES ("
         InsertClause = InsertClause & "'" & Sheets(WSName).Cells(i, l + 1).Value & "', '" & DMCreationDate & "', " & "'No Comments'" & _
        ", '" & DMEffectiveDate & "', '" & Sheets(WSName).Cells(i, l + 8).Value & "', '" & Sheets(WSName).Cells(i, l + 8).Formula & "')"

  
        'SQL Query String
        SQLStmt = InsertClause & ";"
 
         'Create & Open New Recordset
         Set adoRS = New ADODB.Recordset
         

         adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
    End If
Next i
 
 'Close Connection
 adoConn.Close
 
 
MsgBox "The data for " & effDate & " has been loaded into DealManager!", vbInformation, gsAPP_NAME


End Sub


Sub GetCIValues()

Dim l As Long
Dim i As Integer
Dim j As Integer
Dim r As Worksheet
Dim k As Worksheet



Set r = ThisWorkbook.Sheets("Report")
Set k = ThisWorkbook.Sheets("KDI-CI")


r.Activate
' check for the values of 1 in column ia

l = ThisWorkbook.Sheets("Report").Range("rCIind").Column 'End(xlUp).row
j = r.Range("d3").End(xlToRight).Column



 For i = 6 To 200 ' GetLastRow(j, "Report")
 ' figure out of the CI Ind column has a '1', in which case there is a CI
 
    If r.Cells(i, l).Value = "1" Then
        r.Cells(i, l + 6).Value = r.Cells(i, j).Value
        r.Cells(i, l + 7).Value = r.Cells(i, j).Address
    End If
        

Next i
'Sheets("KDI-CI").Range("C" & k).Value = Sheets("Report").Cells(i, j + 2) '(C) Difference

End Sub

Sub oldExportCIDataToDM(connStr As String, effDate As Date) 'This Procedure Exports the CI Data to the DealManager Database
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
' Dim cString As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim SQLStmt As String
 Dim fso As Object
 Dim d As Date
 Dim r As Worksheet
 Dim l As Long
 Dim j As Integer



 'Turn Error Handler On
 On Error GoTo ErrorHandler
 

Set r = ThisWorkbook.Sheets("Report")


r.Activate
' check for the values of 1 in column ia

l = ThisWorkbook.Sheets("Report").Range("rCIind").Column 'get column of Indicator
j = r.Range("d3").End(xlToRight).Column ' get the most recent date
''d = r.Cells(3, j).Value
'd = CDate(d)


 'Set Screen Display Controls
 Call SetScreenControls
 
' set connection string and table name
 
  DBTable = "DealInputValues"
 
 'Set Data Creation & Effective Date Values (Effective Date Is Last Inputs Effective Date)
 DMCreationDate = Now
 DMEffectiveDate = effDate

  'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   



 For i = 3 To 200 ' GetLastRow(j, "Report")
 ' figure out of the CI Ind column has a '1', in which case there is a CI
 
    If r.Cells(i, l).Value = "1" Then
    
        Application.StatusBar = "Exporting " & WSName & " data from row " & i

         'Set Basic INSERT INTO Clause
        InsertClause = "INSERT INTO " & DBTable & "("
 
        'Set Values Portion of INSERT INTO Clause

        InsertClause = InsertClause & "DealInputID, CreateDt, Comments, EffectiveDt, Value, CellReference) VALUES ("
         InsertClause = InsertClause & "'" & Sheets(WSName).Cells(i, l + 1).Value & "', '" & DMCreationDate & "', " & "'No Comments'" & _
        ", '" & DMEffectiveDate & "', '" & Sheets(WSName).Cells(i, l + 6) & "', '" & Sheets(WSName).Cells(i, l + 7) & "')"

  
        'SQL Query String
        SQLStmt = InsertClause & ";"
 
         'Create & Open New Recordset
         Set adoRS = New ADODB.Recordset
         adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
    End If
Next i
 
 'Close Connection
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Or _
 Right(CommandBars.ActionControl.Caption, 13) = "CI Data to DM" Then
  Call ClearScreenControls
 End If

 'Completion Message
 If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
  MsgBox "Export of the Deal Test Data Has Been Completed.", vbExclamation, gsAPP_NAME
 ElseIf Right(CommandBars.ActionControl.Caption, 13) = "CI Data to DM" Then
  MsgBox "Export of the KDI & CI Data Has Been Completed.", vbExclamation, gsAPP_NAME
 End If
 
 'Skip Error Handler
' Exit Sub

'MsgBox "i did it"

'Error Handler
ErrorHandler:
'Call ErrorLogRecord("ExportDataToDM", Err.Number, Err.Description)
End Sub
Sub OnImportToDeal(connStr As String, effDate As Date) 'This Procedure Exports the Deal Test, KDI & CI Data to the DealManager Database
 'Initialize Variables
 '*************adapted from ExportDatatoDM
 
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim DMCreationDate As Date
 Dim DMEffectiveDate As Date
 Dim DBTable As String
 Dim InsertClause As String
 Dim SQLStmt As String
 Dim fso As Object

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
 
 
  WSName = ThisWorkbook.Sheets("KDI-CI").Name
  
  ' make sure the updated KDI-CI values are in worksheet
  Call callfunction
  
 'Set Screen Display Controls
 Call SetScreenControls
 
 'First update the Test Worksheet **************************
' If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
'  Call UpdateTestsWS
' End If
 
 'Determine Data Worksheet Name
'deleted**********************************************
 
 'Check For Data To Be Exported

  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
  'Exit Procedure
'  Exit Sub

 
 'Set DealManager .ini File Name, Server Name & Database Name
 Set fso = CreateObject("Scripting.FileSystemObject")
 DMIniFileName = "C:\Users\" & UserNameWindows & "\AppData\Roaming\NorthBound Solutions\DealManager Suite\dealmanager.ini"
' Mid(ThisWorkbook.path, 1, InStrRev(ThisWorkbook.path, "\")) & "dealmanager.ini"
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
 

 'Set DB Table Name NOT REQUIRED FOR AUTO IMPORT
 'If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
 ' DBTable = "DealTests"
 'Else
  DBTable = "DealInputValues"
 'End If
 
 'Set Data Creation & Effective Date Values (Effective Date Is Last Inputs Effective Date)
 DMCreationDate = Now
' If Right(CommandBars.ActionControl.Caption, 22) = "Upload Test Data to DM" Then
'  DMEffectiveDate = DateValue(Month(Now) + 1 & " 1, " & Year(Now)) - 1
' Else
'  i = Sheets("Inputs").Range("A1").End(xlToRight).Column
'  i = Application.WorksheetFunction.Match("Effective Date", Sheets("Inputs").Range("A1:" & Sheets("Inputs").Cells(1, i).Address), 0)
'  DMEffectiveDate = Sheets("Inputs").Cells(2, i)
' End If
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Determine Last Row of Worksheet Data
 j = GetLastRow(1, WSName)
 
 For i = 2 To j
    
 If Sheets(WSName).Range("F" & i) = "" Then
    GoTo NextOne
Else
 
  'Set Basic INSERT INTO Clause
  InsertClause = "INSERT INTO " & DBTable & "("
 
  'Set Values Portion of INSERT INTO Clause
 '****************************
   InsertClause = InsertClause & "DealMetricID, DealID, CreateDt, Comments, EffectiveDt, Value, Source) VALUES ("
   InsertClause = InsertClause & "'" & Sheets(WSName).Range("A" & i) & "', '" & "1" & "', '" & DMCreationDate & "', " & "'No Comments'" & _
   ", '" & effDate & "', '" & Sheets(WSName).Range("F" & i) & "', '" & Sheets(WSName).Range("f" & i).Formula & "')"
 '************
  
  'SQL Query String
  SQLStmt = InsertClause & ";"
 
  'Create & Open New Recordset
  Set adoRS = New ADODB.Recordset
  adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
End If
NextOne:

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
'Call ErrorLogRecord("ExportDataToDM", Err.Number, Err.Description)
End Sub


