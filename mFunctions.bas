Attribute VB_Name = "mFunctions"
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public vDMValueID As Variant ' use this to update DMetricValue id on ExportData routine
Public vDMid As Variant
Public vValue As Variant
Public vDealID As Variant ' hold dealid fo insert into dmvhist
Public vCreateDt As Variant ' hold create dt for insert into dmvhist
Public vEffectiveDt As Variant
Public vComments As Variant
Public vSource As Variant

'Initialize .ini File Data Extraction Function
Private Declare Function GetPrivateProfileStringA Lib "Kernel32" (ByVal strSection As String, _
ByVal strKey As String, ByVal strDefault As String, ByVal strReturnedString As String, _
ByVal lngSize As Long, ByVal strFileNameName As String) As Long
'Declare Function To Determine System Metric For User's Screen Width
Declare Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex&) As Long
Function UserNameWindows() As String
    ' get the user name, which identifies the user folder to locate the dealmanager.ini file
    ' "C:\Users\" & UserNameWindows & "\AppData\Roaming\NorthBound Solutions\DealManager Suite\dealmanager.ini"
    Dim lngLen As Long
    Dim strBuffer As String
    
    Const dhcMaxUserName = 255
    
        strBuffer = Space(dhcMaxUserName)
    lngLen = dhcMaxUserName
    If CBool(GetUserName(strBuffer, lngLen)) Then
        UserNameWindows = Left$(strBuffer, lngLen - 1)
    Else
        UserNameWindows = ""
    End If
End Function
Function ExtractServerName(ServerName As String)
 'Extract Server Name
 ExtractServerName = Mid(ServerName, InStr(ServerName, "Server=") + 7)
 ExtractServerName = Mid(ExtractServerName, 1, InStr(ExtractServerName, ";") - 1)
End Function
Function ExtractDBName(DBName As String)
 'Extract Database Name
 If InStr(DBName, "Catalog=") > 0 Then
  ExtractDBName = Mid(DBName, InStr(DBName, "Catalog=") + 8)
 ElseIf InStr(DBName, "catalog=") > 0 Then
  ExtractDBName = Mid(DBName, InStr(DBName, "catalog=") + 8)
 ElseIf InStr(DBName, "Database=") > 0 Then
  ExtractDBName = Mid(DBName, InStr(DBName, "Database=") + 9)
 ElseIf InStr(DBName, "database=") > 0 Then
  ExtractDBName = Mid(DBName, InStr(DBName, "database=") + 9)
 End If
 ExtractDBName = Mid(ExtractDBName, 1, InStr(ExtractDBName, ";") - 1)
End Function
Function ExtractUserName(UserName As String)
 'Extract User ID
 If InStr(UserName, "Uid=") > 0 Then
  ExtractUserName = Mid(UserName, InStr(UserName, "Uid=") + 4)
  ExtractUserName = Mid(ExtractUserName, 1, InStr(ExtractUserName, ";") - 1)
 ElseIf InStr(UserName, "User Id=") > 0 Then
  ExtractUserName = Mid(UserName, InStr(UserName, "User Id=") + 8)
  ExtractUserName = Mid(ExtractUserName, 1, InStr(ExtractUserName, ";") - 1)
 Else
  ExtractUserName = ""
 End If
End Function
Function ExtractPassword(PwdName As String)
 'Extract Password
 If InStr(PwdName, "Pwd=") > 0 Then
  ExtractPassword = Mid(PwdName, InStr(PwdName, "Pwd=") + 4)
  ExtractPassword = Mid(ExtractPassword, 1, InStr(ExtractPassword, ";") - 1)
 ElseIf InStr(PwdName, "Password=") > 0 Then
  ExtractPassword = Mid(PwdName, InStr(PwdName, "Password=") + 9)
  ExtractPassword = Mid(ExtractPassword, 1, InStr(ExtractPassword, ";") - 1)
 Else
  ExtractPassword = ""
 End If
End Function
Function ExtractFileName(ReportName As String)
 'Extract File Name
 ExtractFileName = Mid(ReportName, InStrRev(ReportName, "\") + 1)
End Function
Function GetPrivateProfileString32(ByVal strFileName As String, ByVal strSection As String, _
ByVal strKey As String, Optional strDefault) As String
 'Initialize Variables
 Dim strReturnString As String
 Dim lngSize As Long
 Dim lngValid As Long
 
 'Extract Data From .ini File
 On Error Resume Next
 If IsMissing(strDefault) Then
  strDefault = ""
 End If
  
 strReturnString = Space(1024)
 lngSize = Len(strReturnString)
 lngValid = GetPrivateProfileStringA(strSection, strKey, strDefault, strReturnString, lngSize, strFileName)
 GetPrivateProfileString32 = Left(strReturnString, lngValid)
End Function
Function GetLastRow(TNCol As Integer, WSName As String) As Long
 If Sheets.Count > 1 Then
  'Determine Last Row of Data
  GetLastRow = Sheets(WSName).Range("A2").SpecialCells(xlLastCell).Row
  If GetLastRow > 1 Then
   If Sheets(WSName).Cells(GetLastRow, TNCol).Value = "" Then
    GetLastRow = Sheets(WSName).Cells(GetLastRow, TNCol).End(xlUp).Row
    If GetLastRow = 1 Then
     GetLastRow = 2
    End If
   End If
  End If
 End If
End Function
Function GetTNColNum() As Long
 'Initialize Variables
 Dim LastColNum As Integer
 Dim i As Integer
 Dim j As Integer
 
 'Determine Test Name Column Number, If Any
 LastColNum = 234
 GetTNColNum = 0
 For j = 1 To 10
  For i = 1 To LastColNum
   If Sheets("Report").Cells(j, i).Value = "Test Name" Then
    'Set Test Name Column Number
    GetTNColNum = i
   
    If j > 1 Then
     'Reset Column Headers For Deal Test Data
     Sheets("Report").Cells(1, i).Value = "Test Name"
     Sheets("Report").Cells(1, i + 1).Value = "Test Result"
     Sheets("Report").Cells(1, i + 2).Value = "Difference"
     Sheets("Report").Cells(1, i + 3).Value = "Test Type"
     Sheets("Report").Range(Cells(1, i).Address & ":" & _
     Cells(1, i + 3).Address).HorizontalAlignment = xlCenter
    End If
   
    'Exit Loop
    Exit For
   ElseIf Sheets("Report").Cells(j, i).Value = "TestName" Then
    'Set Test Name Column Number
    GetTNColNum = i
   
    'Reset Column Headers For Deal Test Data
    Sheets("Report").Cells(1, i).Value = "Test Name"
    Sheets("Report").Cells(1, i + 1).Value = "Test Result"
    Sheets("Report").Cells(1, i + 2).Value = "Difference"
    Sheets("Report").Cells(1, i + 3).Value = "Test Type"
    Sheets("Report").Range(Cells(1, i).Address & ":" & _
    Cells(1, i + 3).Address).HorizontalAlignment = xlCenter
   
    'Exit Loop
    Exit For
   ElseIf i = LastColNum And j = 10 Then
    'Set Test Name Column Number To Last Column Number
    GetTNColNum = i
   End If
  Next i
  
  If GetTNColNum = i Then
   'Exit Loop
   Exit For
  End If
 Next j
End Function
Function GetCIFirstRow() As Integer
 'Initialize Variable
 Dim WSName As String
 Dim i As Integer
 Dim j As Integer
 
 'Set KDI-CI Worksheet Name
 WSName = "KDI-CI"
 
 'Determine Last Row of KDI-CI Worksheet Data
 j = GetLastRow(1, WSName)
 
 If j > 1 Then
  'Determine First Row of CI Data
  For i = 2 To j
   If Sheets(WSName).Range("B" & i).Value = "Calculated" Then
    'Set CI First Row Number
    GetCIFirstRow = i
   
    'Exit Loop
    Exit For
   ElseIf i = j Then
    'Set CI First Row Number
    GetCIFirstRow = j + 1
   End If
  Next i
 Else
  'Set CI First Row Number
  GetCIFirstRow = j + 1
 End If
End Function
Function GetKDICIWSName() As String
 'Initialize Variable
 Dim sh As Worksheet
 
 'Loop Through Worksheets To Find KDI-IC Worksheet
 For Each sh In Sheets
  If sh.Range("A1") = "ID" And sh.Range("B1") = "Source" And _
  sh.Range("C1") = "Name" And sh.Range("D1") = "Type" Then
   'Assign Variable Value
   GetKDICIWSName = sh.Name
   
   'Exit Loop
   Exit For
  End If
 Next sh
End Function
Function GetNewDealInputID(KDICISheet As String, InputRow As Integer) As Long 'This Function Creates a New Record For the New Calculated Input in the DealManager Database
  'Initialize Variables
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim connStr As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim SelClause As String
 Dim FromClause As String
 Dim WhereClause As String
 Dim SQLStmt As String
 Dim fso As Object

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Check For Data To Be Exported
 If Sheets(KDICISheet).Range("A2") = "" Then
  'Select First Cell in Report Worksheet
  Call ReportFirstCell
  
  'Exit Procedure
  Exit Function
 End If
 
 'Set DealManager .ini File Name, Server Name & Database Name
 Set fso = CreateObject("Scripting.FileSystemObject")
 DMIniFileName = "C:\Documents and Settings\" & UserNameWindows & "\Application Data\NorthBound Solutions\DealManager Suite\dealmanager.ini"
 '"C:\Users\" & UserNameWindows & "\AppData\Roaming\NorthBound Solutions\DealManager Suite\dealmanager.ini" 'Mid(ThisWorkbook.path, 1, InStrRev(ThisWorkbook.path, "\")) & "dealmanager.ini"
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
 DBTable = "DealMetrics"
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Set Basic INSERT INTO Clause
 InsertClause = "INSERT INTO " & DBTable & "("
 
 'Set Values Portion of INSERT INTO Clause
 InsertClause = InsertClause & "DealID, Name, Source, Description, Type) VALUES ("
 InsertClause = InsertClause & "'" & Range("DealID") & "', '" & Sheets(KDICISheet).Range("C" & InputRow) & "', '" & _
 "Calculated" & "', '" & Sheets(KDICISheet).Range("E" & InputRow) & "', '" & _
 Sheets(KDICISheet).Range("D" & InputRow) & "')"
  
 'SQL Query String
 SQLStmt = InsertClause & ";"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
 'Set SELECT Clause
 SelClause = "SELECT * "
 
 'Set FROM Clause Database
 FromClause = "FROM " & DBTable
   
 'Set WHERE Clause Criteria
 WhereClause = " WHERE Name = '" & Sheets(KDICISheet).Range("C" & InputRow) & "'"
 
 'SQL Query String
 SQLStmt = SelClause & FromClause & WhereClause & ";"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
 'Insert Consultant's Top Project
 GetNewDealInputID = adoRS.Fields("DealInputID")
  
 'Close Recordset & Connection
 adoRS.Close
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls

 'Skip Error Handler
 Exit Function

'Error Handler
ErrorHandler:
Call ErrorLogRecord("GetNewDealInputID", Err.Number, Err.Description)
End Function
Function NRCheck(NRName As String) As Boolean
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Check If Named Range Exists
 NRCheck = Len(ThisWorkbook.Names(NRName).Name) <> 0
 
 'Skip Error Handler
 Exit Function

'Error Handler
ErrorHandler:
Call ErrorLogRecord("NRCheck", Err.Number, "Missing Named Range: " & NRName)
End Function
Function NRCheckOpen(NRName As String) As Boolean
 'Check If Named Range Exists
 On Error Resume Next
 NRCheckOpen = Len(ThisWorkbook.Names(NRName).Name) <> 0
End Function

Function GetConnectionString() As String
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim connStr As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim SelClause As String
 Dim FromClause As String
 Dim WhereClause As String
 Dim SQLStmt As String
 Dim fso As Object


' locate and parse the connection string so that related procedures
' can connect to the DealManager DB correctly.

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

GetConnectionString = connStr
 'MsgBox connStr

End Function


Function CheckforDealMetricValue(DealID As Integer, EffectiveDt As Date, DMid As Integer) As Variant

' when updating the database with values, first check to see if a value already exists.  if so, check to see
' if the value is null.  If null, then update the same record with values.  If not null, then check to see
' if the values match, then update the current date.  if not, then insert the existing row into
' DealMetricValuesHistory table and add a new row

' This function returns a value, if one exists for the desired data/deal/id combo.
'   1. Insert means that a record does not exist for this record and should be inserted
'   2. Update All - means that a recrod exists, but all are null
'   3. Update Dt - means that a record exists and value is the same.

Dim sValue As String
Dim adoNewConn As ADODB.Connection
Dim adoRS As ADODB.Recordset
Dim connStr As String
Dim DBTable As String
 Dim InsertClause As String
 Dim SelClause As String
 Dim FromClause As String
 Dim WhereClause As String
 Dim SQLStmt As String
 Dim fso As Object



' 1. Get the connection string

'Call GetConnectionString

connStr = GetConnectionString
' 2. Query the db to get the current value
 DBTable = "DealMetricValues"
 
 'Create & Open New Connection
 Set adoNewConn = New ADODB.Connection
 adoNewConn.Open connStr
'Set SELECT Clause
 SelClause = "SELECT * "
 
 
 'Set FROM Clause Database
 FromClause = "FROM " & DBTable
   
 'Set WHERE Clause Criteria
 WhereClause = " WHERE DealID = '" & DealID & "' and DealMetricID = '"
 WhereClause = WhereClause & DMid & "' and EffectiveDt = '" & EffectiveDt
 
 'SQL Query String
 SQLStmt = SelClause & FromClause & WhereClause & "';"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 'Debug.Print SQLStmt
 'MsgBox SQLStmt
 On Error Resume Next
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoNewConn
 
 CheckforDealMetricValue = adoRS.Fields("Value")
 vDMValueID = adoRS.Fields("DealMetricValueID")

 
 'Get the Value
 If IsEmpty(CheckforDealMetricValue) Then ' = adoRS.Fields("Value")
    ' No record exists for this date - need to insert a new record with new EffectiveDT
    CheckforDealMetricValue = "InsertNew"
 Else
    If CheckforDealMetricValue = "" Then
        CheckforDealMetricValue = "UpdateEmpty"
        
    Else
        CheckforDealMetricValue = adoRS.Fields("Value")
         ' get the values incase we need to insert into Hist table
        vValue = adoRS.Fields("Value")
        vDealID = adoRS.Fields("DealID")
        vDMid = adoRS.Fields("DealMetricID")
        vCreateDt = adoRS.Fields("CreateDt")
        vEffectiveDt = adoRS.Fields("EffectiveDt")
        vComments = adoRS.Fields("Comments")
        vSource = adoRS.Fields("Source")

    End If
 End If
 On Error GoTo 0
 
  
 'Close Recordset & Connection
 adoRS.Close
 adoNewConn.Close
' 3. If nothing exists, then insert the information



End Function

Sub TestCheckForDMV()
Dim dEffDt As Date
Dim iDealID As Integer
Dim iDMid As Integer
Dim sValue As String
Dim sAnswer As String
Dim vSame As Variant
Dim iType As Integer

dEffDt = ThisWorkbook.Sheets("Settings").Range("MEDT").Value '"12/31/2011" '
iDealID = ThisWorkbook.Sheets("Settings").Range("Dealid").Value
iDMid = "45"
sValue = "$200.00"
iType = "1"


sAnswer = CheckforDealMetricValue(iDealID, dEffDt, iDMid)

' remove the $ signs, if any

sAnswer = Replace(sAnswer, "$", "", 1)
sValue = Replace(sValue, "$", "", 1)

vSame = StrComp(sAnswer, sValue, vbBinaryCompare)
If vSame = "0" Then
    MsgBox "the values are equal, only update the CreateDt"
    GoTo NextStop
End If

Select Case sAnswer
    Case "InsertNew"
     MsgBox "no effective date - insert new"
     
    Case "UpdateEmpty"
        MsgBox "the value is null, update all"
           
     Case Else
            MsgBox "the values are not equal, insert a row in DMHist table and update record"
End Select
NextStop:

End Sub


Sub NewGetNewDealInputID()  'KDICISheet As String, InputRow As Integer) As Long 'This Function Creates a New Record For the New Calculated Input in the DealManager Database
  'Initialize Variables
 Dim adoConn As ADODB.Connection
 Dim adoRS As ADODB.Recordset
 Dim DMIniFileName As String
 Dim DMServerName As String
 Dim DMDBName As String
 Dim UserName As String
 Dim PwdName As String
 Dim connStr As String
 Dim DBTable As String
 Dim InsertClause As String
 Dim SelClause As String
 Dim FromClause As String
 Dim WhereClause As String
 Dim SQLStmt As String
 Dim fso As Object
Dim InputRow As Integer
Dim KDICISheet As String
' Created 20140301 to directly add new DealMetric ids to reflect the newrequirements for Express on Servicer Report
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Check For Data To Be Exported
' If Sheets(KDICISheet).Range("A2") = "" Then
  'Select First Cell in Report Worksheet
 ' Call ReportFirstCell
  KDICISheet = "KDI-CI"
  
 InputRow = InputBox("Enter the Row Number that contains information associated with the DealMetricID that you want to create", "Create New Deal ID")
 MsgBox InputRow
 
 
 
 
 
 
  'Exit Procedure
 ' Exit Sub
 'End If
 
 'Set DealManager .ini File Name, Server Name & Database Name
 Set fso = CreateObject("Scripting.FileSystemObject")
 DMIniFileName = "C:\Documents and Settings\" & UserNameWindows & "\Application Data\NorthBound Solutions\DealManager Suite\dealmanager.ini"
 '"C:\Users\" & UserNameWindows & "\AppData\Roaming\NorthBound Solutions\DealManager Suite\dealmanager.ini" 'Mid(ThisWorkbook.path, 1, InStrRev(ThisWorkbook.path, "\")) & "dealmanager.ini"
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
 DBTable = "DealMetrics"
 
 'Create & Open New Connection
 Set adoConn = New ADODB.Connection
 adoConn.Open connStr
   
 'Set Basic INSERT INTO Clause
 InsertClause = "INSERT INTO " & DBTable & "("
 
 'Set Values Portion of INSERT INTO Clause
 InsertClause = InsertClause & "DealID, Name, Category, Description, Type) VALUES ("
 InsertClause = InsertClause & "'" & Range("DealID") & "', '" & Sheets(KDICISheet).Range("C" & InputRow) & "', '" & _
 "Calculated" & "', '" & Sheets(KDICISheet).Range("E" & InputRow) & "', '" & _
 Sheets(KDICISheet).Range("D" & InputRow) & "')"
  
 'SQL Query String
 SQLStmt = InsertClause & ";"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
 'Set SELECT Clause
 SelClause = "SELECT * "
 
 'Set FROM Clause Database
 FromClause = "FROM " & DBTable
   
 'Set WHERE Clause Criteria
 WhereClause = " WHERE Name = '" & Sheets(KDICISheet).Range("C" & InputRow) & "'"
 
 'SQL Query String
 SQLStmt = SelClause & FromClause & WhereClause & ";"
 
 'Create & Open New Recordset
 Set adoRS = New ADODB.Recordset
 adoRS.Open Source:=SQLStmt, ActiveConnection:=adoConn
 
 'Insert Consultant's Top Project
 'NewGetNewDealInputID = adoRS.Fields("DealInputID")
  MsgBox adoRS.Fields("DealMetricID")
 'Close Recordset & Connection
 adoRS.Close
 adoConn.Close
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls

 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("GetNewDealInputID", Err.Number, Err.Description)
End Sub
