VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCalcInput 
   Caption         =   "DealManager by NBS"
   ClientHeight    =   2600
   ClientLeft      =   40
   ClientTop       =   -120
   ClientWidth     =   8240.001
   OleObjectBlob   =   "fCalcInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fCalcInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private Sub ButtonAccept_Click()
 'Initialize Variables
 Dim KDICISheet As String
 Dim CIRange As String
 Dim CICell As String
 Dim i As Integer
 Dim j As Integer
 Dim k As Integer
 Dim l As Integer
 Dim m As Integer
Dim r As Worksheet
 Dim rCol As Integer
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 Set r = ThisWorkbook.ActiveSheet
'rCol = r.Range("rciind").Column
 
 'Select Report Worksheet
' Sheets("Report").Select
 
     If fCalcInput.TextBoxCIValue.Value = "" Then
  'No Value Selected Message
  MsgBox "The selected cell does not contain any data.", vbExclamation, gsAPP_NAME
 Else
  'Find First & Last Calculated Input Rows In KDI-CI Table
  KDICISheet = GetKDICIWSName
  j = GetCIFirstRow
  k = GetLastRow(1, KDICISheet)
  

  
  
  'Find Selected Cell In Report Worksheet
  CIRange = "g" & j & ":g" & k 'check the range on the KDI-CI worksheet
  
  ' add the formula
Call callfunction
  
  CICell = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
  If WorksheetFunction.CountIf(Sheets(KDICISheet).Range(CIRange), CICell) > 0 Then
   i = Application.WorksheetFunction.Match(CICell, Sheets(KDICISheet).Range(CIRange), 0)
   Sheets(KDICISheet).Range("C" & i).Value = Me.TextBoxCIName.Value
   Sheets(KDICISheet).Range("D" & i).Value = Me.ComboBoxCIType.Value
   Sheets(KDICISheet).Range("E" & i).Value = Me.TextBoxCIDescr.Value
    ' add the formula reference not the value
   Sheets(KDICISheet).Range("F" & i).Formula = "=" & r.Name & "!" & Me.TextBoxCICell.Value  'Me.TextBoxCICell.Value

   Sheets(KDICISheet).Range("G" & i).Value = Me.TextBoxCICell.Value
   
 
  Else
   Sheets(KDICISheet).Range("B" & k + 1).Value = "Calculated"
   Sheets(KDICISheet).Range("C" & k + 1).Value = Me.TextBoxCIName.Value
   Sheets(KDICISheet).Range("D" & k + 1).Value = Me.ComboBoxCIType.Value
   Sheets(KDICISheet).Range("E" & k + 1).Value = Me.TextBoxCIDescr.Value
   ' add the formula reference not the value
  Sheets(KDICISheet).Range("F" & k + 1).Formula = "=" & r.Name & "!" & Me.TextBoxCICell.Value ' ActiveCell.Offset(0, 0).Address(False, False)
'Address 'Me.TextBoxCICell.Value

   Sheets(KDICISheet).Range("G" & k + 1).Value = Me.TextBoxCICell.Value
   Sheets(KDICISheet).Range("A" & k + 1).Value = GetNewDealInputID(KDICISheet, k + 1)
   Sheets(KDICISheet).Range("H" & k + 1).Formula = _
   "=VLOOKUP(" & "A" & k + 1 & ",$A$" & j & ":$G$" & k + 1 & ",6,FALSE)"
   Sheets(KDICISheet).Range("I" & k + 1).Formula = _
   "=VLOOKUP(" & "A" & k + 1 & ",$A$" & j & ":$G$" & k + 1 & ",7,FALSE)"

  End If
 
  'Reformat KDI-CI Table, If Necessary
  Call WSFormat(KDICISheet, 9, 2)
  
  'Show Completion Message Label
  Me.LabelComplete.Visible = True
 End If
 
 'Select Last Selected CICell
 Sheets("Report").Select
 Sheets("Report").Range(Me.TextBoxCICell.Value).Select
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
'Hide This UserForm
Me.Hide
 
'Record Error In Error Log
Call ErrorLogRecord("ButtonAccept-fCIData", Err.Number, Err.Description)
End Sub
Private Sub ButtonClose_Click()
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Select Last Selected CICell
 Sheets("Report").Select
 Sheets("Report").Range(Me.TextBoxCICell.Value).Select
 
 'Hide This UserForm
 Me.Hide
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
End Sub





Private Sub UserForm_Activate()
Call callfunction


End Sub

