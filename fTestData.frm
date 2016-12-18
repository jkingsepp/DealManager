VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTestData 
   Caption         =   "DealManager by NBS"
   ClientHeight    =   5280
   ClientLeft      =   40
   ClientTop       =   -120
   ClientWidth     =   8240.001
   OleObjectBlob   =   "fTestData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fTestData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SpinRow_Change()
 'Set Test Row To Value Determined By Spinbutton
 Me.TextBoxTestRow.Value = Format(Me.SpinRow.Value, "#,##0")
 Sheets("Report").Cells(Me.TextBoxTestRow.Value, ActiveCell.Column).Select
End Sub
Private Sub TextBoxTestRow_AfterUpdate()
 'Reset Test Row Box To Value Within Limits
 If Me.TextBoxTestRow.Value > 65000 Then
  Me.TextBoxTestRow.Value = 65000
 ElseIf Me.TextBoxTestRow.Value < 2 Then
  Me.TextBoxTestRow.Value = 2
 End If
 Me.SpinRow.Value = Me.TextBoxTestRow.Value
End Sub
Private Sub TextBoxLeftOperand_AfterUpdate()
 'Set New Resulting Formula
 If Me.TextBoxLeftValue.Value = "" Then
  Me.TextBoxLeftValue.Value = Me.TextBoxLeftOperand.Value
 End If
 Call SetNewResultingFormula
End Sub
Private Sub ComboBoxOperator_AfterUpdate()
 'Set New Resulting Formula
 Call SetNewResultingFormula
End Sub
Private Sub TextBoxRightOperand_AfterUpdate()
 'Set New Resulting Formula
 If Me.TextBoxRightValue.Value = "" Then
  Me.TextBoxRightValue.Value = Me.TextBoxRightOperand.Value
 End If
 Call SetNewResultingFormula
End Sub
Private Sub OptionPass_Click()
 'Set New Resulting Formula
 Call SetNewResultingFormula
End Sub
Private Sub OptionFail_Click()
 'Set New Resulting Formula
 Call SetNewResultingFormula
End Sub
Private Sub TextBoxLeftValue_AfterUpdate()
 'Set New Resulting Formula
 Call SetNewResultingFormula
End Sub
Private Sub TextBoxRightValue_AfterUpdate()
 'Set New Resulting Formula
 Call SetNewResultingFormula
End Sub
Private Sub SetNewResultingFormula()
 'Initialize Variable
 Dim OpType As String
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Operator Type
 If Me.ComboBoxOperator.Value = "" Then
  Me.ComboBoxOperator.Value = "equals"
 End If
 Select Case Me.ComboBoxOperator.Value
  Case "equals": OpType = "="
  Case "does not equal": OpType = "<>"
  Case "is greater than": OpType = ">"
  Case "is greater than or equal to": OpType = ">="
  Case "is less than": OpType = "<"
  Case "is less than or equal to": OpType = "<="
  Case "begins with": OpType = "="
  Case "does not begin with": OpType = "<>"
  Case "ends with": OpType = "="
  Case "does not end with": OpType = "<>"
  Case "contains": OpType = "="
  Case "does not contain": OpType = "="
 End Select
 
 'Set Test Data UserForm Current Left Operand Values, If Any
 Me.TextBoxLeftOperand.Value = Replace(Me.TextBoxLeftOperand.Value, """", "")
  
 'Set Test Data UserForm Current Right Operand Values, If Any
 Me.TextBoxRightOperand.Value = Replace(Me.TextBoxRightOperand.Value, """", "")
 
 'Set Pass/Fail Indicator
 If Me.OptionPass.Value = False And Me.OptionFail.Value = False Then
  Me.OptionPass.Value = True
 End If
 
 'Set New Resulting Formula
 Me.TextBoxResultingFormula.Value = "=IF("
 Me.TextBoxResultingFormula.Value = Me.TextBoxResultingFormula.Value & """" & Me.TextBoxLeftOperand.Value & """"
 Me.TextBoxResultingFormula.Value = Me.TextBoxResultingFormula.Value & OpType
 Me.TextBoxResultingFormula.Value = Me.TextBoxResultingFormula.Value & """" & Me.TextBoxRightOperand.Value & """"
 If Me.OptionPass.Value = True Then
  Me.TextBoxResultingFormula.Value = Me.TextBoxResultingFormula.Value & "," & """Pass""" & "," & """Fail""" & ")"
 Else
  Me.TextBoxResultingFormula.Value = Me.TextBoxResultingFormula.Value & "," & """Fail""" & "," & """Pass""" & ")"
 End If

 'Set New Resulting Difference Formula
 Me.TextBoxResultingFormulaDiff.Value = "="
 Me.TextBoxResultingFormulaDiff.Value = Me.TextBoxResultingFormulaDiff.Value & Me.TextBoxLeftValue.Value
 Me.TextBoxResultingFormulaDiff.Value = Me.TextBoxResultingFormulaDiff.Value & "-"
 Me.TextBoxResultingFormulaDiff.Value = Me.TextBoxResultingFormulaDiff.Value & Me.TextBoxRightValue.Value
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("SetNewResultingFormula", Err.Number, Err.Description)
End Sub
Private Sub ButtonAccept_Click()
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Select Report Worksheet
 Sheets("Report").Select
 
 'Set New Resulting Formula
 Call SetNewResultingFormula
 
 'Determine Deal Test Column
 If GetTNColNum = 234 Then
  'Insert Deal Test Columns
  Call InsertDealTestColumns
 End If
 
 'Check For All Required Entries
 If Me.TextBoxLeftOperand.Value = "" Or Me.ComboBoxOperator.Value = "" Or Me.TextBoxRightOperand.Value = "" Then
  'Set All Test Cells To Blank
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum).Value = ""
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 1).Formula = ""
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 2).Formula = ""
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 3).Value = ""
 Else
  'Set New/Revised Test Name
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum).Value = Me.TextBoxTestName.Value
 
  'Set New/Revised Pass/Fail Formula
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 1).Formula = Me.TextBoxResultingFormula.Value
  
  'Set the TestType Name
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 3).Value = Me.ComboBoxTestType.Value

  'Set the Difference Formula
  If Me.TextBoxLeftValue.Value = "" Or Me.TextBoxRightValue.Value = "" Then
   Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 2).Formula = ""
  Else
   Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 2).Formula = Me.TextBoxResultingFormulaDiff.Value
  End If
 End If
 
 'Highlight Deal Test Row
 If Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 1).Value = "Pass" Then
  'Pass Test Result
  With Sheets("Report").Range("A" & Me.TextBoxTestRow.Value & ":" & _
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 3).Address)
   .Interior.ColorIndex = 4
   .Interior.Pattern = xlSolid
   .Font.ColorIndex = 0
  End With
 Else
  'Fail Test Result
  With Sheets("Report").Range("A" & Me.TextBoxTestRow.Value & ":" & _
  Sheets("Report").Cells(Me.TextBoxTestRow.Value, GetTNColNum + 3).Address)
   .Interior.ColorIndex = 3
   .Interior.Pattern = xlSolid
   .Font.ColorIndex = 2
  End With
 End If
 
 'Select Last Selected Cell
 Sheets("Report").Select
 Sheets("Report").Cells(Format(Me.TextBoxTestRow.Value, "###0"), ActiveCell.Column).Select
 
 'Show Completion Message Label
 Me.LabelComplete.Visible = True
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
'Hide This UserForm
Me.Hide
 
'Record Error In Error Log
Call ErrorLogRecord("ButtonAccept-fTestData", Err.Number, Err.Description)
End Sub
Private Sub ButtonClose_Click()
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Select Last Selected Cell
 Sheets("Report").Select
 Sheets("Report").Cells(Format(Me.TextBoxTestRow.Value, "###0"), ActiveCell.Column).Select
 
 'Hide This UserForm
 Me.Hide
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
End Sub
