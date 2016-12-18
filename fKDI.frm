VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fKDI 
   Caption         =   "DealManager by NBS"
   ClientHeight    =   3160
   ClientLeft      =   40
   ClientTop       =   -120
   ClientWidth     =   10980
   OleObjectBlob   =   "fKDI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fKDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Initialize Module Level Variables
Dim KDICISheet As String
Dim i As Integer
Dim j As Integer
Private Sub ListBoxIndicators_Click()
 'Set Control Tip Value
 i = Me.ListBoxIndicators.ListIndex
 KDICISheet = Me.TextBoxKDICISheet.Value
 Me.ListBoxIndicators.ControlTipText = Sheets(KDICISheet).Range("E" & i + 2)
 Application.StatusBar = Sheets(KDICISheet).Range("E" & i + 2)
End Sub
Private Sub ButtonAccept_Click()
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Set Screen Display Controls
 Call SetScreenControls

 'Select Report Worksheet
 Sheets("Report").Select
 
 'Set KDI Worksheet Name
 KDICISheet = Me.TextBoxKDICISheet.Value
 
 'Determine Last Row of KDI Worksheet Data
 j = GetCIFirstRow - 2
 
 'Save UserForm Data In KDI Worksheet
 For i = 2 To j
  If Right(Me.ListBoxIndicators.List(i - 2, 2), 1) = "?" Then
   Sheets(KDICISheet).Range("F" & i).Value = ""
  Else
   Sheets(KDICISheet).Range("F" & i).Formula = "=Report!$" & _
   Left(Me.ListBoxIndicators.List(i - 2, 2), 1) & "$" & Mid(Me.ListBoxIndicators.List(i - 2, 2), 2)
  End If
  Sheets(KDICISheet).Range("G" & i).Formula = Me.ListBoxIndicators.List(i - 2, 2)
 Next i
 
 'Reformat KDI-CI Table, If Necessary
 Call WSFormat(KDICISheet, 9, 2)
  
 'Select KDICell
 Application.EnableEvents = False
 Sheets("Report").Select
 Sheets("Report").Range(Me.TextBoxKDICell.Value).Select
 Application.EnableEvents = True
 
 'Hide This UserForm
 Me.Hide
  
 'KDI Data Successfully Added
 MsgBox "The Key Deal Indicators Data Has Been" & vbCr & _
 "Successfully Added to the KDI-CI Worksheet.", vbExclamation, gsAPP_NAME
 
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
'Hide This UserForm
Me.Hide
 
'Record Error In Error Log
Call ErrorLogRecord("ButtonAccept-fKDI", Err.Number, Err.Description)
End Sub
Private Sub ButtonCancel_Click()
 'Select KDICell
 Application.EnableEvents = False
 Sheets("Report").Select
 Sheets("Report").Range(Me.TextBoxKDICell.Value).Select
 Application.EnableEvents = True
 
 'Hide This UserForm
 Me.Hide
  
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls
End Sub
