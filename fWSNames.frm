VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fWSNames 
   Caption         =   "DealManager by NBS"
   ClientHeight    =   1080
   ClientLeft      =   40
   ClientTop       =   -120
   ClientWidth     =   6040
   OleObjectBlob   =   "fWSNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fWSNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ComboBoxWSName_Change()
 'Initialize Variable
 Dim MsgBoxQues As String
 
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Hide This UserForm
 Application.ScreenUpdating = True
 Unload Me
 Range("A1").Select
 Application.ScreenUpdating = False
  
 'Confirm Report Deletion
 Beep
 MsgBoxQues = MsgBox("This copy procedure will completely wipe out any existing" & vbCr & _
 "data in this file's 'Report' worksheet.  Continue?", vbYesNo + vbExclamation, gsAPP_NAME)
  
 'Exit Sub On "No"
 If MsgBoxQues = vbNo Then
  Exit Sub
 End If
 
 'Delete Report Worksheet
 Application.DisplayAlerts = False
 Sheets("Report").Delete
 Application.DisplayAlerts = True
 
 'Open File
 Workbooks.Open Me.TextBoxWSFileName.Value, UpdateLinks:=0
 
 'Disable Alerts
 Application.DisplayAlerts = False
 
 'Insert Selected Worksheet
 Sheets("Report").Copy Before:=Workbooks(ThisWorkbook.Name).Sheets(1)
 
 'Enable Alerts
 Application.DisplayAlerts = True
 
 'Select Selected Worksheet
 ActiveSheet.Name = "Report"
 
 'Reactivate New Workbook & Close Without Saving
 Windows(ExtractFileName(Me.TextBoxWSFileName.Value)).Activate
 ActiveWorkbook.Close SaveChanges:=False
 
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
  
 'Clear Status Message & Turn On Screen Updating
 Call ClearScreenControls

 'Completion Message
 MsgBox "The Selected File's Worksheet Data Has Been" & vbCr & _
 "Copied Into This File's 'Report' Worksheet.", vbExclamation, gsAPP_NAME
 
 'Select Report Worksheet
 Sheets("Report").Select
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("ComboBoxWSName_Change", Err.Number, Err.Description)
End Sub
