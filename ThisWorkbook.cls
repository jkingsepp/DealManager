VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()
 'Disable Alerts & Update Links Check
 Application.DisplayAlerts = False
 Call DisableLinks
 
 'Set Screen Display Controls
 Call SetScreenControls
 
 'Show DataInsert Popup & Set Timer To Run Popup Close Procedure In 5 Seconds
 'Triggered Only If No Data Detected Below Row 1 of Report Worksheet
 If ThisWorkbook.Sheets("Report").Range("A1").SpecialCells(xlLastCell).Row = 1 Then
  'Set Top & Left Position & Show UserForm
  fDataInsert.Top = Application.Height / 4
  fDataInsert.Left = (Application.Width - fDataInsert.Width) / 2
  fDataInsert.Show
  Application.OnTime DateAdd("s", 5, Now), "CloseDataInsertUserForm"
 End If
 
 'Disable Events
 Application.EnableEvents = False
 
 'Check For Basic Worksheets
 Call WorksheetsSetup
 
 'Select First Cell in Report Worksheet
 Call ReportFirstCell
  
 'Unlock Report Worksheet
 Call UnlockReport
 
 'Lock Report Worksheet
 Call LockReport
 
 'Enable Events
 Application.EnableEvents = True
 
 'Show Macro Operation
 Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Activate()
 'Set Up DealManager Menu Item & List
 Call CreateMenu
 
 'Set Up DealManager Floating Toolbar
 Call CreateToolbar
 
 'Enable Alerts
 Application.DisplayAlerts = True
End Sub
Private Sub Workbook_Deactivate()
 'Delete DealManager Menu Item
 Call DeleteMenu
 
 'Delete DealManager Floating Toolbar
 Call DeleteToolbar
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
 'Reset Toolbar Sheet Selection Value
 Call ResetToolbarSheetSelection(sh.Name)
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Excel.Range)
 If sh.Name = "Report" And Target.Row > 1 And _
 (fTestData.Visible = True Or fKDI.Visible = True Or fCalcInput.Visible = True) Then
  'Insert Worksheet Selections
  Call SelectionInput(Target)
 End If
End Sub
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
 'Save Toolbar Positon Settings
 Call SaveToolbarPosition
End Sub
