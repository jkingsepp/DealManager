Attribute VB_Name = "mAppControls"
Option Explicit
'Initialize Module Level Variables
Dim sh As Worksheet
Dim i As Integer
Sub CreateMenu()
 'Initialize Variables
 Dim MenuBarObject As CommandBar
 Dim MenuObject As CommandBarPopup
 Dim MenuItem As Object
 Dim SubMenuItem As Object

 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Delete Existing DealManager Menu, If Any
 Call DeleteMenu
    
 'Create DealManager Menu Item
 Set MenuBarObject = Application.CommandBars(1)
 i = MenuBarObject.Controls("&Help").Index
 Set MenuObject = Application.CommandBars(1).Controls.Add(msoControlPopup, , , i, True)
 MenuObject.Caption = "Deal&Manager"

 'Create Menu Button - View Settings
 Set MenuItem = MenuObject.Controls.Add(msoControlButton)
 With MenuItem
  .FaceId = 109
  .Caption = "&View Settings"
  .OnAction = "GoToSettings"
 End With
 
 'Create Menu Button - Format Worksheets
 Set MenuItem = MenuObject.Controls.Add(msoControlButton)
 With MenuItem
    .FaceId = 144
     .Caption = "&Format Custom Worksheets"
     .OnAction = "FormatWorksheets"
 End With
   
 'Create Menu Group Item - Deal Tests & 7 Buttons
 'Set MenuItem = MenuObject.Controls.Add(msoControlPopup)
 'MenuItem.Caption = "&Format Worksheets"
 'For i = 1 To 7
 ' If i <> 2 Then
 '  Set SubMenuItem = MenuItem.Controls.Add(msoControlButton)
 ' End If
 ' Select Case i
 '  Case 1 'Copy In Deal Data
 '  With SubMenuItem
 '   .FaceId = 1642
 '   .Caption = "&Copy In Deal Data"
 '   .OnAction = "CopyDealData"
 '  End With
 '  Case 2 'Format Worksheets
 '  If NRCheckOpen("AgingBal") = True Then
 '   Set SubMenuItem = MenuItem.Controls.Add(msoControlButton)
 '   With SubMenuItem
 '    .FaceId = 144
 '    .Caption = "&Format Custom Worksheets"
 '    .OnAction = "FormatWorksheets"
 '   End With
 '  End If
 '  Case 3 'Define/Modify Tests
 '  With SubMenuItem
 '   .FaceId = 2899
 '   .Caption = "&Define/Modify Deal Tests"
 '   .OnAction = "ConfigureDealTests"
 '  End With
 '  Case 4 'Check for Formula Errors
 '  With SubMenuItem
  '  .FaceId = 5687
  '  .Caption = "Check For Formula &Errors"
 '   .OnAction = "ReviewFormulas"
 '  End With
 '  Case 5 'Check for Violations
 '  With SubMenuItem
 '   .FaceId = 6122
 '   .Caption = "&Check For Deal Violations"
 '   .OnAction = "ReviewDealViolations"
 '  End With
 '  Case 6 'Update Test Worksheet
 '  With SubMenuItem
 '   .FaceId = 459
 '   .Caption = "Update &Tests Worksheet"
 '   .OnAction = "UpdateTestsWS"
 '  End With
 '  Case 7 'Upload Tests to DealManager
 '  With SubMenuItem
 '   .FaceId = 7432
 '   .Caption = "&Upload Test Data to DM"
 '   .OnAction = "ExportDataToDM"
 '  End With
 ' End Select
' Next i
'
' 'Create Menu Group Item - Deal Input Values & 3 Buttons
' Set MenuItem = MenuObject.Controls.Add(msoControlPopup)
' MenuItem.Caption = "Deal &Input Values"
' For i = 1 To 3
'  Set SubMenuItem = MenuItem.Controls.Add(msoControlButton)
'  Select Case i
'   Case 1 'Define KDIs
'   With SubMenuItem
'    .FaceId = 2899
'    .Caption = "&Define Key Deal Indicators"
'    .OnAction = "ConfigureKDIs"
'   End With
'   Case 2 'Define CIs
'   With SubMenuItem
'    .FaceId = 2899
'    .Caption = "Define &Calculated Inputs"
'    .OnAction = "ConfigureCIs"
'   End With
'   Case 3 'Upload KDI & CI Data to DealManager
'   With SubMenuItem
'    .FaceId = 7432
'    .Caption = "&Upload KDI && CI Data to DM"
'    .OnAction = "ExportDataToDM"
'   End With
'  End Select
' Next i
'
 'Create Menu Group Item - Report Center & 2 Buttons
 Set MenuItem = MenuObject.Controls.Add(msoControlPopup)
 MenuItem.Caption = "&Reports"
 For i = 1 To 2
  Set SubMenuItem = MenuItem.Controls.Add(msoControlButton)
  Select Case i
   Case 1 'Create Trial Report
   With SubMenuItem
    .FaceId = 2572
    .Caption = "Create &Trial Report"
    .OnAction = "ExportReport"
   End With
   Case 2 'Create Final Report
   With SubMenuItem
    .FaceId = 2573
    .Caption = "Create &Final Report"
    .OnAction = "ExportReport"
   End With
  End Select
 Next i
           
 'Create Menu Button - Help
 Set MenuItem = MenuObject.Controls.Add(msoControlButton)
 With MenuItem
  .FaceId = 984
  .Caption = "&Help"
  .OnAction = "HelpScreen"
 End With
 
 'Create Menu Button - About NorthBound Solutions
 Set MenuItem = MenuObject.Controls.Add(msoControlButton)
 With MenuItem
  .FaceId = 1000
  .Caption = "&About NorthBound Solutions"
  .OnAction = "AboutNBS"
 End With
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("CreateMenu", Err.Number, Err.Description)
End Sub
Sub DeleteMenu()
 'Delete Menu, If It Exists
 On Error Resume Next
 Application.CommandBars(1).Controls("Deal&Manager").Delete
 On Error GoTo 0
End Sub
Sub DeleteMenu_Old()
 'Initialize Variable
 Dim MenuObject As CommandBarPopup
 
 For Each MenuObject In CommandBars.FindControls(msoControlPopup, , , True)
  If MenuObject.Caption = "Deal&Manager" Then
   'Delete Menu
   MenuObject.Delete
   
   'Exit Loop
   Exit For
  End If
 Next MenuObject
End Sub
Sub CreateToolbar()
 'Initialize Variables

'Error Handler
'ErrorHandler:
'Call ErrorLogRecord("CreateToolbar", Err.Number, Err.Description)
End Sub
Sub DeleteToolbar()
 'Delete Toolbar, If It Exists
 On Error Resume Next
 Application.CommandBars("DealManager").Delete
 On Error GoTo 0
End Sub
Sub DeleteToolbar_Old()
 'Initialize Variable
 Dim MenuObject As CommandBar

 For Each MenuObject In CommandBars
  If MenuObject.Name = "DealManager" Then
   'Save Toolbar Positon Settings
   Call SaveToolbarPosition
   
   'Delete Toolbar
   MenuObject.Delete
   
   'Exit Loop
   Exit For
  End If
 Next MenuObject
End Sub
Sub SaveToolbarPosition()
 'Initialize Variable
' Dim NRName As String
    
 'Turn Error Handler On
 On Error GoTo ErrorHandler
 
 'Check To Make Sure Toolbar Exists
' Call CheckForToolbar
 
 'Save Floating Toolbar Top Position
' NRName = "TBTop"
' If NRCheck(NRName) = True Then
'  Range(NRName).Value = CommandBars("DealManager").Top
' End If
 
 'Save Floating Toolbar Left Position
' NRName = "TBLeft"
' If NRCheck(NRName) = True Then
'  Range(NRName).Value = CommandBars("DealManager").Left
' End If
 
 'Skip Error Handler
 Exit Sub

'Error Handler
ErrorHandler:
Call ErrorLogRecord("SaveToolbarPosition", Err.Number, Err.Description)
End Sub
Sub ResetToolbarSheetSelection(SheetSelection)
 'Check To Make Sure Toolbar Exists
' Call CheckForToolbar
 
 'Select Current Worksheet As Item Showing In Blank
' For i = 1 To CommandBars("DealManager").Controls("Go To Sheet").ListCount
'  If CommandBars("DealManager").Controls("Go To Sheet").List(i) = SheetSelection Then
'   CommandBars("DealManager").Controls("Go To Sheet").ListIndex = i
  
   'Exit Loop
'   Exit For
 ' End If
 'Next i
End Sub
Sub CheckForToolbar()
 'Initialize Variable
 Dim MenuObject As CommandBar

 'Check To Make Sure Toolbar Exists
' For Each MenuObject In CommandBars
'  If MenuObject.Name = "DealManager" Then
   'Exit Loop
'   Exit For
'  ElseIf MenuObject.Name = CommandBars(CommandBars.Count).Name Then
   'Set Up DealManager Menu Item & List
 '  Call CreateMenu
'
   'Set Up DealManager Floating Toolbar
 '  Call CreateToolbar
 ' End If
 'Next MenuObject
End Sub
Sub ResetToolbarPosition()
 'Check To Make Sure Toolbar Exists
 Call CheckForToolbar
 
 'Reset Toolbar Position
 CommandBars("DealManager").Top = 172
 CommandBars("DealManager").Left = GetSystemMetrics32(0) - CommandBars("DealManager").Width - 24
End Sub
Sub SetScreenControls()
 'Hide Macro Operation
 If Application.ScreenUpdating = True Then
  Application.ScreenUpdating = False
 End If

 'Show Status Bar
 If Application.DisplayStatusBar = False Then
  Application.DisplayStatusBar = True
 End If
End Sub
Sub ClearScreenControls()
 'Clear Status Bar
 Application.StatusBar = ""
 
 'Show Screen Updating If Off
 If Application.ScreenUpdating = False Then
  Application.ScreenUpdating = True
 End If
End Sub
Sub UnlockReport()
 'Unlock Report Worksheet For Adjustment
 ActiveSheet.Unprotect Password:=""
End Sub
Sub LockReport()
 'Determine Settings Table Parameters
 i = Sheets("Settings").Range("A1").SpecialCells(xlLastCell).Row
 
 'Lock Report Worksheet
 If Application.WorksheetFunction.VLookup("Worksheet Lock", Sheets("Settings").Range("A2:B" & i), 2, False) = 1 Then
  'Hide Macro Operation
  If Application.ScreenUpdating = True Then
   Application.ScreenUpdating = False
  End If
  
  'Select Report Worksheet
  Sheets("Report").Select
    
  'Lock Worksheet
  ActiveSheet.Protect UserInterfaceOnly:=True, Password:=""
  
  'Show Macro Operation
  Application.ScreenUpdating = True
 End If
End Sub
Sub CloseDataInsertUserForm()
 'Unload DataInsert UserForm
 Unload fDataInsert
End Sub
