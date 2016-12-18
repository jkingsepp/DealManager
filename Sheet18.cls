VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
 intresponse = MsgBox("Are Hedge Reserve Account Inputs Completed?", vbYesNo)
 intresponse = MsgBox("Print Monthly Report", vbYesNo)
 If intresponse = vbYes Then
    Sheets("Final Report").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    Sheets("Final Report").Select
    
Else
intresponse = MsgBox("Ok")
Exit Sub
End If
End Sub

Private Sub CommandButton2_Click()
 intresponse = MsgBox("Please change Date in Cell B7", vbYesNo)
 intresponse = MsgBox("Clear last months Hedge Reserve Account Inputs?", vbYesNo)
  If intresponse = vbYes Then
      Sheets("Hedge Reserve Account").Select
        Range("D42").Select
    Selection.Copy
    Range("D40").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D20:D20").Select
    Selection.ClearContents
    Range("B33:B35").Select
    Selection.ClearContents
    Range("D34:D36").Select
    Selection.ClearContents
    Range("A36:A36").Select
    Selection.ClearContents
    Range("D41:D41").Select
    Selection.ClearContents
       Sheets("Amortization Schedule").Select
       Range("B4:B100").Select
    Selection.ClearContents
         Sheets("Final Report").Select
             Else
intresponse = MsgBox("Ok")
Exit Sub
End If

End Sub
