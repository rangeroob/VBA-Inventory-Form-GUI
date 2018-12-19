Private Sub CLOSEBTN_Click()
Unload Me

End Sub

Private Sub Quit_Click()
Application.QUIT
End Sub


Private Sub SAVEBTN_Click()
Dim iRow As Long
Dim ws As Worksheet
Set ws = Worksheets("Inventory List")

'find first empty row in database
iRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, _
    SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1

'check for a tool type
If Trim(Me.TextBoxTOOLTYPE.Value) = "" Then
  Me.TextBoxTOOLTYPE.SetFocus
  MsgBox "Please enter a tool type" + vbNewLine + "If unknown please write MISC"
  Exit Sub
'check for a software version
ElseIf Trim(Me.TextBoxSWVERSION.Value) = "" Then
  Me.TextBoxSWVERSION.SetFocus
  MsgBox "Please enter software version if unknown please write N/A"
  Exit Sub
'check for a part number
ElseIf Trim(Me.TextBoxPARTNUMBER.Value) = "" Then
  Me.TextBoxPARTNUMBER.SetFocus
  MsgBox "Please enter a part number"
  Exit Sub
'check for a media type
ElseIf Trim(Me.TextBoxTAPECD.Value) = "" Then
  Me.TextBoxSWVERSION.SetFocus
  MsgBox "Please enter what media the software is installed on"
  Exit Sub
'check for a software type
ElseIf Trim(Me.TextBoxRELEASEIMAGE.Value) = "" Then
Me.TextBoxRELEASEIMAGE.SetFocus
MsgBox "Please enter what type of software it is" + vbNewLine + "If it is not a RELEASE, IMAGE or PATCH Put MISC ie(Firmware, etc)"
Exit Sub

End If

'copy the data to the database
'use protect and unprotect lines,
'     with your password
'     if worksheet is protected
With ws
  '.Unprotect Password:="password"
  .Cells(iRow, 1).Value = Me.TextBoxTOOLTYPE.Value
  .Cells(iRow, 2).Value = Me.TextBoxSWVERSION.Value
  .Cells(iRow, 3).Value = Me.TextBoxDESCRIPTION.Value
  .Cells(iRow, 4).Value = Me.TextBoxPARTNUMBER.Value
  .Cells(iRow, 5).Value = Me.TextBoxTAPECD.Value
  .Cells(iRow, 6).Value = Me.TextBoxRELEASEIMAGE.Value
  .Cells(iRow, 7).Value = Me.SOFTWARELOCATION.Value
  '.Protect Password:="password"
End With

'clear the data
Me.TextBoxTOOLTYPE.Value = ""
Me.TextBoxSWVERSION.Value = ""
Me.TextBoxDESCRIPTION.Value = ""
Me.TextBoxPARTNUMBER.Value = ""
Me.TextBoxTAPECD.Value = ""
Me.TextBoxRELEASEIMAGE.Value = ""
Me.TextBoxTOOLTYPE.SetFocus
ThisWorkbook.Save
End Sub

Private Sub UserForm_Initialize()
 With SOFTWARELOCATION
    .AddItem "LOCATION 1"
    .AddItem "LOCATION 2"
    .AddItem "LOCATION 3"
    .AddItem "LOCATION 4"
    End With
lbl_Exit:
    Exit Sub
End Sub
