Attribute VB_Name = "basPatch"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basPatch.bas - Contains functions to patch program

Option Explicit

'Patch form
Public Sub PatchForm(objForm As Form)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchForm(" & objForm.Name & ")", etFullDebug

  PatchFormFont objForm

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchForm"
End Sub

'Patch font of component form
Private Sub PatchFormFont(objForm As Form)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchFormFont(" & objForm.Name & ")", etFullDebug

Dim objCtrl As Control

  For Each objCtrl In objForm.Controls
    If TypeOf objCtrl Is ComboBox Or _
       TypeOf objCtrl Is TextBox Or _
       TypeOf objCtrl Is ListBox Or _
       TypeOf objCtrl Is CheckBox Or _
       TypeOf objCtrl Is ListView Or _
       TypeOf objCtrl Is TreeView Or _
       TypeOf objCtrl Is HBX Or _
       TypeOf objCtrl Is ImageCombo Or _
       TypeOf objCtrl Is OptionButton Then
      objCtrl.Font = ctx.Font
    End If
  Next


  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchFormFont"
End Sub

