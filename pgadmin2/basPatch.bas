Attribute VB_Name = "basPatch"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' basPatch.bas - Contains functions to patch program

Option Explicit

'Patch form
Public Sub PatchForm(objForm As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchForm(" & objForm.Name & ")", etFullDebug

  PatchFormScrObjDB objForm
  PatchFormFont objForm
  PatchFormLang objForm
  PatchFormToolTip objForm

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchForm"
End Sub

'Patch font of component form
Private Sub PatchFormFont(objForm As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchFormFont(" & objForm.Name & ")", etFullDebug

Dim objCtrl As Control

  If ctx.Font Is Nothing Then Exit Sub
  Set objForm.Font = ctx.Font
  On Error Resume Next
  For Each objCtrl In objForm.Controls
    Set objCtrl.Font = ctx.Font
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchFormFont"
End Sub

'Patch ToolTip of component form
Private Sub PatchFormToolTip(objForm As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchFormToolTip(" & objForm.Name & ")", etFullDebug

Dim objToolTip As New clsToolTip
Dim objCtrl As Control
Dim szToolTip As String
Dim vData
Dim ii As Integer
Dim szTemp As String
Dim bBlank As Boolean
Const iMAXLNG_LINE_TOOLTIP As Integer = 50

  On Error Resume Next
  For Each objCtrl In objForm.Controls
    'the HBX have a problem to use this class
    If Not TypeOf objCtrl Is HBX Then
      Err.Clear
      szToolTip = objCtrl.ToolTipText
      If Err.Number = 0 Then
        If Len(szToolTip) > 0 And Len(szToolTip) > iMAXLNG_LINE_TOOLTIP Then
          'create new tooltip
          vData = Split(szToolTip)
          szToolTip = ""
          szTemp = ""
          bBlank = False
          For ii = 0 To UBound(vData)
            If bBlank Then
              'add blank
              szToolTip = szToolTip & " " & vData(ii)
              szTemp = szTemp & " " & vData(ii)
            Else
              'no blank
              szToolTip = szToolTip & vData(ii)
              szTemp = szTemp & vData(ii)
              bBlank = True
            End If
          
            'verify insert crlf
            If Len(szTemp) >= iMAXLNG_LINE_TOOLTIP Then
              szToolTip = szToolTip & vbCrLf
              szTemp = ""
              bBlank = False
            End If
          Next
      
          'add new ToolTip
          objToolTip.AssignToolTip objCtrl, szToolTip
          objCtrl.ToolTipText = ""
        End If
      End If
    End If
  Next
  On Error GoTo 0
  Set objToolTip = Nothing
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchFormToolTip"
End Sub

'Add scroll object database
Private Sub PatchFormScrObjDB(objForm As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchFormScrollObj(" & objForm.Name & ")", etFullDebug

  Select Case Mid(objForm.Name, 4)
    Case "Aggregate", "Domain", "Function", "Operator", "Sequence", "Table", "Type", "View", "User", "Group", "Cast", "Language", "Namespace", "Database", "Column", "ForeignKey", "Rule", "Trigger", "Index", "Conversion", "OperatorClass"
      
      'create object
      objForm.Controls.Add "pgAdmin2.ScrollObjDb", "ScrollObjDb"
      With objForm!ScrollObjDb
        .Left = 25
        .Top = objForm!cmdOK.Top
        .Visible = True
      End With
   
  End Select

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchFormScrollObj"
End Sub

'Patch string lang of component form
Private Sub PatchFormLang(objForm As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basPatch.PatchFormLang(" & objForm.Name & ")", etFullDebug

Dim objCtrl As Control
Dim szTemp As String
Dim ii As Integer

  On Error Resume Next
  For Each objCtrl In objForm.Controls
    'caption
    Err.Clear
    szTemp = objCtrl.Caption
    If Err.Number = 0 Then objCtrl.Caption = §§TrasLang§§(szTemp)
    
    'ToolTip
    Err.Clear
    szTemp = objCtrl.ToolTipText
    If Err.Number = 0 Then objCtrl.ToolTipText = §§TrasLang§§(szTemp)
    
    If TypeOf objCtrl Is SSTab Then
      'SSTab
      For ii = 0 To objCtrl.Tabs - 1
        objCtrl.TabCaption(ii) = §§TrasLang§§(objCtrl.TabCaption(ii))
      Next
    ElseIf TypeOf objCtrl Is ListView Then
      'ListView
      For ii = 1 To objCtrl.ColumnHeaders.Count
        objCtrl.ColumnHeaders(ii).Text = §§TrasLang§§(objCtrl.ColumnHeaders(ii).Text)
      Next
    End If
  Next

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basPatch.PatchFormLang"
End Sub


