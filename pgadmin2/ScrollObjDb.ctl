VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ScrollObjDb 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ScaleHeight     =   375
   ScaleWidth      =   240
   Begin MSComCtl2.UpDown UpDownObjDb 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
End
Attribute VB_Name = "ScrollObjDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' ScrollObjDb.ctl - Scroll object database

Option Explicit

Private Enum EScrObjDb
  EScrObjDb_Up
  EScrObjDb_Down
End Enum

Dim szDatabase As String
Dim szNamespace As String
Dim szTable As String

Private Sub UserControl_Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
If Not frmMain.svr Is Nothing Then frmMain.svr.LogEvent "Entering " & App.Title & ":ScrollObjDb.UserControl_Initialize()", etFullDebug
  
  szDatabase = ctx.CurrentDB
  szNamespace = ctx.CurrentNS
  
  If Len(Trim(szDatabase)) <= 0 Then Exit Sub
  
  'save table name if object depend of table
  Select Case ctx.CurrentObject.ObjectType
    Case "Column", "ForeignKey", "Rule", "Trigger", "Index"
      szTable = ctx.CurrentObject.Table
  End Select

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ScrollObjDb.UserControl_Initialize"
End Sub

Private Sub UpDownObjDb_DownClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ScrollObjDb.UpDownObjDb_DownClick()", etFullDebug
  
  ScrObjDb EScrObjDb_Down
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ScrollObjDb.UpDownObjDb_DownClick"
End Sub

Private Sub UpDownObjDb_UpClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ScrollObjDb.UpDownObjDb_UpClick()", etFullDebug
  
  ScrObjDb EScrObjDb_Up

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ScrollObjDb.UpDownObjDb_UpClick"
End Sub

Private Sub ScrObjDb(ETypeMove As EScrObjDb)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ScrollObjDb.ScrObjDb(" & ETypeMove & ")", etFullDebug

Dim szObjIdentifier As String
Dim szIdentifier As String
Dim objForm As Form
Dim szObjectType As String
  
  'verify if form is in modify
  If InStr(UserControl.Extender.Container.Caption, ":") <= 0 Then Exit Sub
  
  'get identfier object by caption form (????????????????)
  szObjIdentifier = Trim(Mid(UserControl.Extender.Container.Caption, InStr(UserControl.Extender.Container.Caption, ":") + 1))
  szObjectType = Mid(UserControl.Extender.Container.Name, 4)
  Select Case szObjectType
    Case "Aggreagte", "Domain", "Function", "Operator", "Sequence", "Table", "Type", "View"

      'Find next/prev object
      szIdentifier = GetIdentifier(CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace), szObjectType & "s", VbGet), szObjIdentifier, ETypeMove)
      If Len(szIdentifier) <= 0 Then Exit Sub
      
      'load new form
      Select Case szObjectType
        Case "Aggregate"
          Set objForm = New frmAggregate
        
        Case "Domain"
          Set objForm = New frmDomain
        
        Case "Function"
          Set objForm = New frmFunction
        
        Case "Operator"
          Set objForm = New frmOperator
        
        Case "Sequence"
          Set objForm = New frmSequence
        
        Case "Table"
          Set objForm = New frmTable
        
        Case "Type"
          Set objForm = New frmType
        
        Case "View"
          Set objForm = New frmView
      
      End Select
      
      'load object
      StartMsg "Load " & szObjectType
      objForm.Initialise szDatabase, szNamespace, CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace), szObjectType & "s", VbGet, szIdentifier)
      ActivateForm objForm
    
    Case "User", "Group", "Database"
      
      'Find next/prev object
      szIdentifier = GetIdentifier(CallByName(frmMain.svr, szObjectType & "s", VbGet), szObjIdentifier, ETypeMove)
      If Len(szIdentifier) <= 0 Then Exit Sub
  
      'load new form
      Select Case szObjectType
        Case "User"
          Set objForm = New frmUser
        
        Case "Group"
          Set objForm = New frmGroup
      
        Case "Database"
          Set objForm = New frmDatabase
      
      End Select
      
      'load object
      Load objForm
      StartMsg "Load " & szObjectType
      objForm.Initialise CallByName(frmMain.svr, szObjectType & "s", VbGet, szIdentifier)
      ActivateForm objForm
  
    Case "Cast", "Language", "Namespace"
        
      'Find next/prev object
      szIdentifier = GetIdentifier(CallByName(frmMain.svr.Databases(szDatabase), szObjectType & "s", VbGet), szObjIdentifier, ETypeMove)
      If Len(szIdentifier) <= 0 Then Exit Sub
  
      'load new form
      Select Case szObjectType
        Case "Cast"
          Set objForm = New frmCast
        
        Case "Language"
          Set objForm = New frmLanguage
        
        Case "Namespace"
          Set objForm = New frmNamespace
      
      End Select
      
      'load object
      Load objForm
      StartMsg "Load " & szObjectType
      objForm.Initialise szDatabase, CallByName(frmMain.svr.Databases(szDatabase), szObjectType & "s", VbGet, szIdentifier)
      ActivateForm objForm
  
    Case "Column", "ForeignKey", "Rule", "Trigger", "Index"
 
      'Find next/prev object
      Select Case szObjectType
        Case "Index"
          szIdentifier = GetIdentifier(CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(szTable), "Indexes", VbGet), szObjIdentifier, ETypeMove)
        
        Case Else
          szIdentifier = GetIdentifier(CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(szTable), szObjectType & "s", VbGet), szObjIdentifier, ETypeMove)
      
      End Select
      If Len(szIdentifier) <= 0 Then Exit Sub
      
      'load new form
      Select Case szObjectType
        Case "Column"
          Set objForm = New frmColumn
      
        Case "ForeignKey"
          Set objForm = New frmForeignKey
        
        Case "Rule"
          Set objForm = New frmRule
 
        Case "Trigger"
          Set objForm = New frmTrigger
      
        Case "Index"
          Set objForm = New frmIndex
      
      End Select
  
      'load object
      Load objForm
      StartMsg "Load " & szObjectType
      Select Case szObjectType
        Case "Column"
          objForm.Initialise szDatabase, szNamespace, "MP", CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(szTable), szObjectType & "s", VbGet, szIdentifier)
        
        Case "Index"
          objForm.Initialise szDatabase, szNamespace, CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(szTable), "Indexes", VbGet, szIdentifier)
        
        Case Else
          objForm.Initialise szDatabase, szNamespace, CallByName(frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(szTable), szObjectType & "s", VbGet, szIdentifier)
      
      End Select
      ActivateForm objForm
  
  End Select

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ScrollObjDb.ScrObjDb"
End Sub

'Get identifier object
Private Function GetIdentifier(objSearch As Object, szObjIdentifier As String, ETypeMove As EScrObjDb) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ScrollObjDb.ScrObjDb(" & QUOTE & objSearch.Count & QUOTE & "," & QUOTE & szObjIdentifier & QUOTE & "," & ETypeMove & ")", etFullDebug

Dim objTmp
Dim bFound As Boolean
Dim szOldIdentifier As String
      
  GetIdentifier = ""
  bFound = False
  For Each objTmp In objSearch
    If Not (objTmp.SystemObject And Not ctx.IncludeSys) Then
      If bFound Then
        GetIdentifier = objTmp.Identifier
        Exit For
      End If
      If objTmp.Identifier = szObjIdentifier Then
        bFound = True
        If ETypeMove = EScrObjDb_Up Then
          GetIdentifier = szOldIdentifier
          Exit For
        End If
      End If
      szOldIdentifier = objTmp.Identifier
    End If
  Next
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ScrollObjDb.GetIdentifier"
End Function

'activate form and copy position
Private Sub ActivateForm(objForm As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":ScrollObjDb.ActivateForm(" & QUOTE & objForm.Name & QUOTE & ")", etFullDebug

  'copy position
  objForm.Top = UserControl.Extender.Container.Top
  objForm.Left = UserControl.Extender.Container.Left
  objForm.tabProperties.Tab = UserControl.Extender.Container.tabProperties.Tab
    
  objForm.Show
  Unload UserControl.Extender.Container
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":ScrollObjDb.ActivateForm"
End Sub


