VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' clsContext.cls - Contains (Context Sensitive) Globals

' Note: Because these properties may be *very* frequently accessed, no logging is performed.
'       They're highly unlikely to go wrong anyway...

Option Explicit

Dim szServer As String
Dim lPort As Long
Dim szUsername As String
Dim szPassword As String
Dim lLogLevel As Long
Dim bLogView As Boolean
Dim bIncludeSys As Boolean
Dim szAutoHighlight As String
Dim objCurrentObject As Object
Dim szCurrentDB As String

'The connected server name
Public Property Let Server(ByVal szData As String)
  szServer = szData
End Property
Public Property Get Server() As String
  Server = szServer
End Property

'The port number connected to.
Public Property Let Port(ByVal lData As Long)
  lPort = lData
End Property
Public Property Get Port() As Long
  Port = lPort
End Property

'The current username
Public Property Let Username(ByVal szData As String)
  szUsername = szData
End Property
Public Property Get Username() As String
  Username = szUsername
End Property

'The current password
Public Property Let Password(ByVal szData As String)
  szPassword = szData
End Property
Public Property Get Password() As String
  Password = szPassword
End Property

'The current logging level
Public Property Let LogLevel(ByVal lData As Long)
  lLogLevel = lData
End Property
Public Property Get LogLevel() As Long
  LogLevel = lLogLevel
End Property

'Are we displaying the Log Window?
Public Property Let LogView(ByVal bData As Boolean)
  bLogView = bData
End Property
Public Property Get LogView() As Boolean
  LogView = bLogView
End Property

'Are we displaying the System Objects?
Public Property Let IncludeSys(ByVal bData As Boolean)
  bIncludeSys = bData
End Property
Public Property Get IncludeSys() As Boolean
  IncludeSys = bIncludeSys
End Property

'The AutoHighlight String
Public Property Let AutoHighlight(ByVal szData As String)
  szAutoHighlight = szData
End Property
Public Property Get AutoHighlight() As String
  AutoHighlight = szAutoHighlight
End Property

'The Current Object
Public Property Set CurrentObject(ByVal vData As Variant)
On Error Resume Next

  Set objCurrentObject = vData
  
  frmMain.sb.Panels("currentobject").Text = "Object: " & vData.Identifier & " (" & vData.ObjectType & ")"
  
  'You can't drop a Server object.
  If vData.ObjectType = "Server" Then
    frmMain.tb.Buttons("drop").Enabled = False
    frmMain.mnuPopupDrop.Enabled = False
  Else
    frmMain.tb.Buttons("drop").Enabled = True
    frmMain.mnuPopupDrop.Enabled = True
  End If
  
  'Show/Hide the View Data Options
  If vData.ObjectType = "Table" Or vData.ObjectType = "View" Then
    frmMain.tb.Buttons("viewdata").Enabled = True
    frmMain.mnuPopupViewData.Enabled = True
  Else
    frmMain.tb.Buttons("viewdata").Enabled = False
    frmMain.mnuPopupViewData.Enabled = False
  End If
  
  'Revision Control options
  If vData.ObjectType = "Aggregate" Or vData.ObjectType = "Function" Or _
     vData.ObjectType = "Index" Or vData.ObjectType = "Language" Or _
     vData.ObjectType = "Operator" Or vData.ObjectType = "Rule" Or _
     vData.ObjectType = "Sequence" Or vData.ObjectType = "Table" Or _
     vData.ObjectType = "Trigger" Or vData.ObjectType = "Type" Or _
     vData.ObjectType = "View" Then
    If frmMain.svr.Databases(ctx.CurrentDB).RevisionControl And vData.RCStatus <> rcNotApplicable Then
      frmMain.mnuPopupCommit.Enabled = True
      frmMain.tb.Buttons("commit").Enabled = True
      frmMain.mnuPopupHistory.Enabled = True
      frmMain.tb.Buttons("history").Enabled = True
    Else
      frmMain.mnuPopupCommit.Enabled = False
      frmMain.tb.Buttons("commit").Enabled = False
      frmMain.mnuPopupHistory.Enabled = False
      frmMain.tb.Buttons("history").Enabled = False
    End If
  Else
    frmMain.mnuPopupCommit.Enabled = False
    frmMain.tb.Buttons("commit").Enabled = False
    frmMain.mnuPopupHistory.Enabled = False
    frmMain.tb.Buttons("history").Enabled = False
  End If
  
End Property
Public Property Get CurrentObject() As Object
  Set CurrentObject = objCurrentObject
End Property

'The connected server name
Public Property Let CurrentDB(ByVal szData As String)
  szCurrentDB = szData
  If szData <> "" Then
    'Graveyard
    If frmMain.svr.Databases(szData).RevisionControl Then
      frmMain.mnuPopupGraveyard.Enabled = True
      frmMain.tb.Buttons("graveyard").Enabled = True
    Else
      frmMain.mnuPopupGraveyard.Enabled = False
      frmMain.tb.Buttons("graveyard").Enabled = False
    End If
    frmMain.sb.Panels("currentdb").Text = "Database: " & szData
    frmMain.tb.Buttons("create").ButtonMenus("aggregate").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("function").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("index").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("language").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("operator").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("rule").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("sequence").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("table").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("trigger").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("type").Enabled = True
    frmMain.tb.Buttons("create").ButtonMenus("view").Enabled = True
    frmMain.tb.Buttons("sql").Enabled = True
    frmMain.tb.Buttons("vacuum").Enabled = True
    frmMain.mnuFileSaveDBSchema.Enabled = True
    frmMain.mnuPopupCreateAggregate.Enabled = True
    frmMain.mnuPopupCreateFunction.Enabled = True
    frmMain.mnuPopupCreateIndex.Enabled = True
    frmMain.mnuPopupCreateLanguage.Enabled = True
    frmMain.mnuPopupCreateOperator.Enabled = True
    frmMain.mnuPopupCreateRule.Enabled = True
    frmMain.mnuPopupCreateSequence.Enabled = True
    frmMain.mnuPopupCreateTable.Enabled = True
    frmMain.mnuPopupCreateTrigger.Enabled = True
    frmMain.mnuPopupCreateType.Enabled = True
    frmMain.mnuPopupCreateView.Enabled = True
    frmMain.mnuPopupSQL.Enabled = True
    frmMain.mnuPopupVacuum.Enabled = True
  Else
    frmMain.sb.Panels("currentdb").Text = "Database: None"
    frmMain.tb.Buttons("create").ButtonMenus("aggregate").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("function").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("index").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("language").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("operator").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("rule").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("sequence").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("table").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("trigger").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("type").Enabled = False
    frmMain.tb.Buttons("create").ButtonMenus("view").Enabled = False
    frmMain.tb.Buttons("graveyard").Enabled = False
    frmMain.tb.Buttons("sql").Enabled = False
    frmMain.tb.Buttons("vacuum").Enabled = False
    frmMain.mnuFileSaveDBSchema.Enabled = False
    frmMain.mnuPopupCreateAggregate.Enabled = False
    frmMain.mnuPopupCreateFunction.Enabled = False
    frmMain.mnuPopupCreateIndex.Enabled = False
    frmMain.mnuPopupCreateLanguage.Enabled = False
    frmMain.mnuPopupCreateOperator.Enabled = False
    frmMain.mnuPopupCreateRule.Enabled = False
    frmMain.mnuPopupCreateSequence.Enabled = False
    frmMain.mnuPopupCreateTable.Enabled = False
    frmMain.mnuPopupCreateTrigger.Enabled = False
    frmMain.mnuPopupCreateType.Enabled = False
    frmMain.mnuPopupCreateView.Enabled = False
    frmMain.mnuPopupGraveyard.Enabled = False
    frmMain.mnuPopupSQL.Enabled = False
    frmMain.mnuPopupVacuum.Enabled = False
  End If
End Property
Public Property Get CurrentDB() As String
  CurrentDB = szCurrentDB
End Property
