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
' Copyright (C) 2001, The pgAdmin Development Team
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
    frmMain.tb.Buttons("drop").Visible = False
    frmMain.mnuPopupDrop.Visible = False
  Else
    frmMain.tb.Buttons("drop").Visible = True
    frmMain.mnuPopupDrop.Visible = True
  End If
  
  'Show/Hide the View Data Options
  If vData.ObjectType = "Table" Or vData.ObjectType = "View" Then
    frmMain.tb.Buttons("viewdata").Visible = True
    frmMain.mnuPopupViewData.Visible = True
  Else
    frmMain.tb.Buttons("viewdata").Visible = False
    frmMain.mnuPopupViewData.Visible = False
  End If
End Property
Public Property Get CurrentObject() As Object
  Set CurrentObject = objCurrentObject
End Property

'The connected server name
Public Property Let CurrentDB(ByVal szData As String)
  szCurrentDB = szData
  If szData <> "" Then
    frmMain.sb.Panels("currentdb").Text = "Database: " & szData
    frmMain.tb.Buttons("create").ButtonMenus("aggregate").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("function").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("index").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("language").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("operator").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("rule").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("sequence").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("table").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("trigger").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("type").Visible = True
    frmMain.tb.Buttons("create").ButtonMenus("view").Visible = True
    frmMain.tb.Buttons("sep2").Visible = True
    frmMain.tb.Buttons("sql").Visible = True
    frmMain.tb.Buttons("vacuum").Visible = True
    frmMain.mnuFileSaveDBSchema.Visible = True
    frmMain.mnuFileSep2.Visible = True
    frmMain.mnuPopupCreateAggregate.Visible = True
    frmMain.mnuPopupCreateFunction.Visible = True
    frmMain.mnuPopupCreateIndex.Visible = True
    frmMain.mnuPopupCreateLanguage.Visible = True
    frmMain.mnuPopupCreateOperator.Visible = True
    frmMain.mnuPopupCreateRule.Visible = True
    frmMain.mnuPopupCreateSequence.Visible = True
    frmMain.mnuPopupCreateTable.Visible = True
    frmMain.mnuPopupCreateTrigger.Visible = True
    frmMain.mnuPopupCreateType.Visible = True
    frmMain.mnuPopupCreateView.Visible = True
    frmMain.mnuPopupSep2.Visible = True
    frmMain.mnuPopupSQL.Visible = True
    frmMain.mnuPopupVacuum.Visible = True
  Else
    frmMain.sb.Panels("currentdb").Text = "Database: None"
    frmMain.tb.Buttons("create").ButtonMenus("aggregate").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("function").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("index").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("language").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("operator").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("rule").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("sequence").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("table").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("trigger").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("type").Visible = False
    frmMain.tb.Buttons("create").ButtonMenus("view").Visible = False
    frmMain.tb.Buttons("sep2").Visible = False
    frmMain.tb.Buttons("sql").Visible = False
    frmMain.tb.Buttons("vacuum").Visible = False
    frmMain.mnuFileSaveDBSchema.Visible = False
    If frmMain.mnuFileSaveDefinition.Visible = False Then frmMain.mnuFileSep2.Visible = False
    frmMain.mnuPopupCreateAggregate.Visible = False
    frmMain.mnuPopupCreateFunction.Visible = False
    frmMain.mnuPopupCreateIndex.Visible = False
    frmMain.mnuPopupCreateLanguage.Visible = False
    frmMain.mnuPopupCreateOperator.Visible = False
    frmMain.mnuPopupCreateRule.Visible = False
    frmMain.mnuPopupCreateSequence.Visible = False
    frmMain.mnuPopupCreateTable.Visible = False
    frmMain.mnuPopupCreateTrigger.Visible = False
    frmMain.mnuPopupCreateType.Visible = False
    frmMain.mnuPopupCreateView.Visible = False
    frmMain.mnuPopupSep2.Visible = False
    frmMain.mnuPopupSQL.Visible = False
    frmMain.mnuPopupVacuum.Visible = False
  End If
End Property
Public Property Get CurrentDB() As String
  CurrentDB = szCurrentDB
End Property
