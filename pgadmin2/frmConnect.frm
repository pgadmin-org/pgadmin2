VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Server"
   ClientHeight    =   2136
   ClientLeft      =   5016
   ClientTop       =   2664
   ClientWidth     =   3600
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2136
   ScaleWidth      =   3600
   Begin VB.TextBox txtDescription 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   945
      TabIndex        =   9
      ToolTipText     =   "Enter description connection"
      Top             =   1044
      Width           =   2625
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   945
      TabIndex        =   2
      Text            =   "localhost"
      ToolTipText     =   "Enter the hostname or IP address of the server to connect to."
      Top             =   90
      Width           =   2445
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   945
      TabIndex        =   3
      Text            =   "5432"
      ToolTipText     =   "Enter the port that the PostgreSQL server is listening on."
      Top             =   405
      Width           =   870
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   945
      TabIndex        =   4
      Text            =   "postgres"
      ToolTipText     =   "Enter your username on the specified server."
      Top             =   720
      Width           =   2625
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   945
      PasswordChar    =   "*"
      TabIndex        =   0
      ToolTipText     =   "Enter your password on the specified server."
      Top             =   1368
      Width           =   2625
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   2430
      TabIndex        =   1
      ToolTipText     =   "Connect to the Specified Server."
      Top             =   1728
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   192
      Index           =   4
      Left            =   48
      TabIndex        =   10
      Top             =   1092
      Width           =   864
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   192
      Index           =   0
      Left            =   48
      TabIndex        =   8
      Top             =   132
      Width           =   852
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   192
      Index           =   1
      Left            =   48
      TabIndex        =   7
      Top             =   456
      Width           =   864
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   192
      Index           =   2
      Left            =   48
      TabIndex        =   6
      Top             =   768
      Width           =   864
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   192
      Index           =   3
      Left            =   48
      TabIndex        =   5
      Top             =   1416
      Width           =   840
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmConnect.frm - Connect to a Server.

Option Explicit

Private Sub cmdConnect_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConnect.cmdConnect_Click()", etFullDebug

Dim szOriConns(11) As String
Dim szNewConns() As String
Dim iMax As Integer
Dim X As Integer
Dim objNode As Node
Dim vData

  StartMsg "Connecting to " & txtServer.Text & "..."
  
  'Connect the Server Object
  frmMain.svr.MasterDB = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Master DB", "template1")
  frmMain.svr.Connect txtServer.Text, Val(txtPort.Text), txtUsername.Text, txtPassword.Text
  frmMain.Caption = txtServer.Text & ":" & Val(txtPort.Text) & "- " & App.Title
  ctx.dbVer = frmMain.svr.dbVersion.VersionNum
  
  'Write the Values for later
  With ctx
    .Username = txtUsername.Text
    .Password = txtPassword.Text
    .Server = txtServer.Text
    .Port = Val(txtPort.Text)
    .Description = txtDescription.Text
  End With
  
  'Maintain the connection list. Make the current connection the first. Lose #10 if necessary
  szOriConns(0) = ctx.Username & "|" & ctx.Server & "|" & ctx.Port & "|" & ctx.Description
  szOriConns(1) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 1", "")
  szOriConns(2) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 2", "")
  szOriConns(3) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 3", "")
  szOriConns(4) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 4", "")
  szOriConns(5) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 5", "")
  szOriConns(6) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 6", "")
  szOriConns(7) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 7", "")
  szOriConns(8) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 8", "")
  szOriConns(9) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 9", "")
  szOriConns(10) = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 10", "")
 
  'Drop any entries that are the same as the current one.
  For X = 1 To 10
    vData = Split(szOriConns(X), "|")
    If UBound(vData) >= 0 Then
      If vData(0) = ctx.Username And vData(1) = ctx.Server And vData(2) = ctx.Port Then
        szOriConns(X) = ""
      End If
    End If
  Next X
  
  szNewConns = Filter(szOriConns, "|", True)
  
  'Save the connections and rebuild the button menu.
  frmMain.tb.Buttons(1).ButtonMenus.Clear
  iMax = UBound(szNewConns)
  If iMax > 9 Then iMax = 9
  For X = 0 To iMax
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection " & X + 1, regString, szNewConns(X)
  Next X
  BuildConnectionMenu
  
  With frmMain
    'enable menu options
    .mnuFileChangePassword.Enabled = True
    .mnuPopupRefresh.Enabled = True
    .mnuPopupCreate.Enabled = True
    .mnuPopupCreateDatabase.Enabled = True
    .mnuPopupCreateGroup.Enabled = True
    .mnuPopupCreateUser.Enabled = True
    .mnuPopupProperties.Enabled = True
    .mnuToolsFindObject.Enabled = True
  
    'Enable buttons on the toolbar
    With .tb
      .Buttons("refresh").Enabled = True
      .Buttons("create").Enabled = True
      .Buttons("create").ButtonMenus("database").Enabled = True
      .Buttons("create").ButtonMenus("group").Enabled = True
      .Buttons("create").ButtonMenus("user").Enabled = True
      .Buttons("properties").Enabled = True
    End With
  End With
 
  'Rebuild the Plugins Menu
  BuildPluginsMenu
  
  'Start populating the treeview.
  With frmMain
    .tv.Nodes.Clear
    .lv.ListItems.Clear
    .lv.ColumnHeaders.Clear
    Set objNode = .tv.Nodes.Add(, , "SVR-" & GetID, .svr.Server, "server")
    Set .svr.Tag = objNode
  End With
  
  'Set the CurrentObject
  Set ctx.CurrentObject = frmMain.svr
  ctx.CurrentDB = ""
 
  'Expand the node
  frmMain.tv_NodeClick objNode
  objNode.Expanded = True
  
  'Unload the form
  EndMsg
  Unload Me
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConnect.cmdConnect_Click", False
End Sub

Public Sub Load_Defaults(Optional Connection As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConnect.Load_Defaults(" & Connection & ")", etFullDebug

Dim szConnection() As String

  PatchForm Me
  
  'If no connection was specified, then assume connection 1.
  If Connection = 0 Then
    szConnection = Split(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 1", "postgres|localhost|5432|local connection"), "|")
  Else
    szConnection = Split(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection " & Connection, "postgres|localhost|5432|local connection"), "|")
  End If
  txtUsername.Text = szConnection(0)
  txtServer.Text = szConnection(1)
  txtPort.Text = szConnection(2)
  If UBound(szConnection) > 2 Then txtDescription.Text = szConnection(3)

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConnect.Load_Defaults"
End Sub

