VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to Server"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
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
      Top             =   1035
      Width           =   2625
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   2430
      TabIndex        =   1
      ToolTipText     =   "Connect to the Specified Server."
      Top             =   1395
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   450
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   765
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   690
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmConnect.frm - Connect to a Server.

Option Explicit

Private Sub cmdConnect_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConnect.cmdConnect_Click()", etFullDebug

Dim szOriConns(11) As String
Dim szNewConns() As String
Dim iMax As Integer
Dim X As Integer
Dim objNode As Node

  StartMsg "Connecting to " & txtServer.Text & "..."
  
  'Connect the Server Object
  frmMain.svr.Connect txtServer.Text, Val(txtPort.Text), txtUsername.Text, txtPassword.Text
  frmMain.Caption = txtServer.Text & ":" & Val(txtPort.Text) & "- " & App.Title
  
  'Write the Values for later
  ctx.Username = txtUsername.Text
  ctx.Password = txtPassword.Text
  ctx.Server = txtServer.Text
  ctx.Port = Val(txtPort.Text)
  
  'Maintain the connection list. Make the current connection the first. Lose #10 if necessary
  szOriConns(0) = ctx.Username & "|" & ctx.Server & "|" & ctx.Port
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
    If szOriConns(X) = ctx.Username & "|" & ctx.Server & "|" & ctx.Port Then
      szOriConns(X) = ""
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
  
  'enable menu options
  frmMain.mnuFileChangePassword.Visible = True
  frmMain.mnuPopupRefresh.Visible = True
  frmMain.mnuPopupSep1.Visible = True
  frmMain.mnuPopupCreate.Visible = True
  frmMain.mnuPopupCreateDatabase.Visible = True
  frmMain.mnuPopupCreateGroup.Visible = True
  frmMain.mnuPopupCreateUser.Visible = True
  frmMain.mnuPopupProperties.Visible = True
  
  'Enable buttons on the toolbar
  frmMain.tb.Buttons("refresh").Visible = True
  frmMain.tb.Buttons("sep1").Visible = True
  frmMain.tb.Buttons("create").Visible = True
  frmMain.tb.Buttons("create").ButtonMenus("database").Visible = True
  frmMain.tb.Buttons("create").ButtonMenus("group").Visible = True
  frmMain.tb.Buttons("create").ButtonMenus("user").Visible = True
  frmMain.tb.Buttons("properties").Visible = True
 
  'Rebuild the Plugins Menu
  BuildPluginsMenu
  
  'Start populating the treeview.
  frmMain.tv.Nodes.Clear
  frmMain.lv.ListItems.Clear
  frmMain.lv.ColumnHeaders.Clear
  Set objNode = frmMain.tv.Nodes.Add(, , "SVR-" & GetID, frmMain.svr.Server, "server")
  
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
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConnect.cmdConnect_Click"
End Sub

Public Sub Load_Defaults(Optional Connection As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConnect.Load_Defaults(" & Connection & ")", etFullDebug

Dim szConnection() As String
  'If no connection was specified, then assume connection 1.
  If Connection = 0 Then
    szConnection = Split(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection 1", "postgres|localhost|5432"), "|")
  Else
    szConnection = Split(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection " & Connection, "postgres|localhost|5432"), "|")
  End If
  txtUsername.Text = szConnection(0)
  txtServer.Text = szConnection(1)
  txtPort.Text = szConnection(2)
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConnect.Load_Defaults"
End Sub

