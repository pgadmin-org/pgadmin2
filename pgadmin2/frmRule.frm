VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmRule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rule"
   ClientHeight    =   6870
   ClientLeft      =   6315
   ClientTop       =   2970
   ClientWidth     =   5520
   Icon            =   "frmRule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   9
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   6360
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmRule.frx":06C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "hbxProperties(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "hbxProperties(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "hbxProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Do Instead?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "Will the new action suppress the original query?"
         Top             =   3375
         Width           =   1995
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   915
         Index           =   0
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   "Any SQL boolean-condition expression. The condition expression may not refer to any tables except new and old. "
         Top             =   2340
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1614
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Caption         =   "Condition"
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the rule."
         Top             =   675
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The rules OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   3
         ToolTipText     =   "Object is either table or table.column. (Currently, only the table form is actually implemented.)"
         Top             =   1485
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   1
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "The Event that will cause the rule to be invoked. Event is one of SELECT, UPDATE, DELETE or INSERT."
         Top             =   1890
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   915
         Index           =   1
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "The query or queries making up the action can be any SQL SELECT, INSERT, UPDATE, DELETE, or NOTIFY statement."
         Top             =   3690
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   1614
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Caption         =   "Action"
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   1410
         Index           =   2
         Left            =   135
         TabIndex        =   8
         ToolTipText     =   "Comments about the rule."
         Top             =   4725
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2487
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Comments"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Event"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Object"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1575
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   45
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRule.frx":06DE
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRule.frx":0838
            Key             =   "event"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRule.frx":1112
            Key             =   "view"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmRule.frm - Edit/Create a Rule

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objRule As pgRule

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRule.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRule.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRule.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewRule As pgRule

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Rule name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select an object!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select an event!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  If hbxProperties(1).Text = "" Then hbxProperties(1).Text = "NOTHING"
  
  If bNew Then
    StartMsg "Creating Rule..."
    Set objNewRule = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).SelectedItem.Tag.Identifier).Rules.Add(txtProperties(0).Text, cboProperties(1).Text, hbxProperties(0).Text, Bin2Bool(chkProperties(0).Value), hbxProperties(1).Text, hbxProperties(2).Text)
    
    'Add a new node and update the text on the parent
    'verify if rule is for table or view
    On Error Resume Next
    If frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables.Exists(cboProperties(0).SelectedItem.Tag.Identifier) Then
      Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).SelectedItem.Tag.Identifier).Rules.Tag
    ElseIf frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views.Exists(cboProperties(0).SelectedItem.Tag.Identifier) Then
      Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views(cboProperties(0).SelectedItem.Tag.Identifier).Rules.Tag
    End If
    Set objNewRule.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "RUL-" & GetID, txtProperties(0).Text, "rule")
    objNode.Text = "Rules (" & objNode.Children & ")"
    If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler

  Else
    StartMsg "Updating Rule..."
    If hbxProperties(2).Tag = "Y" Then objRule.Comment = hbxProperties(2).Text
  End If
  
  'Simulate a node click to refresh the ListRule
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRule.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional Rule As pgRule)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRule.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objTable As pgTable
Dim objView As pgView
Dim objItem As ComboItem
Dim vArgument As Variant
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  hbxProperties(0).Wordlist = ctx.AutoHighlight
  hbxProperties(1).Wordlist = ctx.AutoHighlight
  
  If Rule Is Nothing Then
  
    'Create a new Rule
    bNew = True
    Me.Caption = "Create Rule"
    
    'Load the combos
    For Each objTable In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables
      If Not objTable.SystemObject Then
        Set objItem = cboProperties(0).ComboItems.Add(, "TBL-" & GetID, objTable.FormattedID, "table")
        Set objItem.Tag = objTable
      End If
    Next objTable
    For Each objView In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Views
      If Not objView.SystemObject Then
        Set objItem = cboProperties(0).ComboItems.Add(, "VIE-" & GetID, objView.FormattedID, "view")
        Set objItem.Tag = objView
      End If
    Next objView
    cboProperties(1).ComboItems.Add , , "INSERT", "event"
    cboProperties(1).ComboItems.Add , , "UPDATE", "event"
    cboProperties(1).ComboItems.Add , , "DELETE", "event"
    cboProperties(1).ComboItems.Add , , "SELECT", "event"

    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    hbxProperties(0).BackColor = &H80000005
    hbxProperties(0).Locked = False
    hbxProperties(1).BackColor = &H80000005
    hbxProperties(1).Locked = False
    
  Else
  
    'Display/Edit the specified Rule.
    Set objRule = Rule
    bNew = False

    Me.Caption = "Rule: " & objRule.Identifier
    txtProperties(0).Text = objRule.Name
    txtProperties(1).Text = objRule.Oid
    
    If frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables.Exists(objRule.Table) Then
      Set objItem = cboProperties(0).ComboItems.Add(, , objRule.Table, "table")
    Else
      Set objItem = cboProperties(0).ComboItems.Add(, , objRule.Table, "view")
    End If
    objItem.Selected = True
    Set objItem = cboProperties(1).ComboItems.Add(, , objRule.RuleEvent, "event")
    objItem.Selected = True
    hbxProperties(0).Text = objRule.Condition
    hbxProperties(1).Text = objRule.Action
    chkProperties(0).Value = Bool2Bin(objRule.DoInstead)
    hbxProperties(2).Text = objRule.Comment
  End If
  
  'Reset the Tags
  hbxProperties(2).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRule.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRule.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRule.hbxProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmRule.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objRule Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objRule.DoInstead)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmRule.chkProperties_Click"
End Sub
