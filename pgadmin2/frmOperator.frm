VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmOperator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operator"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmOperator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   10
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin MSComctlLib.ImageList il 
      Left            =   0
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
            Picture         =   "frmOperator.frx":058A
            Key             =   "function"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperator.frx":0B24
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperator.frx":10BE
            Key             =   "type"
         EndProperty
      EndProperty
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
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Properties 1"
      TabPicture(0)   =   "frmOperator.frx":1658
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblProperties(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProperties(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboProperties(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "hbxProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboProperties(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProperties(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProperties(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProperties(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtProperties(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "P&roperties 2"
      TabPicture(1)   =   "frmOperator.frx":1674
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblProperties(8)"
      Tab(1).Control(1)=   "lblProperties(9)"
      Tab(1).Control(2)=   "lblProperties(10)"
      Tab(1).Control(3)=   "lblProperties(11)"
      Tab(1).Control(4)=   "lblProperties(12)"
      Tab(1).Control(5)=   "lblProperties(13)"
      Tab(1).Control(6)=   "cboProperties(9)"
      Tab(1).Control(7)=   "cboProperties(8)"
      Tab(1).Control(8)=   "cboProperties(7)"
      Tab(1).Control(9)=   "cboProperties(6)"
      Tab(1).Control(10)=   "cboProperties(5)"
      Tab(1).Control(11)=   "cboProperties(4)"
      Tab(1).Control(12)=   "chkProperties(0)"
      Tab(1).ControlCount=   13
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Hashes?"
         Height          =   195
         Index           =   0
         Left            =   -74865
         TabIndex        =   16
         ToolTipText     =   "Indicates this operator can support a hash join."
         Top             =   2340
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The operators owner."
         Top             =   1485
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The operators OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the operator."
         Top             =   675
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "The kind of operator."
         Top             =   3105
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "The function used to implement this operator."
         Top             =   1845
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
         Height          =   2130
         Index           =   0
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   "Comments about the operator."
         Top             =   3915
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   3757
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
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   1
         Left            =   1935
         TabIndex        =   5
         ToolTipText     =   "The type of the left-hand argument of the operator, if any. This option would be omitted for a left-unary operator. "
         Top             =   2250
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
         Index           =   2
         Left            =   1935
         TabIndex        =   6
         ToolTipText     =   "The type of the right-hand argument of the operator, if any. This option would be omitted for a right-unary operator."
         Top             =   2655
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
         Index           =   4
         Left            =   -73065
         TabIndex        =   12
         ToolTipText     =   "The commutator of this operator."
         Top             =   630
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
         Index           =   5
         Left            =   -73065
         TabIndex        =   13
         ToolTipText     =   "The negator of this operator."
         Top             =   1035
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
         Index           =   6
         Left            =   -73065
         TabIndex        =   14
         ToolTipText     =   "The restriction selectivity estimator function for this operator."
         Top             =   1440
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
         Index           =   7
         Left            =   -73065
         TabIndex        =   15
         ToolTipText     =   "The join selectivity estimator function for this operator."
         Top             =   1845
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
         Index           =   8
         Left            =   -73065
         TabIndex        =   17
         ToolTipText     =   "If this operator can support a merge join, the operator that sorts the left-hand data type of this operator. "
         Top             =   2655
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
         Index           =   9
         Left            =   -73065
         TabIndex        =   18
         ToolTipText     =   "If this operator can support a merge join, the operator that sorts the right-hand data type of this operator."
         Top             =   3060
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
         Index           =   3
         Left            =   1935
         TabIndex        =   8
         ToolTipText     =   "The type of the right-hand argument of the operator, if any. This option would be omitted for a right-unary operator."
         Top             =   3465
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Right sort operator"
         Height          =   195
         Index           =   13
         Left            =   -74865
         TabIndex        =   32
         Top             =   3150
         Width           =   1305
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Left sort operator"
         Height          =   195
         Index           =   12
         Left            =   -74865
         TabIndex        =   31
         Top             =   2745
         Width           =   1200
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Join function"
         Height          =   195
         Index           =   11
         Left            =   -74865
         TabIndex        =   30
         Top             =   1935
         Width           =   900
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Restrict function"
         Height          =   195
         Index           =   10
         Left            =   -74865
         TabIndex        =   29
         Top             =   1530
         Width           =   1155
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Negator"
         Height          =   195
         Index           =   9
         Left            =   -74865
         TabIndex        =   28
         Top             =   1125
         Width           =   570
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Commutator"
         Height          =   195
         Index           =   8
         Left            =   -74865
         TabIndex        =   27
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Operator function"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   26
         Top             =   1935
         Width           =   1230
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   25
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   23
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Left operand type"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   22
         Top             =   2340
         Width           =   1245
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Right operand type"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   21
         Top             =   2745
         Width           =   1350
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Result type"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   20
         Top             =   3555
         Width           =   795
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Operator kind"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   19
         Top             =   3150
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmOperator.frm - Edit/Create a Operator

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim objOperator As pgOperator

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperator.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperator.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperator.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim szFunction As String
Dim szCommutator As String
Dim szNegator As String
Dim szRestrict As String
Dim szJoin As String
Dim szLeftSort As String
Dim szRightSort As String

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Operator name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select a Operator function!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select a left operand type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  If cboProperties(2).Text = "" Then
    MsgBox "You must select a right operand type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(2).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Operator..."
    If Not (cboProperties(0).SelectedItem Is Nothing) Then szFunction = cboProperties(0).SelectedItem.Tag
    If Not (cboProperties(4).SelectedItem Is Nothing) Then szCommutator = cboProperties(4).SelectedItem.Tag
    If Not (cboProperties(5).SelectedItem Is Nothing) Then szNegator = cboProperties(5).SelectedItem.Tag
    If Not (cboProperties(6).SelectedItem Is Nothing) Then szRestrict = cboProperties(6).SelectedItem.Tag
    If Not (cboProperties(7).SelectedItem Is Nothing) Then szJoin = cboProperties(7).SelectedItem.Tag
    If Not (cboProperties(8).SelectedItem Is Nothing) Then szLeftSort = cboProperties(8).SelectedItem.Tag
    If Not (cboProperties(9).SelectedItem Is Nothing) Then szRightSort = cboProperties(9).SelectedItem.Tag
    frmMain.svr.Databases(szDatabase).Operators.Add txtProperties(0).Text, szFunction, cboProperties(1).Text, cboProperties(2).Text, szCommutator, szNegator, szRestrict, szJoin, Bin2Bool(chkProperties(0).Value), szLeftSort, szRightSort, hbxProperties(0).Text
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) <> "SVR-" Then
        If (Left(objNode.Key, 4) = "OPR+") And (objNode.Parent.Text = szDatabase) Then
          frmMain.tv.Nodes.Add objNode.Key, tvwChild, "OPR-" & GetID, txtProperties(0).Text & " (" & cboProperties(1).Text & ", " & cboProperties(2).Text & ")", "Operator"
          objNode.Text = "Operators (" & objNode.Children & ")"
        End If
      End If
    Next objNode
    
  Else
    StartMsg "Updating Operator..."
    If hbxProperties(0).Tag = "Y" Then objOperator.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListOperator
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperator.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, Optional Operator As pgOperator)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperator.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objFunction As pgFunction
Dim objType As pgType
Dim objTempOperator As pgOperator
Dim objItem As ComboItem
Dim vArgument As Variant
  
  szDatabase = szDB
  
  If Operator Is Nothing Then
  
    'Create a new Operator
    bNew = True
    Me.Caption = "Create Operator"
    
    'Load the combos
    For Each objFunction In frmMain.svr.Databases(szDatabase).Functions
      Set objItem = cboProperties(0).ComboItems.Add(, , objFunction.Identifier, "function")
      objItem.Tag = objFunction.Name
      Set objItem = cboProperties(6).ComboItems.Add(, , objFunction.Identifier, "function")
      objItem.Tag = objFunction.Name
      Set objItem = cboProperties(7).ComboItems.Add(, , objFunction.Identifier, "function")
      objItem.Tag = objFunction.Name
    Next objFunction
    For Each objType In frmMain.svr.Databases(szDatabase).Types
      If Left(objType.Name, 1) <> "_" Then cboProperties(1).ComboItems.Add , , objType.Name, "type"
      If Left(objType.Name, 1) <> "_" Then cboProperties(2).ComboItems.Add , , objType.Name, "type"
    Next objType
    For Each objTempOperator In frmMain.svr.Databases(szDatabase).Operators
      Set objItem = cboProperties(4).ComboItems.Add(, , objTempOperator.Identifier, "operator")
      objItem.Tag = objTempOperator.Name
      Set objItem = cboProperties(5).ComboItems.Add(, , objTempOperator.Identifier, "operator")
      objItem.Tag = objTempOperator.Name
      Set objItem = cboProperties(8).ComboItems.Add(, , objTempOperator.Identifier, "operator")
      objItem.Tag = objTempOperator.Name
      Set objItem = cboProperties(9).ComboItems.Add(, , objTempOperator.Identifier, "operator")
      objItem.Tag = objTempOperator.Name
    Next objTempOperator
  
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    For X = 0 To 9
      cboProperties(X).BackColor = &H80000005
    Next X
    
  Else
  
    'Display/Edit the specified Operator.
    Set objOperator = Operator
    bNew = False

    Me.Caption = "Operator: " & objOperator.Identifier
    txtProperties(0).Text = objOperator.Name
    txtProperties(1).Text = objOperator.OID
    txtProperties(2).Text = objOperator.Owner
    txtProperties(3).Text = objOperator.Kind
    Set objItem = cboProperties(0).ComboItems.Add(, , objOperator.OperatorFunction, "function")
    objItem.Selected = True
    Set objItem = cboProperties(1).ComboItems.Add(, , objOperator.LeftOperandType, "type")
    objItem.Selected = True
    Set objItem = cboProperties(2).ComboItems.Add(, , objOperator.RightOperandType, "type")
    objItem.Selected = True
    Set objItem = cboProperties(3).ComboItems.Add(, , objOperator.ResultType, "type")
    objItem.Selected = True
    Set objItem = cboProperties(4).ComboItems.Add(, , objOperator.Commutator, "operator")
    objItem.Selected = True
    Set objItem = cboProperties(5).ComboItems.Add(, , objOperator.Negator, "operator")
    objItem.Selected = True
    Set objItem = cboProperties(6).ComboItems.Add(, , objOperator.RestrictFunction, "function")
    objItem.Selected = True
    Set objItem = cboProperties(7).ComboItems.Add(, , objOperator.JoinFunction, "function")
    objItem.Selected = True
    Set objItem = cboProperties(8).ComboItems.Add(, , objOperator.LeftTypeSortOperator, "function")
    objItem.Selected = True
    Set objItem = cboProperties(9).ComboItems.Add(, , objOperator.RightTypeSortOperator, "function")
    objItem.Selected = True

    chkProperties(0).Value = Bool2Bin(objOperator.HashJoins)
    hbxProperties(0).Text = objOperator.Comment
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperator.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperator.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperator.hbxProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperator.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objOperator Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objOperator.HashJoins)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperator.chkProperties_Click"
End Sub



