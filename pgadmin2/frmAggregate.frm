VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmAggregate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aggregate"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmAggregate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
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
      TabPicture(0)   =   "frmAggregate.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblProperties(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblProperties(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProperties(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboProperties(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboProperties(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboProperties(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboProperties(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "hbxProperties(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProperties(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtProperties(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtProperties(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtProperties(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the aggregate."
         Top             =   630
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The aggregates OID (Object ID) in the PostgreSQL Database."
         Top             =   1035
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The aggregates owner."
         Top             =   1440
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   $"frmAggregate.frx":0166
         Top             =   3870
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   1860
         Index           =   0
         Left            =   135
         TabIndex        =   10
         ToolTipText     =   "Comments about the aggregate."
         Top             =   4275
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   3281
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
         Index           =   0
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   $"frmAggregate.frx":0218
         Top             =   1800
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
         TabIndex        =   5
         ToolTipText     =   "The data type for the aggregate's state value."
         Top             =   2205
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
         ToolTipText     =   $"frmAggregate.frx":02C8
         Top             =   2610
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
         TabIndex        =   7
         ToolTipText     =   "The data type returned by the aggregate."
         Top             =   3015
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
         Left            =   1935
         TabIndex        =   8
         ToolTipText     =   $"frmAggregate.frx":04CB
         Top             =   3420
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
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   21
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   20
         Top             =   1080
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   675
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "State function"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   18
         Top             =   2700
         Width           =   990
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "State type"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   17
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Input type"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Initial condition"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   15
         Top             =   3915
         Width           =   1050
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Final function"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   14
         Top             =   3510
         Width           =   945
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Final type"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   13
         Top             =   3105
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   12
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAggregate.frx":0650
            Key             =   "function"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAggregate.frx":0BEA
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAggregate.frx":1184
            Key             =   "any"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAggregate.frx":12DE
            Key             =   "domain"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAggregate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmAggregate.frm - Edit/Create a Aggregate

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objAggregate As pgAggregate

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAggregate.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAggregate.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAggregate.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewAggregate As pgAggregate
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Aggregate name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select an input type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select a state type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  If cboProperties(2).Text = "" Then
    MsgBox "You must select a state function!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(2).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Aggregate..."
    Set objNewAggregate = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Aggregates.Add(txtProperties(0).Text, cboProperties(0).Text, cboProperties(2).Text, cboProperties(1).Text, cboProperties(4).Text, txtProperties(3).Text, hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Aggregates.Tag
    If cboProperties(0).Text = "ANY" Then
      frmMain.tv.Nodes.Add objNode.Key, tvwChild, "AGG-" & GetID, txtProperties(0).Text & " opaque", "aggregate"
    Else
      frmMain.tv.Nodes.Add objNode.Key, tvwChild, "AGG-" & GetID, txtProperties(0).Text & " " & cboProperties(0).Text, "aggregate"
    End If
    objNode.Text = "Aggregates (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Aggregate..."
    If hbxProperties(0).Tag = "Y" Then objAggregate.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListAggregate
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAggregate.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional Aggregate As pgAggregate)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAggregate.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objType As pgType
Dim objDomain As pgDomain
Dim objFunction As pgFunction
Dim objNamespace As pgNamespace
Dim objItem As ComboItem
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  If Aggregate Is Nothing Then
  
    'Create a new Aggregate
    bNew = True
    Me.Caption = "Create Aggregate"
  
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    txtProperties(3).BackColor = &H80000005
    txtProperties(3).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    cboProperties(2).BackColor = &H80000005
    cboProperties(4).BackColor = &H80000005
  
    'Load the combos
    cboProperties(0).ComboItems.Add , , "ANY", "any"
    If ctx.dbVer >= 7.3 Then
      'Load pg_catalog entries first, unqualified
      For Each objDomain In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Domains
        cboProperties(0).ComboItems.Add , , fmtID(objDomain.Name), "domain"
        cboProperties(1).ComboItems.Add , , fmtID(objDomain.Name), "domain"
      Next objDomain
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
        If Left(objType.Name, 1) <> "_" Then cboProperties(0).ComboItems.Add , , fmtID(objType.Name), "type"
        If Left(objType.Name, 1) <> "_" Then cboProperties(1).ComboItems.Add , , fmtID(objType.Name), "type"
      Next objType
      'Now load the rest
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objDomain In objNamespace.Domains
            cboProperties(0).ComboItems.Add , , objDomain.FormattedID, "domain"
            cboProperties(1).ComboItems.Add , , objDomain.FormattedID, "domain"
          Next objDomain
          For Each objType In objNamespace.Types
            If Left(objType.Name, 1) <> "_" Then cboProperties(0).ComboItems.Add , , objType.FormattedID, "type"
            If Left(objType.Name, 1) <> "_" Then cboProperties(1).ComboItems.Add , , objType.FormattedID, "type"
          Next objType
        End If
      Next objNamespace
    Else
      For Each objDomain In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Domains
        cboProperties(0).ComboItems.Add , , objDomain.FormattedID, "domain"
        cboProperties(1).ComboItems.Add , , objDomain.FormattedID, "domain"
      Next objDomain
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types
        If Left(objType.Name, 1) <> "_" Then cboProperties(0).ComboItems.Add , , objType.FormattedID, "type"
        If Left(objType.Name, 1) <> "_" Then cboProperties(1).ComboItems.Add , , objType.FormattedID, "type"
      Next objType
    End If
  
  Else
  
    'Display/Edit the specified Aggregate.
    Set objAggregate = Aggregate
    bNew = False
    
    Me.Caption = "Aggregate: " & objAggregate.Identifier
    txtProperties(0).Text = objAggregate.Name
    txtProperties(1).Text = objAggregate.Oid
    txtProperties(2).Text = objAggregate.Owner
    Set objItem = cboProperties(0).ComboItems.Add(, , objAggregate.InputType, "type")
    objItem.Selected = True
    Set objItem = cboProperties(1).ComboItems.Add(, , objAggregate.StateType, "type")
    objItem.Selected = True
    Set objItem = cboProperties(2).ComboItems.Add(, , objAggregate.StateFunction, "function")
    objItem.Selected = True
    Set objItem = cboProperties(3).ComboItems.Add(, , objAggregate.FinalType, "type")
    objItem.Selected = True
    Set objItem = cboProperties(4).ComboItems.Add(, , objAggregate.FinalFunction, "function")
    objItem.Selected = True
    txtProperties(3).Text = objAggregate.InitialCondition
    hbxProperties(0).Text = objAggregate.Comment
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAggregate.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAggregate.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAggregate.hbxProperties_Change"
End Sub

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmAggregate.cboProperties_Click(" & Index & ")", etFullDebug

Dim objFunction As pgFunction
Dim objNamespace As pgNamespace

  'Populate the StateFunction Combo. StateFunctions have 1 or 2 arguments, the first of which
  'is the StateType. The second is the InputType The return type must also be the StateType.
  'For the FinalFunction Combo, FinalFunction have 1 argument = StateType
  If (Index = 1) And (objAggregate Is Nothing) Then
    cboProperties(2).ComboItems.Clear
    cboProperties(2).Text = ""
    cboProperties(4).ComboItems.Clear
    cboProperties(4).Text = ""
    
    If ctx.dbVer >= 7.3 Then
    
      'Add pg_catalog items first, unqualified
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Functions
        If objFunction.Arguments.Count >= 1 Then
          'StateFunction - Single argument functions
          If objFunction.Arguments.Count = 1 Then
            If objFunction.Arguments(1) = cboProperties(1).Text And objFunction.Returns = cboProperties(1).Text Then
              cboProperties(2).ComboItems.Add , , fmtID(objFunction.Name), "function"
            End If
          End If
          'StateFunction - Double argument functions
          If objFunction.Arguments.Count = 2 Then
            If objFunction.Arguments(1) = cboProperties(1).Text And objFunction.Arguments(2) = cboProperties(0).Text And objFunction.Returns = cboProperties(1).Text Then
              cboProperties(2).ComboItems.Add , , fmtID(objFunction.Name), "function"
            End If
          End If
          'FinalFunction
          If objFunction.Arguments(1) = cboProperties(1).Text Then
            cboProperties(4).ComboItems.Add , , fmtID(objFunction.Name), "function"
          End If
        End If
      Next objFunction
      
      'Now add other items
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objFunction In objNamespace.Functions
            If objFunction.Arguments.Count >= 1 Then
              'StateFunction - Single argument functions
              If objFunction.Arguments.Count = 1 Then
                If objFunction.Arguments(1) = cboProperties(1).Text And objFunction.Returns = cboProperties(1).Text Then
                  cboProperties(2).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
                End If
              End If
              'StateFunction - Double argument functions
              If objFunction.Arguments.Count = 2 Then
                If objFunction.Arguments(1) = cboProperties(1).Text And objFunction.Arguments(2) = cboProperties(0).Text And objFunction.Returns = cboProperties(1).Text Then
                  cboProperties(2).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
                End If
              End If
              'FinalFunction
              If objFunction.Arguments(1) = cboProperties(1).Text Then
                cboProperties(4).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
              End If
            End If
          Next objFunction
        End If
      Next objNamespace
      
    Else
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions
        If objFunction.Arguments.Count >= 1 Then
        
          'StateFunction - Single argument functions
          If objFunction.Arguments.Count = 1 Then
            If objFunction.Arguments(1) = cboProperties(1).Text And objFunction.Returns = cboProperties(1).Text Then
              cboProperties(2).ComboItems.Add , , objFunction.Name, "function"
            End If
          End If
          
          'StateFunction - Double argument functions
          If objFunction.Arguments.Count = 2 Then
            If objFunction.Arguments(1) = cboProperties(1).Text And objFunction.Arguments(2) = cboProperties(0).Text And objFunction.Returns = cboProperties(1).Text Then
              cboProperties(2).ComboItems.Add , , objFunction.Name, "function"
            End If
          End If
          
          'FinalFunction
          If objFunction.Arguments(1) = cboProperties(1).Text Then
            cboProperties(4).ComboItems.Add , , objFunction.Name, "function"
          End If
        End If
      Next objFunction
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAggregate.cboProperties_Click"
End Sub
