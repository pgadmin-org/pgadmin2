VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmDomain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Domain"
   ClientHeight    =   6870
   ClientLeft      =   7350
   ClientTop       =   1935
   ClientWidth     =   5520
   Icon            =   "frmDomain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDomain.frx":06C2
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDomain.frx":0D94
            Key             =   "type"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   6360
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmDomain.frx":1466
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "hbxProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkProperties(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProperties(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProperties(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   $"frmDomain.frx":1482
         Top             =   3105
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "The defined length of the column."
         Top             =   2295
         Width           =   3390
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Not null?"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         ToolTipText     =   $"frmDomain.frx":1557
         Top             =   3555
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "The domains owner."
         Top             =   1485
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "The domains OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "The name of the domain."
         Top             =   675
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The numeric scale of the column (applicable to numeric columns only)."
         Top             =   2700
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   7
         ToolTipText     =   "The data type of the domain."
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
         Height          =   2175
         Index           =   0
         Left            =   135
         TabIndex        =   8
         ToolTipText     =   "Comments about the operator."
         Top             =   3915
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   3836
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
         Caption         =   "Default"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   18
         Top             =   3150
         Width           =   510
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Base type"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   1935
         Width           =   705
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   10
         Top             =   2340
         Width           =   495
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Numeric scale"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   9
         Top             =   2745
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmDomain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmDomain.frm - Edit/Create a Domain

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objDomain As pgDomain

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDomain.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDomain.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDomain.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewDomain As pgDomain
Dim szInputfunction As String
Dim szOutputFunction As String
Dim szSendFunction As String
Dim szReceiveFunction As String

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Domain name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select a base type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If

  
  If bNew Then
    StartMsg "Creating Domain..."
    Set objNewDomain = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Domains.Add(txtProperties(0).Text, cboProperties(0).Text, Val(txtProperties(3).Text), Val(txtProperties(4).Text), txtProperties(5).Text, Bin2Bool(chkProperties(0).Value), hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Domains.Tag
    Set objNewDomain.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "DOM-" & GetID, txtProperties(0).Text, "domain")
    objNode.Text = "Domains (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Domain..."
    If hbxProperties(0).Tag = "Y" Then objDomain.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListDomain
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDomain.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional Domain As pgDomain)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDomain.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objType As pgType
Dim objTempDomain As pgDomain
Dim objNamespace As pgNamespace
Dim objItem As ComboItem
Dim vArgument As Variant
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  If Domain Is Nothing Then
  
    'Create a new Domain
    bNew = True
    Me.Caption = "Create Domain"
    
    'Load the combos
    If ctx.dbVer >= 7.3 Then
      'Add pg_catalog items first, unqualified
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
        cboProperties(0).ComboItems.Add , , fmtTypeName(objType), "type", "type"
      Next objType
      'Now add other items
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objType In objNamespace.Types
            cboProperties(0).ComboItems.Add , , fmtTypeName(objType), "type", "type"
          Next objType
        End If
      Next objNamespace
    Else
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types
        cboProperties(0).ComboItems.Add , , fmtTypeName(objType), "type", "type"
      Next objType
    End If

    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    txtProperties(5).BackColor = &H80000005
    txtProperties(5).Locked = False
    cboProperties(0).BackColor = &H80000005

  Else
  
    'Display/Edit the specified Domain.
    Set objDomain = Domain
    bNew = False

    Me.Caption = "Domain: " & objDomain.Identifier
    txtProperties(0).Text = objDomain.Name
    txtProperties(1).Text = objDomain.Oid
    txtProperties(2).Text = objDomain.Owner
    If objDomain.Length = 0 Then
      txtProperties(3).Text = "Variable"
    Else
      txtProperties(3).Text = objDomain.Length
    End If
    If objDomain.BaseType = "numeric" Then txtProperties(4).Text = objDomain.NumericScale
    txtProperties(5).Text = objDomain.Default
    Set objItem = cboProperties(0).ComboItems.Add(, , objDomain.BaseType, "type")
    objItem.Selected = True
    chkProperties(0).Value = Bool2Bin(objDomain.NotNull)
    hbxProperties(0).Text = objDomain.Comment
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDomain.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDomain.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDomain.hbxProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDomain.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objDomain Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objDomain.NotNull)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDomain.chkProperties_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDomain.cboProperties_Click(" & Index & ")", etFullDebug

  If Index = 0 Then
     
    'Lock first
    txtProperties(3).BackColor = &H8000000F
    txtProperties(3).Locked = True
    txtProperties(4).BackColor = &H8000000F
    txtProperties(4).Locked = True
    
    'Now unlock based on the data type
    Select Case cboProperties(0).Text
      Case "numeric"
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
        txtProperties(4).BackColor = &H80000005
        txtProperties(4).Locked = False
      Case "char"
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
      Case "varchar"
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
    End Select
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDomain.cboProperties_Click"
End Sub

