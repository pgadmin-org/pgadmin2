VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index"
   ClientHeight    =   6876
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5520
   Icon            =   "frmIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6876
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
      TabPicture(0)   =   "frmIndex.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "hbxProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkProperties(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lvProperties(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "hbxProperties(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin HighlightBox.HBX hbxProperties 
         Height          =   870
         Index           =   1
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   "Comments about the index."
         Top             =   5310
         Width           =   5190
         _ExtentX        =   9165
         _ExtentY        =   1545
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Comments"
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   1140
         Index           =   0
         Left            =   1935
         TabIndex        =   7
         ToolTipText     =   "Lists the indexed columns."
         Top             =   3195
         Width           =   3390
         _ExtentX        =   5990
         _ExtentY        =   2011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Columns"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Unique Index?"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   6
         ToolTipText     =   "Is the index Unique?"
         Top             =   2835
         Width           =   2040
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Primary Index?"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   5
         ToolTipText     =   "Is the index a Primary Key?"
         Top             =   2430
         Width           =   2040
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The indexes OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         Index           =   0
         Left            =   1935
         TabIndex        =   1
         ToolTipText     =   "The name of the index."
         Top             =   675
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   300
         Index           =   0
         Left            =   1932
         TabIndex        =   3
         ToolTipText     =   "The indexed table."
         Top             =   1488
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   300
         Index           =   1
         Left            =   1932
         TabIndex        =   4
         ToolTipText     =   $"frmIndex.frx":05A6
         Top             =   1932
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   825
         Index           =   0
         Left            =   135
         TabIndex        =   8
         ToolTipText     =   "Defines the constraint expression for a partial index (PostgreSQL 7.2+)."
         Top             =   4410
         Width           =   5190
         _ExtentX        =   9165
         _ExtentY        =   1461
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         Caption         =   "Constraint"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Indexed Columns"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   16
         Top             =   3285
         Width           =   1215
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Index Type"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   15
         Top             =   2025
         Width           =   795
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   1575
         Width           =   405
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
   End
   Begin MSComctlLib.ImageList il 
      Left            =   0
      Top             =   6300
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":06D6
            Key             =   "column"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":0C70
            Key             =   "table"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndex.frx":0DCA
            Key             =   "index"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmIndex.frm - Edit/Create a Index

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objIndex As pgIndex

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.cmdOK_Click()", etFullDebug

Dim szOldName As String
Dim objNode As Node
Dim objItem As ListItem
Dim objNewIndex As pgIndex
Dim szColumns As String

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Index name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select a table!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    For Each objItem In lvProperties(0).ListItems
      If objItem.Checked = True Then szColumns = szColumns & QUOTE & objItem.Text & QUOTE & ", "
    Next objItem
    If Len(szColumns) > 2 Then szColumns = Left(szColumns, Len(szColumns) - 2)
    If szColumns = "" Then
      MsgBox "You must select at least one column!", vbExclamation, "Error"
      tabProperties.Tab = 0
      lvProperties(0).SetFocus
      Exit Sub
    End If
    
    StartMsg "Creating Index..."
    Set objNewIndex = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).SelectedItem.Tag.Identifier).Indexes.Add(txtProperties(0).Text, Bin2Bool(chkProperties(1).Value), szColumns, cboProperties(1).Text, hbxProperties(1).Text, hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    On Error Resume Next
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).SelectedItem.Tag.Identifier).Indexes.Tag
    Set objNewIndex.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "IND-" & GetID, txtProperties(0).Text, "index")
    objNode.Text = "Indexes (" & objNode.Children & ")"
    If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
    
  Else
    StartMsg "Updating Index..."
    
    'Update the index name if required
    If txtProperties(0).Tag = "Y" Then
      szOldName = objIndex.Name
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).Text).Indexes.Rename szOldName, txtProperties(0).Text
        
      'Update the node text
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).Text).Indexes(txtProperties(0).Text).Tag.Text = txtProperties(0).Text
    End If
    
    If hbxProperties(1).Tag = "Y" Then objIndex.Comment = hbxProperties(1).Text
  End If
  
  'Simulate a node click to refresh the ListIndex
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional Index As pgIndex)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objTable As pgTable
Dim objItem As ComboItem
Dim objColumn As pgColumn
Dim vColumn As Variant
Dim vArgument As Variant
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  If Index Is Nothing Then
  
    'Create a new Index
    bNew = True
    Me.Caption = "Create Index"
    
    'Load the combos
    For Each objTable In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables
      If Not objTable.SystemObject Then
        Set objItem = cboProperties(0).ComboItems.Add(, , objTable.FormattedID, "table")
        Set objItem.Tag = objTable
      End If
    Next objTable
    Set objItem = cboProperties(1).ComboItems.Add(, , "btree", "index")
    objItem.Selected = True
    cboProperties(1).ComboItems.Add , , "rtree", "index"
    cboProperties(1).ComboItems.Add , , "hash", "index"
    'Gist indexes are in 7.2
    If ctx.dbVer > 7.1 Then
      cboProperties(1).ComboItems.Add , , "gist", "index"
    End If

    'Unlock the edittable fields
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    lvProperties(0).BackColor = &H80000005
    If ctx.dbVer >= 7.2 Then
      hbxProperties(0).BackColor = &H80000005
      hbxProperties(0).Locked = False
    End If
    
  Else
  
    'Display/Edit the specified Index.
    Set objIndex = Index
    bNew = False

    Me.Caption = "Index: " & objIndex.Identifier
    txtProperties(0).Text = objIndex.Name
    txtProperties(1).Text = objIndex.Oid
    Set objItem = cboProperties(0).ComboItems.Add(, , objIndex.Table, "table")
    objItem.Selected = True
    Set objItem = cboProperties(1).ComboItems.Add(, , objIndex.IndexType, "index")
    objItem.Selected = True
    chkProperties(0).Value = Bool2Bin(objIndex.Primary)
    chkProperties(1).Value = Bool2Bin(objIndex.Unique)
    For Each objColumn In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(objIndex.Table).Columns
      If Not objColumn.SystemObject Then lvProperties(0).ListItems.Add , objColumn.Identifier, objColumn.Identifier, "column", "column"
    Next objColumn
    For Each vColumn In objIndex.IndexedColumns
      lvProperties(0).ListItems(vColumn).Checked = True
      lvProperties(0).ListItems(vColumn).Tag = "Y"
    Next vColumn
    hbxProperties(0).Text = objIndex.Constraint
    hbxProperties(1).Text = objIndex.Comment
  End If
  
  'Reset the Tags
  txtProperties(0).Tag = "N"
  hbxProperties(1).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.hbxProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objIndex Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objIndex.Primary)
    chkProperties(1).Value = Bool2Bin(objIndex.Unique)
  Else
    chkProperties(0).Value = 0
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.chkProperties_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.cboProperties_Click(" & Index & ")", etFullDebug

Dim objColumn As pgColumn

  If (Index = 0) And (objIndex Is Nothing) Then
    lvProperties(0).ListItems.Clear
    For Each objColumn In cboProperties(Index).SelectedItem.Tag.Columns
      If Not objColumn.SystemObject Then lvProperties(0).ListItems.Add , , objColumn.Identifier, "column", "column"
    Next objColumn
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.cboProperties_Click"
End Sub

Private Sub lvProperties_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmIndex.lvProperties_ItemCheck(" & Index & ", " & QUOTE & Item.Text & QUOTE & ")", etFullDebug

  If Not (objIndex Is Nothing) Then
    If Item.Tag = "Y" Then
      Item.Checked = True
    Else
      Item.Checked = False
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmIndex.cboProperties_Click"
End Sub
