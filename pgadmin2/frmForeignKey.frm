VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmForeignKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foreign Key"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmForeignKey.frx":0000
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
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmForeignKey.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProperties(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboProperties(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProperties(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkProperties(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "&Relationships"
      TabPicture(1)   =   "frmForeignKey.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblProperties(7)"
      Tab(1).Control(1)=   "lblProperties(8)"
      Tab(1).Control(2)=   "cboProperties(6)"
      Tab(1).Control(3)=   "cboProperties(5)"
      Tab(1).Control(4)=   "lvProperties(0)"
      Tab(1).Control(5)=   "cmdRemove"
      Tab(1).Control(6)=   "cmdAdd"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74685
         TabIndex        =   13
         ToolTipText     =   "Add the defined relationship."
         Top             =   1575
         Width           =   1320
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74685
         TabIndex        =   14
         ToolTipText     =   "Remove the selected relationship."
         Top             =   1980
         Width           =   1320
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Deferrable?"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "This controls whether the constraint can be deferred to the end of the transaction."
         Top             =   3105
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the foreign key."
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
         ToolTipText     =   "The foreign keys OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   3
         ToolTipText     =   "The table that the foreign key will be part of."
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
         Index           =   1
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "The table referenced by the foreign key."
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
         Index           =   2
         Left            =   1935
         TabIndex        =   5
         ToolTipText     =   "The action to take when a referenced row in the referenced table is being deleted."
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
         Index           =   3
         Left            =   1935
         TabIndex        =   6
         ToolTipText     =   "The action to take when a referenced column in the referenced table is being updated to a new value."
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
         Left            =   1935
         TabIndex        =   8
         ToolTipText     =   $"frmForeignKey.frx":05C2
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
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4515
         Index           =   0
         Left            =   -73065
         TabIndex        =   15
         ToolTipText     =   "Lists the relationships in the foreign key."
         Top             =   1575
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   7964
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Local Column"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "References"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   5
         Left            =   -73065
         TabIndex        =   11
         ToolTipText     =   "Select a column in the local table."
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
         Index           =   6
         Left            =   -73065
         TabIndex        =   12
         ToolTipText     =   "Select the column to be referenced in the referenced table."
         Top             =   1080
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
         Caption         =   "Local column"
         Height          =   195
         Index           =   8
         Left            =   -74865
         TabIndex        =   24
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Referenced column"
         Height          =   195
         Index           =   7
         Left            =   -74865
         TabIndex        =   23
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Initially"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   22
         Top             =   3510
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "On update"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   21
         Top             =   2700
         Width           =   750
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "On delete"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   2295
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Referenced table"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   1890
         Width           =   1230
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   18
         Top             =   1530
         Width           =   405
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   720
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   0
      Top             =   6345
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
            Picture         =   "frmForeignKey.frx":0660
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForeignKey.frx":07BA
            Key             =   "column"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForeignKey.frx":0D54
            Key             =   "relationship"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmForeignKey.frx":12EE
            Key             =   "key"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmForeignKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmForeignKey.frm - Edit/Create a ForeignKey

Option Explicit

Dim szDatabase As String
Dim szMode As String
Dim frmCallingForm As Form
Dim objForeignKey As pgForeignKey

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.cmdRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then Exit Sub
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.cmdAdd_Click()", etFullDebug

Dim objItem As ListItem

  If cboProperties(5).Text = "" Then Exit Sub
  If cboProperties(6).Text = "" Then Exit Sub

  Set objItem = lvProperties(0).ListItems.Add(, , cboProperties(5).Text, "relationship", "relationship")
  objItem.SubItems(1) = cboProperties(6).Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdAdd_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmForeignKey.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objRelItem As ListItem
Dim szOldName As String

  If Not frmCallingForm Is Nothing Then
    If Not frmCallingForm.Visible Then
      MsgBox "The form that called this form has been destroyed!", vbExclamation, "Error"
      Unload Me
      Exit Sub
    End If
  End If
  
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Foreign Key name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select a referenced table!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  If lvProperties(0).ListItems.Count < 1 Then
    MsgBox "You must specify at least one relationship!", vbExclamation, "Error"
    tabProperties.Tab = 1
    cboProperties(5).SetFocus
    Exit Sub
  End If
  
  Select Case szMode
    Case "TA"
      For Each objItem In frmCallingForm.lvProperties(2).ListItems
        If objItem.Text = txtProperties(0).Text Then
          MsgBox "A foreign key with that name already exists!", vbExclamation, "Error"
          tabProperties.Tab = 0
          txtProperties(0).SetFocus
          Exit Sub
        End If
      Next objItem
      
      Set objItem = frmCallingForm.lvProperties(2).ListItems.Add(, , txtProperties(0).Text, "foreignkey", "foreignkey")
      objItem.SubItems(1) = cboProperties(1).Text
      For Each objRelItem In lvProperties(0).ListItems
        objItem.SubItems(2) = objItem.SubItems(2) & QUOTE & objRelItem.Text & QUOTE & ", "
        objItem.SubItems(3) = objItem.SubItems(3) & QUOTE & objRelItem.SubItems(1) & QUOTE & ", "
      Next objRelItem
      If Len(objItem.SubItems(2)) > 2 Then objItem.SubItems(2) = Left(objItem.SubItems(2), Len(objItem.SubItems(2)) - 2)
      If Len(objItem.SubItems(3)) > 2 Then objItem.SubItems(3) = Left(objItem.SubItems(3), Len(objItem.SubItems(3)) - 2)
      objItem.SubItems(4) = cboProperties(2).Text
      objItem.SubItems(5) = cboProperties(3).Text
      If chkProperties(0).Value = 0 Then
        objItem.SubItems(6) = "No"
      Else
        objItem.SubItems(6) = "Yes"
      End If
      objItem.SubItems(7) = cboProperties(4).Text
      
      frmCallingForm.lvProperties(2).Tag = "Y"
      
    Case "MP"
      'we can't update foreign keys...
      
      'Simulate a node click to refresh the List
      frmMain.tv_NodeClick frmMain.tv.SelectedItem
  End Select
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmForeignKey.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szMD As String, Optional ForeignKey As pgForeignKey, Optional frmCF As Form)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.Initialise(" & QUOTE & szDB & QUOTE & ", " & QUOTE & szMD & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ComboItem
Dim objLItem As ListItem
Dim objTable As pgTable
Dim objRelationship As pgRelationship

  szDatabase = szDB
  
  'The mode indicates the way the form is being used:
  'MP = Called from frmMain, viewing the properties of a Foreign Key object.
  'TA = Called from frmTable (ref: frmCallingForm), adding a new Foreign Key object.
  
  szMode = szMD
  Set frmCallingForm = frmCF
    
  Select Case szMode
    Case "TA"
  
      'Create a new ForeignKey
      Me.Caption = "Create Foreign Key"
      
      'Unlock the edittable fields
      txtProperties(0).BackColor = &H80000005
      txtProperties(0).Locked = False
      cboProperties(1).BackColor = &H80000005
      cboProperties(2).BackColor = &H80000005
      cboProperties(3).BackColor = &H80000005
      cboProperties(4).BackColor = &H80000005
      cboProperties(5).BackColor = &H80000005
      cboProperties(6).BackColor = &H80000005
      lvProperties(0).BackColor = &H80000005
      cmdAdd.Enabled = True
      cmdRemove.Enabled = True
      
      Set objItem = cboProperties(0).ComboItems.Add(, , frmCallingForm.txtProperties(0).Text, "table", "table")
      objItem.Selected = True
      
      'Populate the Referenced Tables combo
      For Each objTable In frmMain.svr.Databases(szDatabase).Tables
        If Not objTable.SystemObject Then cboProperties(1).ComboItems.Add , , objTable.Name, "table", "table"
      Next objTable
      
      'Populate the Actions combos
      Set objItem = cboProperties(2).ComboItems.Add(, , "No Action", "key", "key")
      objItem.Selected = True
      Set objItem = cboProperties(3).ComboItems.Add(, , "No Action", "key", "key")
      objItem.Selected = True
      cboProperties(2).ComboItems.Add , , "Restrict", "key", "key"
      cboProperties(3).ComboItems.Add , , "Restrict", "key", "key"
      cboProperties(2).ComboItems.Add , , "Cascade", "key", "key"
      cboProperties(3).ComboItems.Add , , "Cascade", "key", "key"
      cboProperties(2).ComboItems.Add , , "Set Null", "key", "key"
      cboProperties(3).ComboItems.Add , , "Set Null", "key", "key"
      cboProperties(2).ComboItems.Add , , "Set Default", "key", "key"
      cboProperties(3).ComboItems.Add , , "Set Default", "key", "key"
      
      'Populate the Initially combo
      Set objItem = cboProperties(4).ComboItems.Add(, , "Immediate", "key", "key")
      objItem.Selected = True
      cboProperties(4).ComboItems.Add , , "Deferred", "key", "key"
      
      'Populate the local columns combo
      For Each objLItem In frmCallingForm.lvProperties(0).ListItems
        cboProperties(5).ComboItems.Add , , objLItem.Text, "column", "column"
      Next objLItem
      
    Case "MP"
  
      'Display/Edit the specified ForeignKey.
      Set objForeignKey = ForeignKey
    
      Me.Caption = "ForeignKey: " & objForeignKey.Identifier
      txtProperties(0).Text = objForeignKey.Name
      txtProperties(1).Text = objForeignKey.OID
      Set objItem = cboProperties(0).ComboItems.Add(, , objForeignKey.Table, "table")
      objItem.Selected = True
      Set objItem = cboProperties(1).ComboItems.Add(, , objForeignKey.ReferencedTable, "table")
      objItem.Selected = True
      Set objItem = cboProperties(2).ComboItems.Add(, , objForeignKey.OnDelete, "key")
      objItem.Selected = True
      Set objItem = cboProperties(3).ComboItems.Add(, , objForeignKey.OnUpdate, "key")
      objItem.Selected = True
      Set objItem = cboProperties(4).ComboItems.Add(, , objForeignKey.Initially, "key")
      objItem.Selected = True
      chkProperties(0).Value = Bool2Bin(objForeignKey.Deferrable)
      For Each objRelationship In objForeignKey.Relationships
        Set objLItem = lvProperties(0).ListItems.Add(, , objRelationship.LocalColumn, "relationship", "relationship")
        objLItem.SubItems(1) = objRelationship.ReferencedColumn
      Next objRelationship
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmForeignKey.Initialise"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objForeignKey Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objForeignKey.Deferrable)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmForeignKey.chkProperties_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmForeignKey.cboProperties_Click(" & Index & ")", etFullDebug

Dim objColumn As pgColumn

  If objForeignKey Is Nothing Then
    If Index = 1 Then
      cboProperties(6).ComboItems.Clear
      lvProperties(0).ListItems.Clear
      For Each objColumn In frmMain.svr.Databases(szDatabase).Tables(cboProperties(1).Text).Columns
        If Not objColumn.SystemObject Then cboProperties(6).ComboItems.Add , , objColumn.Name, "column", "column"
      Next objColumn
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmForeignKey.cboProperties_Click"
End Sub
