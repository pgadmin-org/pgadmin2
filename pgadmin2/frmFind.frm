VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Object"
   ClientHeight    =   7524
   ClientLeft      =   2232
   ClientTop       =   1572
   ClientWidth     =   9840
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7524
   ScaleWidth      =   9840
   Begin VB.CheckBox chkOnlyOid 
      Caption         =   "Find only for &Oid"
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      ToolTipText     =   "Check to find only Oid."
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Frame fraByOid 
      Caption         =   "Find only for Oid"
      Enabled         =   0   'False
      Height          =   612
      Left            =   4920
      TabIndex        =   23
      Top             =   3960
      Width           =   2412
      Begin VB.TextBox txtOid 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         ToolTipText     =   "Enter an object Oid."
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Oid"
         Enabled         =   0   'False
         Height          =   192
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   252
      End
   End
   Begin VB.Frame fraCol 
      Caption         =   "Display columns"
      Height          =   1212
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   4695
      Begin MSComctlLib.ListView lvColResult 
         Height          =   852
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "What columns should be included in the results?"
         Top             =   240
         Width           =   4452
         _ExtentX        =   7853
         _ExtentY        =   1503
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraFind 
      Caption         =   "Find options"
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtSql 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Enter an object's DDL, or part of an object's DDL."
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox txtComment 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Enter an object comment or part of an object comment."
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Enter an object name, or part of an object name."
         Top             =   1560
         Width           =   3375
      End
      Begin MSComctlLib.ImageCombo cboDatabase 
         Height          =   300
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Select a database to search."
         Top             =   276
         Width           =   3372
         _ExtentX        =   5948
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboSearchFor 
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "What search type should be used?"
         Top             =   2640
         Width           =   3372
         _ExtentX        =   5948
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ListView lvNameSpace 
         Height          =   855
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Select the schemas to search in."
         Top             =   600
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   1503
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "SQL"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   21
         Top             =   2385
         Width           =   315
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Comment"
         Height          =   192
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Database"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Schema"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   675
         Width           =   585
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1620
         Width           =   420
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Search for"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   735
      End
   End
   Begin VB.CheckBox chkAdvOpt 
      Caption         =   "&Advanced Options"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      ToolTipText     =   "Check to apply advanced search options."
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame fraAdvFind 
      Caption         =   "Advanced options"
      Enabled         =   0   'False
      Height          =   2892
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   4815
      Begin MSComctlLib.ListView lvObjType 
         Height          =   1176
         Left            =   1200
         TabIndex        =   8
         ToolTipText     =   "Select the object types to search for."
         Top             =   240
         Width           =   3516
         _ExtentX        =   6181
         _ExtentY        =   2074
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvOwner 
         Height          =   1176
         Left            =   1200
         TabIndex        =   9
         ToolTipText     =   "Select the object owners whose objects will be searched."
         Top             =   1560
         Width           =   3516
         _ExtentX        =   6181
         _ExtentY        =   2074
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "Object Type"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   270
         Width           =   870
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "By Owner"
         Enabled         =   0   'False
         Height          =   192
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   696
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   9000
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":1CFA
            Key             =   "aggregate"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":23CC
            Key             =   "check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":2A9E
            Key             =   "column"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":3170
            Key             =   "function"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":3842
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":3F14
            Key             =   "index"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":44AE
            Key             =   "indexcolumn"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":4B80
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":5252
            Key             =   "language"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":5924
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":5FF6
            Key             =   "property"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":6590
            Key             =   "relationship"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":66EA
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":6DBC
            Key             =   "server"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":6F16
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":75E8
            Key             =   "table"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":7CBA
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":838C
            Key             =   "type"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":8A5E
            Key             =   "user"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":8BB8
            Key             =   "view"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":928A
            Key             =   "hiproperty"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":9824
            Key             =   "database"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":997E
            Key             =   "closeddatabase"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":9AD8
            Key             =   "baddatabase"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":9C32
            Key             =   "statistics"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":A804
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":AED6
            Key             =   "namespace"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":BAA8
            Key             =   "all"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":BC02
            Key             =   "cast"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":C7D4
            Key             =   "conversion"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFind.frx":D0AE
            Key             =   "operatorclass"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2772
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Displays the results of the search."
      Top             =   4680
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   4890
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "il"
      SmallIcons      =   "il"
      ColHdrIcons     =   "il"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmFind.frm - Find Object Database

Option Explicit

Public Sub Initialise()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.Initialise()", etFullDebug
  
Dim objDatabase As pgDatabase
Dim objUser As pgUser
Dim lvItem As ListItem

  PatchForm Me
  
  'load database
  cboDatabase.ComboItems.Clear
  For Each objDatabase In frmMain.svr.Databases
    If Not (objDatabase.SystemObject And Not ctx.IncludeSys) And objDatabase.AllowConnections Then
      cboDatabase.ComboItems.Add , objDatabase.Name, objDatabase.Name, "database", "database"
    End If
  Next
  cboDatabase.ComboItems(1).Selected = True
  cboDatabase_Click
  
  'search for modal
  cboSearchFor.ComboItems.Clear
  cboSearchFor.ComboItems.Add , "WWR", §§TrasLang§§("Whole Word"), "all"
  cboSearchFor.ComboItems.Add , "BGN", §§TrasLang§§("Beginning"), "all"
  cboSearchFor.ComboItems.Add , "END", §§TrasLang§§("Ending"), "all"
  cboSearchFor.ComboItems.Add , "SBR", §§TrasLang§§("Substring"), "all"
  cboSearchFor.ComboItems(1).Selected = True
  
  'load object type
  lvObjType.ListItems.Clear
  lvObjType.ListItems.Add , "AGG", §§TrasLang§§("Aggregate"), "aggregate", "aggregate"
  If ctx.dbVer >= 7.3 Then
    lvObjType.ListItems.Add , "CST", §§TrasLang§§("Cast"), "cast", "cast"
    lvObjType.ListItems.Add , "DOM", §§TrasLang§§("Domain"), "domain", "domain"
    lvObjType.ListItems.Add , "CNV", §§TrasLang§§("Conversion"), "conversion", "conversion"
  End If
  lvObjType.ListItems.Add , "FNC", §§TrasLang§§("Function"), "function", "function"
  lvObjType.ListItems.Add , "LNG", §§TrasLang§§("Language"), "language", "language"
  lvObjType.ListItems.Add , "OPR", §§TrasLang§§("Operator"), "operator", "operator"
  If ctx.dbVer >= 7.3 Then
    lvObjType.ListItems.Add , "OPC", §§TrasLang§§("Operator Class"), "operatorclass", "operatorclass"
  End If
  lvObjType.ListItems.Add , "SEQ", §§TrasLang§§("Sequence"), "sequence", "sequence"
  lvObjType.ListItems.Add , "TBL", §§TrasLang§§("Table"), "table", "table"
  lvObjType.ListItems.Add , "TYP", §§TrasLang§§("Type"), "type", "type"
  lvObjType.ListItems.Add , "VIE", §§TrasLang§§("View"), "view", "view"
  
  lvOwner.ListItems.Clear
  For Each objUser In frmMain.svr.Users
    lvOwner.ListItems.Add , , objUser.Name, "user", "user"
  Next
  
  'column result
  lvColResult.ListItems.Clear
  Set lvItem = lvColResult.ListItems.Add(, "NAM", §§TrasLang§§("Name"), "column", "column")
  lvItem.Checked = True
  Set lvItem = lvColResult.ListItems.Add(, "SCH", §§TrasLang§§("Schema"), "column", "column")
  lvItem.Checked = True
  lvColResult.ListItems.Add , "OID", §§TrasLang§§("Oid Object"), "column", "column"
  lvColResult.ListItems.Add , "COM", §§TrasLang§§("Comment"), "column", "column"
  lvColResult.ListItems.Add , "SQL", §§TrasLang§§("SQL"), "column", "column"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.Initialise"
End Sub

Private Sub cboDatabase_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.cboSchema_Click()", etFullDebug
  
Dim objNamespace As pgNamespace
  
  'load schema
  lvNameSpace.ListItems.Clear
  For Each objNamespace In frmMain.svr.Databases(cboDatabase.SelectedItem.Text).Namespaces
    lvNameSpace.ListItems.Add , , objNamespace.Name, "namespace", "namespace"
  Next

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.cboDatabase_Click"
End Sub

Private Sub chkAdvOpt_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.chkAdvOpt_Click()", etFullDebug
  
  If chkAdvOpt.Value = 0 Then
    fraAdvFind.Enabled = False
    lblFind(4).Enabled = False
    lblFind(5).Enabled = False
    lvObjType.Enabled = False
    lvOwner.Enabled = False
  ElseIf chkAdvOpt.Value = 1 Then
    fraAdvFind.Enabled = True
    lblFind(4).Enabled = True
    lblFind(5).Enabled = True
    lvObjType.Enabled = True
    lvOwner.Enabled = True
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.chkAdvOpt_Click"
End Sub

Private Sub chkOnlyOid_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.chkOnlyOid_Click()", etFullDebug
  
  If chkOnlyOid.Value = 0 Then
    lblFind(8).Enabled = False
    fraByOid.Enabled = False
    chkAdvOpt.Enabled = True
    lblFind(3).Enabled = True
    lblFind(6).Enabled = True
    lblFind(7).Enabled = True
    lblFind(0).Enabled = True
    lblFind(2).Enabled = True
    txtName.Enabled = True
    txtComment.Enabled = True
    txtSql.Enabled = True
    cboSearchFor.Enabled = True
    lvNameSpace.Enabled = True
    txtOid.Enabled = False
  ElseIf chkOnlyOid.Value = 1 Then
    lblFind(8).Enabled = True
    fraByOid.Enabled = True
    chkAdvOpt.Value = 0
    chkAdvOpt.Enabled = False
    lblFind(3).Enabled = False
    lblFind(6).Enabled = False
    lblFind(7).Enabled = False
    lblFind(0).Enabled = False
    lblFind(2).Enabled = False
    txtName.Enabled = False
    txtComment.Enabled = False
    txtSql.Enabled = False
    cboSearchFor.Enabled = False
    lvNameSpace.Enabled = False
    txtOid.Enabled = True
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.chkOnlyOid_Click"
End Sub

Private Sub cmdFind_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.cmdFind_Click()", etFullDebug

Dim szName As String
Dim szComment As String
Dim szSQL As String
Dim iLenName As Integer
Dim iLenComment As Integer
Dim iLenSql As Integer
Dim objDatabase As pgDatabase
Dim szSearchFor As String
Dim lvItem As ListItem
Dim colObj As Collection
Dim objCol As Variant
Dim objTmp As Variant
Dim bFound As Boolean
Dim bFoundName As Boolean
Dim bFoundComment As Boolean
Dim bFoundSql As Boolean
Dim szNamespace As String
Dim szOwner() As String
Dim bSreachOwner As Boolean
Dim lOid As Long
  
  StartMsg §§TrasLang§§("Find in progress...")
  
  'find object
  szName = txtName.Text:  iLenName = Len(szName)
  szComment = txtComment.Text:   iLenComment = Len(szComment)
  szSQL = txtSql.Text:  iLenSql = Len(szSQL)
  
  Set objDatabase = frmMain.svr.Databases(cboDatabase.SelectedItem.Text)
  szSearchFor = cboSearchFor.SelectedItem.Key
  
  'columns result
  lvResult.ListItems.Clear
  lvResult.ColumnHeaders.Clear
  lvResult.ColumnHeaders.Add , , §§TrasLang§§("Name")
  lvResult.ColumnHeaders.Add , , §§TrasLang§§("Schema")
  If lvColResult.ListItems("COM").Checked Then lvResult.ColumnHeaders.Add , , §§TrasLang§§("Comment")
  If lvColResult.ListItems("SQL").Checked Then lvResult.ColumnHeaders.Add , , §§TrasLang§§("SQL")
  If lvColResult.ListItems("OID").Checked Then lvResult.ColumnHeaders.Add , , §§TrasLang§§("Oid Object")
  
  'find only oid
  If chkOnlyOid.Value = 1 Then
    If Not IsNumeric(txtOid.Text) Then
      MsgBox §§TrasLang§§("Oid is not numeric!"), vbExclamation, §§TrasLang§§("Error")
      txtOid.SetFocus
      EndMsg
      Exit Sub
    End If
      
    lOid = CLng(txtOid.Text)
    bFound = False
      
    For Each lvItem In lvNameSpace.ListItems
      'Load object for find
      Set colObj = New Collection
      szNamespace = lvItem.Text
      
      colObj.Add objDatabase.Namespaces(szNamespace).Aggregates
      If ctx.dbVer >= 7.3 Then
        colObj.Add objDatabase.Casts
        colObj.Add objDatabase.Namespaces(szNamespace).Domains
        colObj.Add objDatabase.Namespaces(szNamespace).Conversions
        colObj.Add objDatabase.Namespaces(szNamespace).OperatorsClass
      End If
      colObj.Add objDatabase.Namespaces(szNamespace).Functions
      colObj.Add objDatabase.Languages
      colObj.Add objDatabase.Namespaces
      colObj.Add objDatabase.Namespaces(szNamespace).Operators
      colObj.Add objDatabase.Namespaces(szNamespace).Sequences
      colObj.Add objDatabase.Namespaces(szNamespace).Tables
      colObj.Add objDatabase.Namespaces(szNamespace).Types
      colObj.Add objDatabase.Namespaces(szNamespace).Views
        
      For Each objCol In colObj
        For Each objTmp In objCol
          If objTmp.Oid = lOid Then
            bFound = True
            Exit For
          End If
        Next
        If bFound Then Exit For
      Next
      If bFound Then Exit For
    Next
      
    'object is found
    If bFound Then AddResult objTmp, szNamespace
    EndMsg
    Exit Sub
  End If
  
  'owner search
  If chkAdvOpt.Value = 1 Then
    ReDim szOwner(0) As String
    For Each lvItem In lvOwner.ListItems
      If lvItem.Checked Then
        ReDim Preserve szOwner(UBound(szOwner) + 1) As String
        szOwner(UBound(szOwner)) = lvItem.Text
      End If
    Next
    bSreachOwner = True
    If UBound(szOwner) = 0 Then bSreachOwner = False
  End If
  
  'verify if select
  bFound = False
  For Each lvItem In lvNameSpace.ListItems
    If lvItem.Checked Then
      bFound = True
      Exit For
    End If
  Next
  If Not bFound Then
    MsgBox §§TrasLang§§("Schema not selected!"), vbExclamation, §§TrasLang§§("Error")
    EndMsg
    Exit Sub
  End If
  
  For Each lvItem In lvNameSpace.ListItems
    If lvItem.Checked = True Then
      'Load object for find
      Set colObj = New Collection
      szNamespace = lvItem.Text
      If chkAdvOpt.Value = 1 Then
        If lvObjType.ListItems("AGG").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Aggregates
        If ctx.dbVer >= 7.3 Then
          If lvObjType.ListItems("CST").Checked Then colObj.Add objDatabase.Casts
          If lvObjType.ListItems("DOM").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Domains
          If lvObjType.ListItems("CNV").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Conversions
          If lvObjType.ListItems("OPC").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).OperatorsClass
        End If
        If lvObjType.ListItems("FNC").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Functions
        If lvObjType.ListItems("LNG").Checked Then colObj.Add objDatabase.Languages
        If lvObjType.ListItems("OPR").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Operators
        If lvObjType.ListItems("SEQ").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Sequences
        If lvObjType.ListItems("TBL").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Tables
        If lvObjType.ListItems("TYP").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Types
        If lvObjType.ListItems("VIE").Checked Then colObj.Add objDatabase.Namespaces(szNamespace).Views
      Else
        colObj.Add objDatabase.Namespaces(szNamespace).Aggregates
        If ctx.dbVer >= 7.3 Then
          colObj.Add objDatabase.Casts
          colObj.Add objDatabase.Namespaces(szNamespace).Domains
          colObj.Add objDatabase.Namespaces(szNamespace).Conversions
          colObj.Add objDatabase.Namespaces(szNamespace).OperatorsClass
        End If
        colObj.Add objDatabase.Namespaces(szNamespace).Functions
        colObj.Add objDatabase.Languages
        colObj.Add objDatabase.Namespaces(szNamespace).Operators
        colObj.Add objDatabase.Namespaces(szNamespace).Sequences
        colObj.Add objDatabase.Namespaces(szNamespace).Tables
        colObj.Add objDatabase.Namespaces(szNamespace).Types
        colObj.Add objDatabase.Namespaces(szNamespace).Views
      End If
    
      'loop collection
      For Each objCol In colObj
        For Each objTmp In objCol
          bFoundName = True
          bFoundComment = True
          bFoundSql = True
          
          'find by name
          If iLenName > 0 Then
            bFoundName = False
            Select Case szSearchFor
              Case "WWR"
                bFoundName = objTmp.Name = szName
              Case "BGN"
                bFoundName = Left(objTmp.Name, iLenName) = szName
              Case "END"
                bFoundName = Right(objTmp.Name, iLenName) = szName
              Case "SBR"
                bFoundName = InStr(objTmp.Name, szName) > 0
            End Select
          End If
          
          'find by comment
          If iLenComment > 0 Then
            bFoundComment = True
            If ObjcetHaveComment(objTmp.ObjectType) Then
              Select Case szSearchFor
                Case "WWR"
                  bFoundComment = objTmp.Comment = szComment
                Case "BGN"
                  bFoundComment = Left(objTmp.Comment, iLenComment) = szComment
                Case "END"
                  bFoundComment = Right(objTmp.Comment, iLenComment) = szComment
                Case "SBR"
                  bFoundComment = InStr(objTmp.Comment, szComment) > 0
              End Select
            End If
          End If
          
          'find by sql
          If iLenSql > 0 Then
            bFoundSql = False
            Select Case szSearchFor
              Case "WWR"
                bFoundSql = objTmp.SQL = szSQL
              Case "BGN"
                bFoundSql = Left(objTmp.SQL, iLenSql) = szSQL
              Case "END"
                bFoundSql = Right(objTmp.SQL, iLenSql) = szSQL
              Case "SBR"
                bFoundSql = InStr(objTmp.SQL, szSQL) > 0
            End Select
          End If
       
          bFound = bFoundName And bFoundComment And bFoundSql
        
          'advanced options
          If fraFind.Visible Then
            'search by owner
            If bFound And bSreachOwner Then
              If objTmp.ObjectType <> "Cast" And objTmp.ObjectType <> "Language" Then
                If UBound(Filter(szOwner, objTmp.Owner)) > -1 Then
                  bFound = True
                Else
                  bFound = False
                End If
              End If
            End If
          End If
          
          'object is found
          If bFound Then AddResult objTmp, szNamespace
        Next
      Next
    End If
  Next
  AutoSizeColumnLv lvResult
  
  EndMsg
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.cmdFind_Click"
End Sub

Private Sub lvResult_DblClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.lv_DblClick()", etFullDebug

Dim objDatabase As pgDatabase
Dim objNamespace As pgNamespace
Dim szName As String
  
  Set objDatabase = frmMain.svr.Databases(cboDatabase.Text)
  'if object don't have a schema
  If Len(lvResult.SelectedItem.SubItems(1)) > 0 Then
    Set objNamespace = objDatabase.Namespaces(lvResult.SelectedItem.SubItems(1))
  End If
  szName = lvResult.SelectedItem.Text
  
  Select Case Left(lvResult.SelectedItem.Key, 3)
    Case "NSP"
      Dim objNamespaceForm As New frmNamespace
      Load objNamespaceForm
      objNamespaceForm.Initialise objDatabase.Name, objNamespace
      objNamespaceForm.Show
    
    Case "AGG"
      Dim objAggregateForm As New frmAggregate
      Load objAggregateForm
      objAggregateForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Aggregates(szName)
      objAggregateForm.Show
      
    Case "CST"
      Dim objCastForm As New frmCast
      Load objCastForm
      objCastForm.Initialise objDatabase.Name, objDatabase.Casts(szName)
      objCastForm.Show
    
    Case "DOM"
      Dim objDomainForm As New frmDomain
      Load objDomainForm
      objDomainForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Domains(szName)
      objDomainForm.Show
      
    Case "CNV"
      Dim objConversionForm As New frmConversion
      Load objConversionForm
      objConversionForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Conversions(szName)
      objConversionForm.Show
      
    Case "FNC"
      Dim objFunctionForm As New frmFunction
      Load objFunctionForm
      objFunctionForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Functions(szName)
      objFunctionForm.Show
      
    Case "LNG"
      Dim objLanguageForm As New frmLanguage
      Load objLanguageForm
      objLanguageForm.Initialise objDatabase.Name, objDatabase.Languages(szName)
      objLanguageForm.Show
    
    Case "OPR"
      Dim objOperatorForm As New frmOperator
      Load objOperatorForm
      objOperatorForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Operators(szName)
      objOperatorForm.Show
    
    Case "OPC"
      Dim objOperatorClassForm As New frmOperatorClass
      Load objOperatorClassForm
      objOperatorClassForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.OperatorsClass(szName)
      objOperatorClassForm.Show
    
    Case "SEQ"
      Dim objSequenceForm As New frmSequence
      Load objSequenceForm
      objSequenceForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Sequences(szName)
      objSequenceForm.Show
    
    Case "TBL"
      Dim objTableForm As New frmTable
      Load objTableForm
      objTableForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Tables(szName)
      objTableForm.Show
    
    Case "TYP"
      Dim objTypeForm As New frmType
      Load objTypeForm
      objTypeForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Types(szName)
      objTypeForm.Show
    
    Case "VIE"
      Dim objViewForm As New frmView
      Load objViewForm
      objViewForm.Initialise objDatabase.Name, objNamespace.Name, objNamespace.Views(szName)
      objViewForm.Show
      
  End Select
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.lv_DblClick"
End Sub

Private Sub lvColResult_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.lvColResult_ItemCheck(" & Item.Text & ")", etFullDebug

  If Item.Key = "NAM" Then
    Item.Checked = True
    Exit Sub
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.lvColResult_ItemCheck"
End Sub

'return if object have comment
Private Function ObjcetHaveComment(ByVal ObjectType As String) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.ObjcetHaveComment(" & ObjectType & ")", etFullDebug

  Select Case ObjectType
    Case "Cast", "Language", "Conversion", "OperatorClass"
      ObjcetHaveComment = False
    Case Else
      ObjcetHaveComment = True
  End Select
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.ObjcetHaveComment"
End Function

'add element in result
Private Sub AddResult(ByVal ObjResult As Object, szNamespace As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFind.AddResult(" & ObjResult.ObjectType & ",'" & szNamespace & "')", etFullDebug

Dim iCol As Integer
Dim szImg As String
Dim szKey As String
Dim lvItem As ListItem
  
  szImg = NameImageByObjectType(ObjResult.ObjectType)

  Select Case ObjResult.ObjectType
    Case "Aggregate"
      szKey = "AGG"
    Case "Cast"
      szKey = "CST"
    Case "Domain"
      szKey = "DOM"
    Case "Conversion"
      szKey = "CNV"
    Case "Function"
      szKey = "FNC"
    Case "Language"
      szKey = "LNG"
    Case "Operator"
      szKey = "OPR"
    Case "OperatorClass"
      szKey = "OPC"
    Case "Sequence"
      szKey = "SEQ"
    Case "Table"
      szKey = "TBL"
    Case "Type"
      szKey = "TYP"
    Case "View"
      szKey = "VIE"
    Case "Schema"
      szKey = "NSP"
  End Select
  Set lvItem = lvResult.ListItems.Add(, szKey & "_" & GetID, ObjResult.Identifier)
  If Len(szImg) > 0 Then
    lvItem.Icon = szImg
    lvItem.SmallIcon = szImg
  End If
  If ObjResult.ObjectType <> "Cast" And ObjResult.ObjectType <> "Language" Then lvItem.SubItems(1) = szNamespace
  iCol = 1
  
  If lvColResult.ListItems("COM").Checked Then
    iCol = iCol + 1
     If ObjcetHaveComment(ObjResult.ObjectType) Then lvItem.SubItems(iCol) = ObjResult.Comment
  End If

  If lvColResult.ListItems("SQL").Checked Then
    iCol = iCol + 1
    lvItem.SubItems(iCol) = ObjResult.SQL
  End If

  If lvColResult.ListItems("OID").Checked Then
    iCol = iCol + 1
    lvItem.SubItems(iCol) = ObjResult.Oid
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFind.AddResult"
End Sub
