VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOperatorClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operator Class"
   ClientHeight    =   6876
   ClientLeft      =   3612
   ClientTop       =   1668
   ClientWidth     =   5520
   Icon            =   "frmOperatorClass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6876
   ScaleWidth      =   5520
   Begin MSComctlLib.ImageList il 
      Left            =   45
      Top             =   6300
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperatorClass.frx":058A
            Key             =   "index"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperatorClass.frx":0B24
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperatorClass.frx":11F6
            Key             =   "function"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOperatorClass.frx":18C8
            Key             =   "operator"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   18
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   19
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
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmOperatorClass.frx":1F9A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboProperties(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&Operator"
      TabPicture(1)   =   "frmOperatorClass.frx":1FB6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtProperties(3)"
      Tab(1).Control(1)=   "chkProperties(1)"
      Tab(1).Control(2)=   "cmdRemoveOp"
      Tab(1).Control(3)=   "cmdAddOp"
      Tab(1).Control(4)=   "lvProperties(0)"
      Tab(1).Control(5)=   "cboProperties(2)"
      Tab(1).Control(6)=   "lblProperties(5)"
      Tab(1).Control(7)=   "lblProperties(4)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&Function"
      TabPicture(2)   =   "frmOperatorClass.frx":1FD2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtProperties(4)"
      Tab(2).Control(1)=   "cmdAddFnc"
      Tab(2).Control(2)=   "cmdRemoveFnc"
      Tab(2).Control(3)=   "lvProperties(1)"
      Tab(2).Control(4)=   "cboProperties(3)"
      Tab(2).Control(5)=   "lblProperties(8)"
      Tab(2).Control(6)=   "lblProperties(7)"
      Tab(2).ControlCount=   7
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "The index access method's support procedure number for a function associated with the operator class."
         Top             =   1020
         Width           =   3390
      End
      Begin VB.CommandButton cmdAddFnc 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72360
         TabIndex        =   16
         ToolTipText     =   "Add argument."
         Top             =   5880
         Width           =   1320
      End
      Begin VB.CommandButton cmdRemoveFnc 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -70980
         TabIndex        =   17
         ToolTipText     =   "Remove the selected argument."
         Top             =   5880
         Width           =   1320
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "The index access method's strategy number for an operator associated with the operator class."
         Top             =   1020
         Width           =   3390
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Recheck?"
         Height          =   240
         Index           =   1
         Left            =   -74880
         TabIndex        =   9
         ToolTipText     =   $"frmOperatorClass.frx":1FEE
         Top             =   1380
         Width           =   1995
      End
      Begin VB.CommandButton cmdRemoveOp 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -70980
         TabIndex        =   12
         ToolTipText     =   "Remove the selected argument."
         Top             =   5880
         Width           =   1320
      End
      Begin VB.CommandButton cmdAddOp 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72360
         TabIndex        =   11
         ToolTipText     =   "Add argument."
         Top             =   5880
         Width           =   1320
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4092
         Index           =   0
         Left            =   -74880
         TabIndex        =   10
         Top             =   1680
         Width           =   5196
         _ExtentX        =   9165
         _ExtentY        =   7218
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Strategy Number"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Recheck"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operator"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Default?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   $"frmOperatorClass.frx":20BA
         Top             =   1872
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The Operator Class owner."
         Top             =   1440
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The Operator Class OID (Object ID) in the PostgreSQL Database."
         Top             =   1035
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the operator class."
         Top             =   630
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   300
         Index           =   2
         Left            =   -73080
         TabIndex        =   7
         ToolTipText     =   "The identifier of an operator associated with the operator class."
         Top             =   660
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
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "The column data type that this operator class is for."
         Top             =   2220
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
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "The name of the index access method this operator class is for."
         Top             =   2640
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4068
         Index           =   1
         Left            =   -74880
         TabIndex        =   15
         Top             =   1680
         Width           =   5196
         _ExtentX        =   9165
         _ExtentY        =   7176
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
            Text            =   "Support Number"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Function"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   300
         Index           =   3
         Left            =   -73080
         TabIndex        =   13
         ToolTipText     =   "The name of a function that is an index access method support procedure for the operator class. "
         Top             =   660
         Width           =   3396
         _ExtentX        =   5990
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Support Number"
         Height          =   192
         Index           =   8
         Left            =   -74880
         TabIndex        =   28
         Top             =   1092
         Width           =   1176
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   192
         Index           =   7
         Left            =   -74868
         TabIndex        =   27
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Strategy Number "
         Height          =   192
         Index           =   5
         Left            =   -74880
         TabIndex        =   26
         Top             =   1068
         Width           =   1248
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Access Method"
         Height          =   192
         Index           =   6
         Left            =   135
         TabIndex        =   25
         Top             =   2736
         Width           =   1116
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "For Data Type"
         Height          =   192
         Index           =   3
         Left            =   135
         TabIndex        =   24
         Top             =   2316
         Width           =   1044
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Operator"
         Height          =   192
         Index           =   4
         Left            =   -74868
         TabIndex        =   23
         Top             =   720
         Width           =   636
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   675
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   1080
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   20
         Top             =   1485
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmOperatorClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmOperatorClass.frm - Edit/Create a Function

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objOperatorClass As pgOperatorClass

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cboProperties_Click(" & Index & ")", etFullDebug

Dim objOperator As pgOperator
Dim objFunction As pgFunction
Dim ii As Integer
Dim lii As Long
Dim bFound As Boolean

  If Index = 0 Then
    '///////////////////
    'load operator
    cboProperties(2).ComboItems.Clear
    cboProperties(2).Text = ""
    
    'Add pg_catalog items first, unqualified
    For Each objOperator In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Operators
      If objOperator.LeftOperandType = cboProperties(0).Text And objOperator.RightOperandType = cboProperties(0).Text Then
        cboProperties(2).ComboItems.Add , , fmtID(objOperator.Namespace) & "." & objOperator.Name, "operator"
      End If
    Next
    If cboProperties(2).ComboItems.Count > 0 Then cboProperties(2).ComboItems(1).Selected = True
  
    '///////////////////
    'load function
    cboProperties(3).ComboItems.Clear
    cboProperties(3).Text = ""
    
    'Add pg_catalog items first, unqualified
    For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Functions
      bFound = False
      If objFunction.Arguments.Count = 1 Then
        bFound = (objFunction.Arguments(1) = "internal")
      ElseIf objFunction.Arguments.Count > 1 Then
        For lii = 1 To objFunction.Arguments.Count - 1
          If objFunction.Arguments(lii) = cboProperties(0).Text Then
            bFound = True
            Exit For
          End If
        Next
      End If
      If bFound Then
        cboProperties(3).ComboItems.Add , , fmtID(objFunction.Namespace) & "." & objFunction.Identifier, "function"
      End If
    Next
    If cboProperties(3).ComboItems.Count > 0 Then cboProperties(3).ComboItems(1).Selected = True
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmAggregate.cboProperties_Click"
End Sub

Private Sub cmdAddOp_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cmdAddOp_Click()", etFullDebug

Dim objLItem As ListItem

  If Len(cboProperties(2).Text) = 0 Then
    MsgBox §§TrasLang§§("You must select a operator!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If
  If Len(txtProperties(3).Text) = 0 Then
    MsgBox §§TrasLang§§("You must specify a strategy number!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If

  'verify if exist in list
  For Each objLItem In lvProperties(0).ListItems
    If objLItem.SubItems(2) = cboProperties(2).Text Then
      MsgBox "'" & objLItem.SubItems(2) & §§TrasLang§§("' already appears in the Operator List!"), vbExclamation, §§TrasLang§§("Error")
      Exit Sub
    End If
  Next

  Set objLItem = lvProperties(0).ListItems.Add(, , txtProperties(3).Text, "operator", "operator")
  objLItem.SubItems(1) = BoolToYesNo(Bin2Bool(chkProperties(1).Value))
  objLItem.SubItems(2) = cboProperties(2).Text
  
  AutoSizeColumnLv lvProperties(0)

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.cmdAddOp_Click"
End Sub

Private Sub cmdRemoveOp_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cmdRemoveOp_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then
    MsgBox §§TrasLang§§("You must select a operator to remove!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If

  If MsgBox(§§TrasLang§§("Are you sure you wish to remove operator '") & lvProperties(0).SelectedItem.SubItems(2) & "' ?", vbQuestion + vbYesNo, §§TrasLang§§("Remove Operator")) = vbNo Then Exit Sub
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  AutoSizeColumnLv lvProperties(0)

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.cmdRemoveOp_Click"
End Sub

Private Sub cmdAddFnc_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cmdAddFnc_Click()", etFullDebug

Dim objLItem As ListItem

  If Len(cboProperties(3).Text) = 0 Then
    MsgBox §§TrasLang§§("You must select a function!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If
  If Len(txtProperties(4).Text) = 0 Then
    MsgBox §§TrasLang§§("You must specify a support number!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If

  'verify if exist in list
  For Each objLItem In lvProperties(1).ListItems
    If objLItem.SubItems(1) = cboProperties(2).Text Then
      MsgBox "'" & objLItem.SubItems(2) & §§TrasLang§§("' already appears in the Function List!"), vbExclamation, §§TrasLang§§("Error")
      Exit Sub
    End If
  Next

  Set objLItem = lvProperties(1).ListItems.Add(, , txtProperties(4).Text, "function", "function")
  objLItem.SubItems(1) = cboProperties(3).Text
  
  AutoSizeColumnLv lvProperties(1)

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.cmdAddFnc_Click"
End Sub

Private Sub cmdRemoveFnc_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cmdRemoveFnc_Click()", etFullDebug

  If lvProperties(1).SelectedItem Is Nothing Then
    MsgBox §§TrasLang§§("You must select a function to remove!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If

  If MsgBox(§§TrasLang§§("Are you sure you wish to remove function '") & lvProperties(1).SelectedItem.SubItems(1) & "' ?", vbQuestion + vbYesNo, §§TrasLang§§("Remove Function")) = vbNo Then Exit Sub
  lvProperties(1).ListItems.Remove lvProperties(1).SelectedItem.Index
  AutoSizeColumnLv lvProperties(1)

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.cmdRemoveFnc_Click"
End Sub

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cmdCancel_Click()", etFullDebug

  Unload Me

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.cmdCancel_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional oOperatorClass As pgOperatorClass)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.Initialise(" & QUOTE & szDB & QUOTE & "," & QUOTE & szNS & QUOTE & ")", etFullDebug

Dim rs As Recordset
Dim objType As pgType
Dim objOpClassFnc As OpClassFnc
Dim objOpClassOp As OpClassOp
Dim objLItem As ListItem
Dim ii As Integer

  szDatabase = szDB
  szNamespace = szNS

  PatchForm Me

  If oOperatorClass Is Nothing Then

    'Create a new Operator Class
    bNew = True
    Me.Caption = §§TrasLang§§("Create Operator Class")

    'Load the combo
    'for data type
    For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
      cboProperties(0).ComboItems.Add , , fmtTypeName(objType), "type"
    Next
    cboProperties(0).ComboItems(1).Selected = True
    
    'index access methods
    Set rs = frmMain.svr.Databases(szDatabase).Execute("SELECT amname FROM pg_am ORDER BY amname")
    While Not rs.EOF
      cboProperties(1).ComboItems.Add , , rs!amname, "index"
      rs.MoveNext
    Wend
    cboProperties(1).ComboItems(1).Selected = True
  
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    txtProperties(3).BackColor = &H80000005
    txtProperties(3).Locked = False
    txtProperties(4).BackColor = &H80000005
    txtProperties(4).Locked = False
    
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    cboProperties(2).BackColor = &H80000005
    cboProperties(3).BackColor = &H80000005
    lvProperties(0).BackColor = &H80000005
    lvProperties(1).BackColor = &H80000005
    
    cmdAddOp.Enabled = True
    cmdRemoveOp.Enabled = True
    cmdAddFnc.Enabled = True
    cmdRemoveFnc.Enabled = True
  
    cboProperties_Click 0
  Else

    'Display/Edit the specified Function.
    Set objOperatorClass = oOperatorClass
    bNew = False

    With objOperatorClass
      Me.Caption = §§TrasLang§§("Operator Class: ") & .Identifier
      txtProperties(0).Text = .Name
      txtProperties(1).Text = .Oid
      txtProperties(2).Text = .Owner
      chkProperties(0).Value = IIf(.Default, 1, 0)
      cboProperties(0).ComboItems.Add , , .InputType, "type"
      cboProperties(0).ComboItems(1).Selected = True
      cboProperties(1).ComboItems.Add , , .AccessMethod, "index"
      cboProperties(1).ComboItems(1).Selected = True
    
      'operator
      For Each objOpClassOp In .OpClassOps
        Set objLItem = lvProperties(0).ListItems.Add(, , objOpClassOp.StrategyNumber, "operator", "operator")
        objLItem.SubItems(1) = BoolToYesNo(objOpClassOp.Rechecked)
        objLItem.SubItems(2) = objOpClassOp.Operator
      Next
      AutoSizeColumnLv lvProperties(0)
      
      'function
      For Each objOpClassFnc In .OpClassFncs
        Set objLItem = lvProperties(1).ListItems.Add(, , objOpClassFnc.ProcedureIndex, "function", "function")
        objLItem.SubItems(1) = objOpClassFnc.Procedure
      Next
      AutoSizeColumnLv lvProperties(1)
    End With
    chkProperties(0).Enabled = False
    chkProperties(1).Enabled = False

  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.Initialise"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmOperatorClass.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objNewOperatoClass As pgOperatorClass
Dim DataFncs As New OpClassFncs
Dim DataOps As New OpClassOps
Dim OpClassFnc As OpClassFnc
Dim OpClassOp As OpClassOp
Dim objLItem As ListItem

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox §§TrasLang§§("You must specify a operator class name!"), vbExclamation, §§TrasLang§§("Error")
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox §§TrasLang§§("You must select a data type!"), vbExclamation, §§TrasLang§§("Error")
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox §§TrasLang§§("You must select a access method!"), vbExclamation, §§TrasLang§§("Error")
    tabProperties.Tab = 1
    cboProperties(1).SetFocus
    Exit Sub
  End If
  
  'operator
  For Each objLItem In lvProperties(0).ListItems
    Set OpClassOp = New OpClassOp
    With OpClassOp
      .Operator = objLItem.SubItems(2)
      .Rechecked = YesNoToBool(objLItem.SubItems(1))
      .StrategyNumber = objLItem.Text
    End With
    DataOps.Add OpClassOp
  Next
  
  'function
  For Each objLItem In lvProperties(1).ListItems
    Set OpClassFnc = New OpClassFnc
    With OpClassFnc
      .Procedure = objLItem.SubItems(1)
      .ProcedureIndex = objLItem.Text
    End With
    DataFncs.Add OpClassFnc
  Next
  
  If bNew Then
    StartMsg §§TrasLang§§("Creating Operator Class...")
    Set objNewOperatoClass = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).OperatorsClass.Add(txtProperties(0).Text, cboProperties(1).Text, cboProperties(0).Text, Bin2Bool(chkProperties(0).Value), DataOps, DataFncs)

    'Add a new node and update the text on the parent
    On Error Resume Next
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).OperatorsClass.Tag
    Set objNewOperatoClass.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "OPC-" & GetID, txtProperties(0).Text & " (" & cboProperties(1).Text & ")", "operatorclass")
    objNode.Text = §§TrasLang§§("Operators Class (") & objNode.Children & ")"
    If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
  Else
    StartMsg §§TrasLang§§("Updating Operator Class...")
  End If
  
  'Simulate a node click to refresh the ListFunction
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
  
  EndMsg
  Unload Me
  Exit Sub

Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmOperatorClass.cmdOK_Click"
End Sub
