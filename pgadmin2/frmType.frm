VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Type"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmType.frx":0000
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
      TabIndex        =   8
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   9
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
      TabCaption(0)   =   "&Properties 1"
      TabPicture(0)   =   "frmType.frx":06C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "hbxProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProperties(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "P&roperties 2"
      TabPicture(1)   =   "frmType.frx":06DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblProperties(11)"
      Tab(1).Control(1)=   "lblProperties(10)"
      Tab(1).Control(2)=   "lblProperties(9)"
      Tab(1).Control(3)=   "lblProperties(6)"
      Tab(1).Control(4)=   "lblProperties(7)"
      Tab(1).Control(5)=   "cboProperties(4)"
      Tab(1).Control(6)=   "cboProperties(3)"
      Tab(1).Control(7)=   "cboProperties(2)"
      Tab(1).Control(8)=   "chkProperties(0)"
      Tab(1).Control(9)=   "txtProperties(4)"
      Tab(1).Control(10)=   "txtProperties(5)"
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -73065
         Locked          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "The delimiter character for the array elements."
         Top             =   1485
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   -73065
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "The default value for the data type. Usually this is omitted, so that the default is NULL."
         Top             =   675
         Width           =   3390
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Passed by value?"
         Height          =   195
         Index           =   0
         Left            =   -74910
         TabIndex        =   13
         ToolTipText     =   $"frmType.frx":06FA
         Top             =   1935
         Width           =   2040
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "A literal value, which specifies the internal length of the new type (0 = Variable)."
         Top             =   2700
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the type."
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
         ToolTipText     =   "The types OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The types owner."
         Top             =   1485
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "The name of a function, created by CREATE FUNCTION, which converts data from its external form to the type's internal form."
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
         Height          =   2985
         Index           =   0
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "Comments about the operator."
         Top             =   3105
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5265
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
         ToolTipText     =   "The name of a function, created by CREATE FUNCTION, which converts data from its internal form to a form suitable for display."
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
         Left            =   -73065
         TabIndex        =   11
         ToolTipText     =   "The type being created is an array; this specifies the type of the array elements."
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
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   3
         Left            =   -73065
         TabIndex        =   14
         ToolTipText     =   "Storage alignment requirement of the data type."
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
         Index           =   4
         Left            =   -73065
         TabIndex        =   15
         ToolTipText     =   "Storage technique for the data type. If specified, must be 'plain', 'external', 'extended', or 'main'; the default is 'plain'."
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
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Alignment"
         Height          =   195
         Index           =   7
         Left            =   -74865
         TabIndex        =   26
         Top             =   2340
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Storage"
         Height          =   195
         Index           =   6
         Left            =   -74865
         TabIndex        =   25
         Top             =   2745
         Width           =   555
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Default"
         Height          =   195
         Index           =   9
         Left            =   -74865
         TabIndex        =   24
         Top             =   675
         Width           =   510
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Element type"
         Height          =   195
         Index           =   10
         Left            =   -74865
         TabIndex        =   23
         Top             =   1125
         Width           =   915
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Delimiter"
         Height          =   195
         Index           =   11
         Left            =   -74865
         TabIndex        =   22
         Top             =   1530
         Width           =   600
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Internal length"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   21
         Top             =   2745
         Width           =   1005
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Output Function"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   18
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Input function"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Top             =   1935
         Width           =   975
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
            Picture         =   "frmType.frx":07DE
            Key             =   "function"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmType.frx":0D78
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmType.frx":1312
            Key             =   "storage"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmType.frm - Edit/Create a Type

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objType As pgType

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmType.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmType.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmType.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewType As pgType
Dim szInputfunction As String
Dim szOutputFunction As String

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Type name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select an input function!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select an output function!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Type..."
    If Not (cboProperties(0).SelectedItem Is Nothing) Then szInputfunction = cboProperties(0).SelectedItem.Text
    If Not (cboProperties(1).SelectedItem Is Nothing) Then szOutputFunction = cboProperties(1).SelectedItem.Text
    Set objNewType = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types.Add(txtProperties(0).Text, szInputfunction, szOutputFunction, Val(txtProperties(3).Text), txtProperties(4).Text, cboProperties(2).Text, txtProperties(5).Text, Bin2Bool(chkProperties(0).Value), cboProperties(3).Text, cboProperties(4).Text, hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types.Tag
    Set objNewType.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "TYP-" & GetID, txtProperties(0).Text, "type")
    objNode.Text = "Types (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Type..."
    If hbxProperties(0).Tag = "Y" Then objType.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListType
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmType.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional oType As pgType)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmType.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objFunction As pgFunction
Dim objTempType As pgType
Dim objNamespace As pgNamespace
Dim objItem As ComboItem
Dim vArgument As Variant
  
  szDatabase = szDB
  szNamespace = szNS
  
  'Set the font
  For X = 0 To 5
    Set txtProperties(X).Font = ctx.Font
  Next X
  For X = 0 To 4
    Set cboProperties(X).Font = ctx.Font
  Next X
  Set hbxProperties(0).Font = ctx.Font
  
  If oType Is Nothing Then
  
    'Create a new Type
    bNew = True
    Me.Caption = "Create Type"
    
    'Load the combos
    If ctx.dbVer >= 7.3 Then
      'First add pg_catalog items, unqualified
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Functions
        'Input functions can be either xxx(opaque) or xx(opaque, oid, int4)
        If objFunction.Arguments.Count = 1 Then
          If objFunction.Arguments(1) = "opaque" Then cboProperties(0).ComboItems.Add , , fmtID(objFunction.Name), "function"
        ElseIf objFunction.Arguments.Count = 3 Then
          If objFunction.Arguments(1) = "opaque" And _
             objFunction.Arguments(2) = "oid" And _
             objFunction.Arguments(3) = "int4" Then cboProperties(0).ComboItems.Add , , fmtID(objFunction.Name), "function"
        End If
        'Output functions can be either xxx(opaque) or xx(opaque, oid)
        If objFunction.Arguments.Count = 1 Then
          If objFunction.Arguments(1) = "opaque" Then cboProperties(1).ComboItems.Add , , fmtID(objFunction.Name), "function"
        ElseIf objFunction.Arguments.Count = 2 Then
          If objFunction.Arguments(1) = "opaque" And _
             objFunction.Arguments(2) = "oid" Then cboProperties(1).ComboItems.Add , , fmtID(objFunction.Name), "function"
        End If
      Next objFunction
      For Each objTempType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
        If ((objTempType.InternalLength <> -1) Or (objTempType.Element = "")) Then cboProperties(2).ComboItems.Add , , fmtID(objTempType.Name), "type"
      Next objTempType
      'Now load the rest
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objFunction In objNamespace.Functions
            'Input functions can be either xxx(opaque) or xx(opaque, oid, int4)
            If objFunction.Arguments.Count = 1 Then
              If objFunction.Arguments(1) = "opaque" Then cboProperties(0).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
            ElseIf objFunction.Arguments.Count = 3 Then
              If objFunction.Arguments(1) = "opaque" And _
                 objFunction.Arguments(2) = "oid" And _
                 objFunction.Arguments(3) = "int4" Then cboProperties(0).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
            End If
            'Output functions can be either xxx(opaque) or xx(opaque, oid)
            If objFunction.Arguments.Count = 1 Then
              If objFunction.Arguments(1) = "opaque" Then cboProperties(1).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
            ElseIf objFunction.Arguments.Count = 2 Then
              If objFunction.Arguments(1) = "opaque" And _
                 objFunction.Arguments(2) = "oid" Then cboProperties(1).ComboItems.Add , , objNamespace.FormattedID & "." & fmtID(objFunction.Name), "function"
            End If
          Next objFunction
          For Each objTempType In objNamespace.Types
            If ((objTempType.InternalLength <> -1) Or (objTempType.Element = "")) Then cboProperties(2).ComboItems.Add , , objTempType.FormattedID, "type"
          Next objTempType
        End If
      Next objNamespace
    Else
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions
        'Input functions can be either xxx(opaque) or xx(opaque, oid, int4)
        If objFunction.Arguments.Count = 1 Then
          If objFunction.Arguments(1) = "opaque" Then cboProperties(0).ComboItems.Add , , fmtID(objFunction.Name), "function"
        ElseIf objFunction.Arguments.Count = 3 Then
          If objFunction.Arguments(1) = "opaque" And _
             objFunction.Arguments(2) = "oid" And _
             objFunction.Arguments(3) = "int4" Then cboProperties(0).ComboItems.Add , , fmtID(objFunction.Name), "function"
        End If
        'Output functions can be either xxx(opaque) or xx(opaque, oid)
        If objFunction.Arguments.Count = 1 Then
          If objFunction.Arguments(1) = "opaque" Then cboProperties(1).ComboItems.Add , , fmtID(objFunction.Name), "function"
        ElseIf objFunction.Arguments.Count = 2 Then
          If objFunction.Arguments(1) = "opaque" And _
             objFunction.Arguments(2) = "oid" Then cboProperties(1).ComboItems.Add , , fmtID(objFunction.Name), "function"
        End If
      Next objFunction
      For Each objTempType In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types
        If ((objTempType.InternalLength <> -1) Or (objTempType.Element = "")) Then cboProperties(2).ComboItems.Add , , objTempType.FormattedID, "type"
      Next objTempType
    End If
    
    cboProperties(3).ComboItems.Add , , "char", "type"
    cboProperties(3).ComboItems.Add , , "double", "type"
    cboProperties(3).ComboItems.Add , , "int2", "type"
    cboProperties(3).ComboItems.Add , , "int4", "type"
    cboProperties(4).ComboItems.Add , , "PLAIN", "storage"
    cboProperties(4).ComboItems.Add , , "EXTERNAL", "storage"
    cboProperties(4).ComboItems.Add , , "EXTENDED", "storage"
    cboProperties(4).ComboItems.Add , , "MAIN", "storage"
  
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    For X = 3 To 5
      txtProperties(X).BackColor = &H80000005
      txtProperties(X).Locked = False
    Next X
    For X = 0 To 4
      cboProperties(X).BackColor = &H80000005
    Next X
    
  Else
  
    'Display/Edit the specified Type.
    Set objType = oType
    bNew = False

    Me.Caption = "Type: " & objType.Identifier
    txtProperties(0).Text = objType.Name
    txtProperties(1).Text = objType.OID
    txtProperties(2).Text = objType.Owner
    txtProperties(3).Text = objType.InternalLength
    txtProperties(4).Text = objType.Default
    txtProperties(5).Text = objType.Delimiter
    Set objItem = cboProperties(0).ComboItems.Add(, , objType.InputFunction, "function")
    objItem.Selected = True
    Set objItem = cboProperties(1).ComboItems.Add(, , objType.OutputFunction, "type")
    objItem.Selected = True
    Set objItem = cboProperties(2).ComboItems.Add(, , objType.Element, "type")
    objItem.Selected = True
    Set objItem = cboProperties(5).ComboItems.Add(, , objType.Alignment, "type")
    objItem.Selected = True
    Set objItem = cboProperties(6).ComboItems.Add(, , objType.Storage, "storage")
    objItem.Selected = True
    chkProperties(0).Value = Bool2Bin(objType.PassedByValue)
    hbxProperties(0).Text = objType.Comment
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmType.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmType.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmType.hbxProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmType.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objType Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objType.PassedByValue)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmType.chkProperties_Click"
End Sub
