VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmFunction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmFunction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":058A
            Key             =   "language"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":0B24
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":10BE
            Key             =   "opaque"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":1218
            Key             =   "table"
         EndProperty
      EndProperty
   End
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
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmFunction.frx":1372
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "hbxProperties(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtProperties(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkProperties(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "&Input/Output"
      TabPicture(1)   =   "frmFunction.frx":138E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRemove"
      Tab(1).Control(1)=   "cmdAdd"
      Tab(1).Control(2)=   "lvProperties(0)"
      Tab(1).Control(3)=   "cboProperties(1)"
      Tab(1).Control(4)=   "cboProperties(2)"
      Tab(1).Control(5)=   "lblProperties(5)"
      Tab(1).Control(6)=   "lblProperties(4)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Definition"
      TabPicture(2)   =   "frmFunction.frx":13AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "hbxProperties(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74685
         TabIndex        =   13
         ToolTipText     =   "Remove the selected argument."
         Top             =   1980
         Width           =   1320
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74685
         TabIndex        =   12
         ToolTipText     =   "Add argument."
         Top             =   1575
         Width           =   1320
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4515
         Index           =   0
         Left            =   -73065
         TabIndex        =   14
         ToolTipText     =   $"frmFunction.frx":13C6
         Top             =   1530
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Included Arguments"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Strict?"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   $"frmFunction.frx":148F
         Top             =   2655
         Width           =   1995
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Cachable?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   $"frmFunction.frx":1633
         Top             =   2295
         Width           =   1995
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "May be 'sql', 'C', 'internal', or 'plname', where 'plname' is the name of a created procedural language."
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
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The functions owner."
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
         ToolTipText     =   "The functions OID (Object ID) in the PostgreSQL Database."
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
         ToolTipText     =   "The name of the function."
         Top             =   630
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   2985
         Index           =   0
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "Comments about the function."
         Top             =   3015
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
      Begin HighlightBox.HBX hbxProperties 
         Height          =   5730
         Index           =   1
         Left            =   -74865
         TabIndex        =   15
         ToolTipText     =   $"frmFunction.frx":176E
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   10107
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
         Caption         =   "Function Definition/Object Library"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   1
         Left            =   -73065
         TabIndex        =   10
         ToolTipText     =   $"frmFunction.frx":1827
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
         Index           =   2
         Left            =   -73065
         TabIndex        =   11
         ToolTipText     =   "Select an agument data type to add to the argument list."
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
         Caption         =   "Arguments"
         Height          =   195
         Index           =   5
         Left            =   -74865
         TabIndex        =   21
         Top             =   1170
         Width           =   750
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Return Type"
         Height          =   195
         Index           =   4
         Left            =   -74865
         TabIndex        =   20
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Language"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   1890
         Width           =   720
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   675
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   1080
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   16
         Top             =   1485
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmFunction.frm - Edit/Create a Function

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim objFunction As pgFunction

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.cmdRemove_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then Exit Sub
  lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.cmdAdd_Click()", etFullDebug

  If cboProperties(2).Text = "" Then Exit Sub
  Select Case cboProperties(2).SelectedItem.Image
    Case "type"
      lvProperties(0).ListItems.Add , , cboProperties(2).Text, "type", "type"
    Case "opaque"
      lvProperties(0).ListItems.Add , , cboProperties(2).Text, "opaque", "opaque"
    Case "table"
      lvProperties(0).ListItems.Add , , cboProperties(2).Text, "table", "table"
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdAdd_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim szArguments As String
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a function name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select a function language!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select a return type!", vbExclamation, "Error"
    tabProperties.Tab = 1
    cboProperties(1).SetFocus
    Exit Sub
  End If
  If hbxProperties(1).Text = "" Then
    MsgBox "You must specify the function definition or object library!", vbExclamation, "Error"
    tabProperties.Tab = 2
    hbxProperties(1).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Function..."
    For Each objItem In lvProperties(0).ListItems
      szArguments = szArguments & objItem.Text & ", "
    Next objItem
    If Len(szArguments) > 2 Then szArguments = Left(szArguments, Len(szArguments) - 2)
    frmMain.svr.Databases(szDatabase).Functions.Add txtProperties(0).Text, szArguments, cboProperties(1).Text, hbxProperties(1).Text, cboProperties(0).Text, Bin2Bool(chkProperties(0).Value), Bin2Bool(chkProperties(1).Value), hbxProperties(0).Text
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) <> "SVR-" Then
        If (Left(objNode.Key, 4) = "FNC+") And (objNode.Parent.Text = szDatabase) Then
          frmMain.tv.Nodes.Add objNode.Key, tvwChild, "FNC-" & GetID, txtProperties(0).Text & "(" & szArguments & ")", "function"
          objNode.Text = "Functions (" & objNode.Children & ")"
          Exit For
        End If
      End If
    Next objNode
    
  Else
    StartMsg "Updating Function..."
    If hbxProperties(0).Tag = "Y" Then objFunction.Comment = hbxProperties(0).Text
    If hbxProperties(1).Tag = "Y" Then objFunction.Source = hbxProperties(1).Text
  End If
  
  'Simulate a node click to refresh the ListFunction
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, Optional oFunction As pgFunction)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objLanguage As pgLanguage
Dim objType As pgType
Dim objTable As pgTable
Dim objItem As ComboItem
Dim vArgument As Variant
  
  szDatabase = szDB
  hbxProperties(1).Wordlist = ctx.AutoHighlight
  
  If oFunction Is Nothing Then
  
    'Create a new Function
    bNew = True
    Me.Caption = "Create Function"
    
    'Load the combo
    For Each objLanguage In frmMain.svr.Databases(szDatabase).Languages
      cboProperties(0).ComboItems.Add , , objLanguage.Identifier, "language"
    Next objLanguage
    cboProperties(1).ComboItems.Add , , "opaque", "opaque"
    cboProperties(2).ComboItems.Add , , "opaque", "opaque"
    For Each objType In frmMain.svr.Databases(szDatabase).Types
      If Left(objType.Identifier, 1) <> "_" Then cboProperties(1).ComboItems.Add , , objType.Identifier, "type"
      If Left(objType.Identifier, 1) <> "_" Then cboProperties(2).ComboItems.Add , , objType.Identifier, "type"
    Next objType
    For Each objTable In frmMain.svr.Databases(szDatabase).Tables
      cboProperties(1).ComboItems.Add , , objTable.Identifier, "table"
      cboProperties(2).ComboItems.Add , , objTable.Identifier, "table"
    Next objTable
  
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    cboProperties(2).BackColor = &H80000005
    lvProperties(0).BackColor = &H80000005
    hbxProperties(1).BackColor = &H80000005
    hbxProperties(1).Locked = False
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    
  Else
  
    'Display/Edit the specified Function.
    Set objFunction = oFunction
    bNew = False
    
    Me.Caption = "Function: " & objFunction.Identifier
    txtProperties(0).Text = objFunction.Name
    txtProperties(1).Text = objFunction.OID
    txtProperties(2).Text = objFunction.Owner
    Set objItem = cboProperties(0).ComboItems.Add(, , objFunction.Language, "language")
    objItem.Selected = True
    If objFunction.Returns = "opaque" Then
      Set objItem = cboProperties(1).ComboItems.Add(, , objFunction.Returns, "opaque")
    Else
      Set objItem = cboProperties(1).ComboItems.Add(, , objFunction.Returns, "type")
    End If
    objItem.Selected = True
    For Each vArgument In objFunction.Arguments
      lvProperties(0).ListItems.Add , , vArgument, "type", "type"
    Next vArgument
    chkProperties(0).Value = Bool2Bin(objFunction.Cachable)
    chkProperties(1).Value = Bool2Bin(objFunction.Strict)
    hbxProperties(0).Text = objFunction.Comment
    hbxProperties(1).Text = objFunction.Source
    
    'You can edit functions in 7.2 :-)
    If (frmMain.svr.dbVersion.VersionNum >= 7.2) And Not objFunction.SystemObject Then
      hbxProperties(1).BackColor = &H80000005
      hbxProperties(1).Locked = False
    End If
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  hbxProperties(1).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.hbxProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objFunction Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objFunction.Cachable)
    chkProperties(1).Value = Bool2Bin(objFunction.Strict)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.chkProperties_Click"
End Sub


