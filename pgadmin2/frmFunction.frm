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
         NumListImages   =   9
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
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":1372
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":1A44
            Key             =   "public"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":1B9E
            Key             =   "user"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":1CF8
            Key             =   "group"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunction.frx":23CA
            Key             =   "volatility"
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
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      TabPicture(0)   =   "frmFunction.frx":2CA4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboProperties(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "hbxProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkProperties(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkProperties(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkProperties(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "&Input/Output"
      TabPicture(1)   =   "frmFunction.frx":2CC0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblProperties(4)"
      Tab(1).Control(1)=   "lblProperties(5)"
      Tab(1).Control(2)=   "cboProperties(2)"
      Tab(1).Control(3)=   "cboProperties(1)"
      Tab(1).Control(4)=   "lvProperties(0)"
      Tab(1).Control(5)=   "cmdAdd"
      Tab(1).Control(6)=   "cmdRemove"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Definition"
      TabPicture(2)   =   "frmFunction.frx":2CDC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "hbxProperties(1)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Security"
      TabPicture(3)   =   "frmFunction.frx":2CF8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvProperties(1)"
      Tab(3).Control(1)=   "fraAdd"
      Tab(3).Control(2)=   "cmdAddPrivilege"
      Tab(3).Control(3)=   "cmdRemovePrivilege"
      Tab(3).ControlCount=   4
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Security Definer?"
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   32
         ToolTipText     =   "Indicates whether the function is a security definer."
         Top             =   3375
         Width           =   1995
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Returns a Set?"
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   29
         ToolTipText     =   "Indicates whether the function returns a set (ie, multiple values of the specified data type)."
         Top             =   3015
         Width           =   1995
      End
      Begin VB.CommandButton cmdRemovePrivilege 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73515
         TabIndex        =   27
         ToolTipText     =   "Remove the selected entry."
         Top             =   3900
         Width           =   1230
      End
      Begin VB.CommandButton cmdAddPrivilege 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74865
         TabIndex        =   26
         ToolTipText     =   "Add the defined entry."
         Top             =   3900
         Width           =   1230
      End
      Begin VB.Frame fraAdd 
         Caption         =   "Define Privilege"
         Height          =   1815
         Left            =   -74865
         TabIndex        =   22
         Top             =   4380
         Width           =   5190
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Execute"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   23
            ToolTipText     =   "Give execute privilege to the selected entity."
            Top             =   1125
            Width           =   1590
         End
         Begin MSComctlLib.ImageCombo cboEntities 
            Height          =   330
            Left            =   1260
            TabIndex        =   24
            ToolTipText     =   "Select a user, group or 'PUBLIC'."
            Top             =   315
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483633
            Locked          =   -1  'True
            ImageList       =   "il"
         End
         Begin VB.Label lblProperties 
            AutoSize        =   -1  'True
            Caption         =   "User/Group"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   25
            Top             =   360
            Width           =   840
         End
      End
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
         ToolTipText     =   $"frmFunction.frx":2D14
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
         ToolTipText     =   $"frmFunction.frx":2DDD
         Top             =   2655
         Width           =   1995
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Cachable?"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   $"frmFunction.frx":2F81
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
         Height          =   1860
         Index           =   0
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "Comments about the function."
         Top             =   4140
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
      Begin HighlightBox.HBX hbxProperties 
         Height          =   5730
         Index           =   1
         Left            =   -74865
         TabIndex        =   15
         ToolTipText     =   $"frmFunction.frx":30BC
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
         ToolTipText     =   $"frmFunction.frx":3175
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
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3390
         Index           =   1
         Left            =   -74865
         TabIndex        =   28
         ToolTipText     =   "The access control list for the schema."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5980
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
            Text            =   "User/Group name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Privileges"
            Object.Width           =   4939
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   3
         Left            =   1935
         TabIndex        =   31
         ToolTipText     =   "Indicates whether the function's result depends only on its input arguments, or is affected by outside factors."
         Top             =   3690
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
         Caption         =   "Volatility"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   30
         Top             =   3780
         Width           =   570
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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmFunction.frm - Edit/Create a Function

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim szUsers() As String
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
    Case "domain"
      lvProperties(0).ListItems.Add , , cboProperties(2).Text, "domain", "domain"
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
Dim objNewFunction As pgFunction
Dim szArguments As String
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant
Dim szIdentifier As String

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

  'Get the identifier/arguments in case we need it
  For Each objItem In lvProperties(0).ListItems
    szArguments = szArguments & objItem.Text & ", "
  Next objItem
  If Len(szArguments) > 2 Then szArguments = Left(szArguments, Len(szArguments) - 2)
  szIdentifier = fmtID(txtProperties(0).Text) & "(" & szArguments & ")"
  
  If bNew Then
    StartMsg "Creating Function..."
    Set objNewFunction = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions.Add(txtProperties(0).Text, szArguments, cboProperties(1).Text, hbxProperties(1).Text, cboProperties(0).Text, Bin2Bool(chkProperties(0).Value), Bin2Bool(chkProperties(1).Value), hbxProperties(0).Text, cboProperties(3).Text, Bin2Bool(chkProperties(3).Value))
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions.Tag
    Set objNewFunction.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "FNC-" & GetID, txtProperties(0).Text & "(" & szArguments & ")", "function")
    objNode.Text = "Functions (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Function..."
    If hbxProperties(0).Tag = "Y" Then objFunction.Comment = hbxProperties(0).Text
    If hbxProperties(1).Tag = "Y" Then objFunction.Source = hbxProperties(1).Text
  End If
  
  'Set the ACL on the Database as required
  If lvProperties(1).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then
        If vEntity = "PUBLIC" Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions(szIdentifier).Revoke vEntity, aclAll
        ElseIf Left(vEntity, 6) = "GROUP " Then
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions(szIdentifier).Revoke "GROUP " & fmtID(Mid(vEntity, 7)), aclAll
        Else
          frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions(szIdentifier).Revoke fmtID(vEntity), aclAll
        End If
      End If
    Next vEntity
    
    'Now Grant the new permissions
    For Each objItem In lvProperties(1).ListItems
      If objItem.Icon = "group" Then
        szEntity = "GROUP " & fmtID(objItem.Text)
      ElseIf objItem.Icon = "public" Then
        szEntity = "PUBLIC"
      Else
        szEntity = fmtID(objItem.Text)
      End If
      lACL = 0
      If InStr(1, objItem.SubItems(1), "Execute") <> 0 Then lACL = lACL + aclExecute
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions(szIdentifier).Grant szEntity, lACL
    Next objItem
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

Public Sub Initialise(szDB As String, szNS As String, Optional oFunction As pgFunction)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objLanguage As pgLanguage
Dim objDomain As pgDomain
Dim objType As pgType
Dim objTable As pgTable
Dim objNamespace As pgNamespace
Dim objItem As ComboItem
Dim objUser As pgUser
Dim objGroup As pgGroup
Dim vArgument As Variant
Dim objLItem As ListItem
Dim szUserlist As String
Dim szAccesslist As String
Dim szAccess() As String
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  'Unlock the edittable fields
  If ctx.dbVer >= 7.3 Then
    cmdAddPrivilege.Enabled = True
    cmdRemovePrivilege.Enabled = True
    lvProperties(1).BackColor = &H80000005
    cboEntities.BackColor = &H80000005
    chkPrivilege(0).Enabled = True
    chkProperties(2).Enabled = True
    chkProperties(3).Enabled = True
  Else
    chkProperties(0).Enabled = True
  End If
  
  If oFunction Is Nothing Then
  
    'Create a new Function
    bNew = True
    Me.Caption = "Create Function"
    
    'Load the combo
    For Each objLanguage In frmMain.svr.Databases(szDatabase).Languages
      cboProperties(0).ComboItems.Add , , objLanguage.Identifier, "language"
    Next objLanguage
    If ctx.dbVer < 7.3 Then
      cboProperties(1).ComboItems.Add , , "opaque", "opaque"
      cboProperties(2).ComboItems.Add , , "opaque", "opaque"
    End If
    
    If ctx.dbVer >= 7.3 Then
      'Load pg_catalog entries first, unqualified
      For Each objDomain In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Domains
        cboProperties(1).ComboItems.Add , , fmtID(objDomain.Name), "domain"
        cboProperties(2).ComboItems.Add , , fmtID(objDomain.Name), "domain"
      Next objDomain
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
        cboProperties(1).ComboItems.Add , , fmtTypeName(objType), "type"
        cboProperties(2).ComboItems.Add , , fmtTypeName(objType), "type"
      Next objType
      For Each objTable In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Tables
        cboProperties(1).ComboItems.Add , , fmtID(objTable.Name), "table"
        cboProperties(2).ComboItems.Add , , fmtID(objTable.Name), "table"
      Next objTable
      'Now load the rest
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objDomain In objNamespace.Domains
            cboProperties(1).ComboItems.Add , , objDomain.FormattedID, "domain"
            cboProperties(2).ComboItems.Add , , objDomain.FormattedID, "domain"
          Next objDomain
          For Each objType In objNamespace.Types
            cboProperties(1).ComboItems.Add , , fmtTypeName(objType), "type"
            cboProperties(2).ComboItems.Add , , fmtTypeName(objType), "type"
          Next objType
          For Each objTable In objNamespace.Tables
            cboProperties(1).ComboItems.Add , , objTable.FormattedID, "table"
            cboProperties(2).ComboItems.Add , , objTable.FormattedID, "table"
          Next objTable
        End If
      Next objNamespace
    Else
      For Each objDomain In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Domains
        cboProperties(1).ComboItems.Add , , objDomain.FormattedID, "domain"
        cboProperties(2).ComboItems.Add , , objDomain.FormattedID, "domain"
      Next objDomain
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types
        cboProperties(1).ComboItems.Add , , fmtTypeName(objType), "type"
        cboProperties(2).ComboItems.Add , , fmtTypeName(objType), "type"
      Next objType
      For Each objTable In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables
        cboProperties(1).ComboItems.Add , , objTable.FormattedID, "table"
        cboProperties(2).ComboItems.Add , , objTable.FormattedID, "table"
      Next objTable
    End If
  
    If ctx.dbVer >= 7.3 Then
      Set objItem = cboProperties(3).ComboItems.Add(, , "IMMUTABLE", "volatility", "volatility")
      cboProperties(3).ComboItems.Add , , "STABLE", "volatility", "volatility"
      cboProperties(3).ComboItems.Add , , "VOLATILE", "volatility", "volatility"
      objItem.Selected = True
      cboProperties(3).BackColor = &H80000005
    End If
    
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

    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)
    
  Else
  
    'Display/Edit the specified Function.
    Set objFunction = oFunction
    bNew = False
    
    Me.Caption = "Function: " & objFunction.Identifier
    txtProperties(0).Text = objFunction.Name
    txtProperties(1).Text = objFunction.OID
    txtProperties(2).Text = objFunction.Owner
    
    If ctx.dbVer >= 7.3 Then
      Set objItem = cboProperties(3).ComboItems.Add(, , UCase(objFunction.Volatility), "volatility", "volatility")
      objItem.Selected = True
    End If

    Set objItem = cboProperties(0).ComboItems.Add(, , objFunction.Language, "language")
    objItem.Selected = True
    If objFunction.Returns = "opaque" And ctx.dbVer < 7.3 Then
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
    chkProperties(2).Value = Bool2Bin(objFunction.RetSet)
    chkProperties(3).Value = Bool2Bin(objFunction.SecDef)
    hbxProperties(0).Text = objFunction.Comment
    hbxProperties(1).Text = objFunction.Source
    
    'You can edit functions in 7.2 :-)
    If (ctx.dbVer >= 7.2) And Not objFunction.SystemObject Then
      hbxProperties(1).BackColor = &H80000005
      hbxProperties(1).Locked = False
    End If
    
    ParseACL objFunction.ACL, szUserlist, szAccesslist
    szUsers = Split(szUserlist, "|")
    szAccess = Split(szAccesslist, "|")
    For X = 0 To UBound(szUsers)
      If UCase(Left(szUsers(X), 6)) = "GROUP " Then
        Set objLItem = lvProperties(1).ListItems.Add(, , Mid(szUsers(X), 7), "group", "group")
      Else
        If UCase(szUsers(X)) = "PUBLIC" Then
          Set objLItem = lvProperties(1).ListItems.Add(, , szUsers(X), "public", "public")
        Else
          Set objLItem = lvProperties(1).ListItems.Add(, , szUsers(X), "user", "user")
        End If
      End If
      objLItem.SubItems(1) = szAccess(X)
    Next X
    
  End If
  
  'Load the Entities combo
  If ctx.dbVer >= 7.3 Then
    cboEntities.ComboItems.Add , , "PUBLIC", "public"
    For Each objUser In frmMain.svr.Users
      cboEntities.ComboItems.Add , , objUser.Name, "user"
    Next objUser
    For Each objGroup In frmMain.svr.Groups
      cboEntities.ComboItems.Add , , objGroup.Name, "group"
    Next objGroup
    cboEntities.ComboItems(1).Selected = True
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  hbxProperties(1).Tag = "N"
  lvProperties(1).Tag = "N"
  
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
    chkProperties(2).Value = Bool2Bin(objFunction.RetSet)
    chkProperties(3).Value = Bool2Bin(objFunction.SecDef)
  Else
    chkProperties(2).Value = 0
    If ctx.dbVer < 7.3 Then chkProperties(3).Value = 0
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.chkProperties_Click"
End Sub


Private Sub cmdRemovePrivilege_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.cmdRemovePrivilege_Click()", etFullDebug

  If lvProperties(1).SelectedItem Is Nothing Then Exit Sub
  lvProperties(1).ListItems.Remove lvProperties(1).SelectedItem.Index
  lvProperties(1).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdRemovePrivilege_Click"
End Sub

Private Sub cmdAddPrivilege_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmFunction.cmdAddPrivilege_Click()", etFullDebug

Dim szAccess As String
Dim objItem As ListItem

  If cboEntities.Text = "" Then Exit Sub
  
  'Check the entry doesn't already exist
  For Each objItem In lvProperties(1).ListItems
    If (objItem.Text = cboEntities.SelectedItem.Text) And (objItem.SmallIcon = cboEntities.SelectedItem.Image) Then
      MsgBox "'" & objItem.Text & "' already appears in the Access Control List. If you wish to modify this entry, it must be removed, and then replaced.", vbExclamation, "Error"
      Exit Sub
    End If
  Next objItem
  
  'Build the access string
  If chkPrivilege(0).Value = 1 Then szAccess = szAccess & "Execute, "
  If Len(szAccess) > 2 Then szAccess = Left(szAccess, Len(szAccess) - 2)
  If szAccess = "" Then szAccess = "None"
  
  Set objItem = lvProperties(1).ListItems.Add(, , cboEntities.SelectedItem.Text, cboEntities.SelectedItem.Image, cboEntities.SelectedItem.Image)
  objItem.SubItems(1) = szAccess
  lvProperties(1).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmFunction.cmdAddPrivilege_Click"
End Sub
