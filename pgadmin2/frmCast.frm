VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmCast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cast"
   ClientHeight    =   6885
   ClientLeft      =   7530
   ClientTop       =   1875
   ClientWidth     =   5520
   Icon            =   "frmCast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
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
      TabPicture(0)   =   "frmCast.frx":0BC2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboProperties(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The Casts OID (Object ID) in the PostgreSQL Database."
         Top             =   675
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   2
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
         Index           =   1
         Left            =   1935
         TabIndex        =   3
         Top             =   1485
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
         TabIndex        =   4
         Top             =   1890
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
         TabIndex        =   5
         ToolTipText     =   "The data type returned by the Cast."
         Top             =   2280
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
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Top             =   720
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Type target"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   1575
         Width           =   810
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Type source"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Context"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   2400
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   7
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
            Picture         =   "frmCast.frx":0BDE
            Key             =   "function"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCast.frx":1178
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCast.frx":1712
            Key             =   "cast"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmCast.frm - Edit/Create a Cast

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim objCast As pgCast

Public Sub Initialise(szDB As String, Optional Cast As pgCast)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCast.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ComboItem
Dim objType As pgType
  
  szDatabase = szDB
  
  PatchForm Me
  
  If Cast Is Nothing Then
  
    'Create a new Cast
    bNew = True
    Me.Caption = "Create Cast"
  
    For X = 0 To 3
      cboProperties(X).BackColor = &H80000005
    Next X
  
    'Load the combo
    For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
      cboProperties(0).ComboItems.Add , , fmtID(objType.Name), "type"
    Next
  
    cboProperties(3).ComboItems.Add , , "Assignment", "cast"
    cboProperties(3).ComboItems.Add , , "Explicit", "cast"
    cboProperties(3).ComboItems.Add , , "Implicit", "cast"
    cboProperties(3).ComboItems(2).Selected = True
  
  Else
  
    'Display/Edit the specified Cast.
    Set objCast = Cast
    bNew = False
    
    Me.Caption = "Cast: " & objCast.Identifier
    txtProperties(0).Text = objCast.Oid
    cboProperties(0).ComboItems.Add , , fmtID(objCast.Source), "type"
    cboProperties(0).ComboItems(1).Selected = True
    cboProperties(1).ComboItems.Add , , fmtID(objCast.Target), "type"
    cboProperties(1).ComboItems(1).Selected = True
    If Len(objCast.Funct) > 0 Then
      cboProperties(2).ComboItems.Add , , fmtID(objCast.Funct), "function"
      cboProperties(2).ComboItems(1).Selected = True
    End If
    
    Select Case objCast.Context
      Case "ASSIGNMENT"
        cboProperties(3).ComboItems.Add , , "Assignment", "cast"
      Case "EXPLICIT"
        cboProperties(3).ComboItems.Add , , "Explicit", "cast"
      Case "IMPLICIT"
        cboProperties(3).ComboItems.Add , , "Implicit", "cast"
    End Select
    cboProperties(3).ComboItems(1).Selected = True
    
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCast.Initialise"
End Sub

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCast.cboProperties_Click(" & Index & ")", etFullDebug

Dim objFunction As pgFunction
Dim objType As pgType

  'Populate the StateFunction Combo. StateFunctions have 1 or 2 arguments
  If objCast Is Nothing Then
    If (Index = 0) Then
      cboProperties(1).ComboItems.Clear
      cboProperties(1).Text = ""
      cboProperties(2).ComboItems.Clear
      cboProperties(2).Text = ""
    
      cboProperties(1).ComboItems.Clear
      For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
        If objType.Name <> cboProperties(0).Text Then
          cboProperties(1).ComboItems.Add , , fmtID(objType.Name), "type"
        End If
      Next
    ElseIf (Index = 1) Then
      cboProperties(2).ComboItems.Clear
      cboProperties(2).Text = ""
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Functions
        If objFunction.Arguments.Count = 1 Then
          If objFunction.Arguments(1) = cboProperties(0).Text And objFunction.Returns = cboProperties(1).Text Then
            cboProperties(2).ComboItems.Add , , fmtID(objFunction.Name), "function"
          End If
        End If
      Next
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCast.cboProperties_Click"
End Sub

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCast.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCast.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmCast.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewCast As pgCast
Dim szContext As String

  'Check the data
  If cboProperties(0).Text = "" Then
    MsgBox "You must select an source type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must select a target type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Cast..."
    
    szContext = UCase(cboProperties(3).Text)
    Set objNewCast = frmMain.svr.Databases(szDatabase).Casts.Add(cboProperties(0).Text, cboProperties(1).Text, cboProperties(2).Text, szContext)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Casts.Tag
    frmMain.tv.Nodes.Add objNode.Key, tvwChild, "CST-" & GetID, objNewCast.Identifier, "cast"
    objNode.Text = "Casts (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Cast..."
  End If
  
  'Simulate a node click to refresh the ListCast
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmCast.cmdOK_Click"
End Sub


