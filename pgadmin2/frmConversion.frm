VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversion"
   ClientHeight    =   6885
   ClientLeft      =   5070
   ClientTop       =   1770
   ClientWidth     =   5520
   Icon            =   "frmConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   7
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
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Properties"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkProperties(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Default conversion?"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         ToolTipText     =   "This controls whether the constraint can be deferred to the end of the transaction."
         Top             =   3060
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The Conversions OID (Object ID) in the PostgreSQL Database."
         Top             =   1440
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the Conversion."
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
         ToolTipText     =   "The Conversions OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1920
         TabIndex        =   4
         ToolTipText     =   "The table that the foreign key will be part of."
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
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "The table referenced by the foreign key."
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
         Left            =   1920
         TabIndex        =   6
         ToolTipText     =   "The table referenced by the foreign key."
         Top             =   2640
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
         Caption         =   "Conversion procedure"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   14
         Top             =   2685
         Width           =   1560
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Source encoding"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Destination encoding"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   12
         Top             =   2295
         Width           =   1500
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
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   720
         Width           =   420
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   540
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
            Picture         =   "frmConversion.frx":08CA
            Key             =   "function"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConversion.frx":0F9C
            Key             =   "encoding"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmConversion.frm - Edit/Create a Conversion

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objConversion As pgConversion

Public Sub Initialise(szDB As String, szNS As String, Optional Conversion As pgConversion)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConversion.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim bFound As Boolean
Dim rs As Recordset
Dim ii As Integer
Dim objFunction As pgFunction
Dim vData
  
  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  If Conversion Is Nothing Then
  
    'Create a new Conversion
    bNew = True
    Me.Caption = "Create Conversion"
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    cboProperties(2).BackColor = &H80000005
    
    'load encoding
    ii = 0
    bFound = True
    While bFound
      Set rs = frmMain.svr.Databases(szDatabase).Execute("select pg_encoding_to_char(" & ii & ") ")
      If Len(rs.Fields(0).Value) <= 0 Then
        bFound = False
      Else
        cboProperties(0).ComboItems.Add , , rs.Fields(0).Value, "encoding", "encoding"
        cboProperties(1).ComboItems.Add , , rs.Fields(0).Value, "encoding", "encoding"
      End If
      ii = ii + 1
    Wend
    
    'load function
    vData = Array("pg_catalog", szNamespace)
    For ii = 0 To UBound(vData)
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces(CStr(vData(ii))).Functions
        'verify param function
        If objFunction.Arguments.Count = 5 Then
          'e.g ascii_to_utf8(int4, int4, cstring, cstring, int4)
          If objFunction.Arguments(1) = "int4" And objFunction.Arguments(2) = "int4" And objFunction.Arguments(3) = "cstring" And objFunction.Arguments(4) = "cstring" And objFunction.Arguments(5) = "int4" Then
            cboProperties(2).ComboItems.Add , , objFunction.Namespace & "." & objFunction.Name, "function", "function"
          End If
        End If
      Next
    Next
  Else
  
    'Display/Edit the specified Conversion.
    Set objConversion = Conversion
    bNew = False
    
    Me.Caption = "Conversion: " & objConversion.Identifier
    txtProperties(0).Text = objConversion.Name
    txtProperties(1).Text = objConversion.Oid
    txtProperties(2).Text = objConversion.Owner
    
    cboProperties(0).ComboItems.Add , , objConversion.ForEncoding, "encoding", "encoding"
    cboProperties(0).ComboItems(1).Selected = True
    cboProperties(1).ComboItems.Add , , objConversion.ToEncoding, "encoding", "encoding"
    cboProperties(1).ComboItems(1).Selected = True
    cboProperties(2).ComboItems.Add , , objConversion.Proc, "function", "function"
    cboProperties(2).ComboItems(1).Selected = True
    chkProperties(0).Value = Bool2Bin(objConversion.Default)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConversion.Initialise"
End Sub

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConversion.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConversion.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConversion.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objNewConversion As pgConversion
Dim vData
Dim ii As Integer
    
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a conversion name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  vData = Array("Source encoding", "Destination encoding", "Conversion procedure")
  For ii = 0 To 2
    If cboProperties(ii).Text = "" Then
      MsgBox "You must specify a " & vData(ii) & "!", vbExclamation, "Error"
      tabProperties.Tab = 0
      cboProperties(ii).SetFocus
      Exit Sub
    End If
  Next
  
  If bNew Then
    StartMsg "Creating Conversion..."
    Set objNewConversion = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Conversions.Add(txtProperties(0).Text, Bin2Bool(chkProperties(0).Value), cboProperties(0).SelectedItem.Text, cboProperties(1).SelectedItem.Text, cboProperties(2).SelectedItem.Text)
   
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Conversions.Tag
    Set objNewConversion.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "CNV-" & GetID, txtProperties(0).Text, "conversion")
    objNode.Text = "Conversions (" & objNode.Children & ")"
  Else
    StartMsg "Updating Conversion..."
  End If
  
  'Simulate a node click to refresh the ListConversion
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConversion.cmdOK_Click"
End Sub

Private Sub chkProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmConversion.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objConversion Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objConversion.Default)
  Else
    If Index = 0 Then chkProperties(0).Tag = "Y"
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmConversion.chkProperties_Click"
End Sub
