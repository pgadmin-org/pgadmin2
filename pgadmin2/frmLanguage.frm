VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmLanguage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Language"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLanguage.frx":058A
            Key             =   "function"
         EndProperty
      EndProperty
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
      TabPicture(0)   =   "frmLanguage.frx":0B24
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "hbxProperties(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtProperties(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtProperties(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkProperties(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   3
         ToolTipText     =   "The name of a previously registered function that will be called to execute the PL procedures."
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
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Trusted?"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   4
         ToolTipText     =   $"frmLanguage.frx":0B40
         Top             =   1935
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the language."
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
         ToolTipText     =   "The languages OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   3840
         Index           =   0
         Left            =   135
         TabIndex        =   5
         ToolTipText     =   "Comments about the language."
         Top             =   2340
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   6773
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
         Caption         =   "Handler"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
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
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmLanguage.frm - Edit/Create a Language

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim objLanguage As pgLanguage

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLanguage.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLanguage.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLanguage.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Language name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(0).Text = "" Then
    MsgBox "You must select a Language handler!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Language..."
    frmMain.svr.Databases(szDatabase).Languages.Add txtProperties(0).Text, Bin2Bool(chkProperties(0).Value), Left(cboProperties(0).Text, Len(cboProperties(0).Text) - 2), hbxProperties(0).Text
    
    'Add a new node and update the text on the parent
    For Each objNode In frmMain.tv.Nodes
      If Left(objNode.Key, 4) <> "SVR-" Then
        If (Left(objNode.Key, 4) = "LNG+") And (objNode.Parent.Text = szDatabase) Then
          frmMain.tv.Nodes.Add objNode.Key, tvwChild, "LNG-" & GetID, txtProperties(0).Text, "language"
          objNode.Text = "Languages (" & objNode.Children & ")"
        End If
      End If
    Next objNode
    
  Else
    StartMsg "Updating Language..."
    If hbxProperties(0).Tag = "Y" Then objLanguage.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListLanguage
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLanguage.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, Optional Language As pgLanguage)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLanguage.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objFunction As pgFunction
Dim objItem As ComboItem
  
  szDatabase = szDB
  
  If Language Is Nothing Then
  
    'Create a new Language
    bNew = True
    Me.Caption = "Create Language"
    
    'Load the combo
    For Each objFunction In frmMain.svr.Databases(szDatabase).Functions
      If (objFunction.Returns = "opaque") And (objFunction.Arguments.Count = 0) Then
        cboProperties(0).ComboItems.Add , , objFunction.Identifier, "function"
      End If
    Next objFunction
  
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    
  Else
  
    'Display/Edit the specified Language.
    Set objLanguage = Language
    bNew = False
    
    cboProperties(0).BackColor = &H8000000F
    Me.Caption = "Language: " & objLanguage.Identifier
    txtProperties(0).Text = objLanguage.Name
    txtProperties(1).Text = objLanguage.OID
    Set objItem = cboProperties(0).ComboItems.Add(, , objLanguage.Handler & "()", "function")
    objItem.Selected = True
    chkProperties(0).Value = Bool2Bin(objLanguage.Trusted)
    hbxProperties(0).Text = objLanguage.Comment
  End If
  
  'Reset the Tags
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLanguage.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLanguage.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLanguage.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLanguage.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLanguage.txtProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmLanguage.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objLanguage Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objLanguage.Trusted)
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmLanguage.chkProperties_Click"
End Sub

