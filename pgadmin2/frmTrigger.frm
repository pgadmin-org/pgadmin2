VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmTrigger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trigger"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmTrigger.frx":0000
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":06C2
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":081C
            Key             =   "function"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":0DB6
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrigger.frx":1350
            Key             =   "event"
         EndProperty
      EndProperty
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
      TabPicture(0)   =   "frmTrigger.frx":1C2A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboProperties(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "hbxProperties(0)"
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
      Tab(0).Control(15)=   "chkProperties(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkProperties(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.CheckBox chkProperties 
         Caption         =   "&Delete"
         Height          =   195
         Index           =   2
         Left            =   4545
         TabIndex        =   19
         Top             =   2475
         Width           =   780
      End
      Begin VB.CheckBox chkProperties 
         Caption         =   "&Update"
         Height          =   195
         Index           =   1
         Left            =   3195
         TabIndex        =   18
         Top             =   2475
         Width           =   870
      End
      Begin VB.CheckBox chkProperties 
         Caption         =   "&Insert"
         Height          =   195
         Index           =   0
         Left            =   1935
         TabIndex        =   17
         Top             =   2475
         Width           =   780
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the trigger."
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
         ToolTipText     =   "The triggers OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   3
         ToolTipText     =   "The name of a table."
         Top             =   1485
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   1
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "When the Trigger should fire."
         Top             =   1935
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         ImageList       =   "il"
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   2445
         Index           =   0
         Left            =   135
         TabIndex        =   7
         ToolTipText     =   "Comments about the trigger."
         Top             =   3735
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   4313
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
         Index           =   2
         Left            =   1935
         TabIndex        =   5
         ToolTipText     =   "Should the Trigger fire for each row or statement? As of PostgreSQL v7.1.2, Statement level triggers are not yet supported."
         Top             =   2835
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   3
         Left            =   1935
         TabIndex        =   6
         ToolTipText     =   "A user-supplied function that the trigger will execute."
         Top             =   3285
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         ImageList       =   "il"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Function"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   16
         Top             =   3375
         Width           =   615
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "For each"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   15
         Top             =   2925
         Width           =   630
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Event"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Executes"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   10
         Top             =   2025
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmTrigger.frm - Edit/Create a Trigger

Option Explicit

Dim bNew As Boolean
Dim szDatabase As String
Dim szNamespace As String
Dim objTrigger As pgTrigger

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewTrigger As pgTrigger
Dim szEvent As String
Dim szOldName As String

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Trigger name!", vbExclamation, "Error"
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
  If cboProperties(3).Text = "" Then
    MsgBox "You must select a function!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(3).SetFocus
    Exit Sub
  End If
  If Right(Trim(cboProperties(3).Text), 1) <> ")" Then
    MsgBox "The function must contain a pair of parentheses even if it takes no arguments!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(3).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    
    StartMsg "Creating Trigger..."
    
    If chkProperties(0).Value = 1 Then szEvent = szEvent & "INSERT OR "
    If chkProperties(1).Value = 1 Then szEvent = szEvent & "UPDATE OR "
    If chkProperties(2).Value = 1 Then szEvent = szEvent & "DELETE OR "
    If Len(szEvent) > 4 Then
      szEvent = Left(szEvent, Len(szEvent) - 4)
    Else
      MsgBox "You must select at least one event!", vbExclamation, "Error"
      tabProperties.Tab = 0
      chkProperties(0).SetFocus
      Exit Sub
    End If
    Set objNewTrigger = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).Text).Triggers.Add(txtProperties(0).Text, cboProperties(3).Text, cboProperties(1).Text, szEvent, cboProperties(2).Text, hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(cboProperties(0).Text).Triggers.Tag
    Set objNewTrigger.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "TRG-" & GetID, objNewTrigger.Identifier, "trigger")
    objNode.Text = "Triggers (" & objNode.Children & ")"
    
  Else

    'Update the triggername if required
    If txtProperties(0).Tag = "Y" Then
      szOldName = objTrigger.Name
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(objTrigger.Table).Triggers.Rename szOldName, txtProperties(0).Text
        
      'Update the node text
      frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(objTrigger.Table).Triggers(objTrigger.Identifier).Tag.Text = objTrigger.Identifier
    End If
    
    StartMsg "Updating Trigger..."
    If hbxProperties(0).Tag = "Y" Then objTrigger.Comment = hbxProperties(0).Text
  End If
  
  'Simulate a node click to refresh the ListTrigger
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, Optional Trigger As pgTrigger)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.Initialise(" & QUOTE & szDB & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objTable As pgTable
Dim objFunction As pgFunction
Dim objNamespace As pgNamespace
Dim objItem As ComboItem
  
  szDatabase = szDB
  szNamespace = szNS
    
  'Set the font
  For X = 0 To 1
    Set txtProperties(X).Font = ctx.Font
  Next X
  For X = 0 To 3
    Set cboProperties(X).Font = ctx.Font
  Next X
  Set hbxProperties(0).Font = ctx.Font
  
  If Trigger Is Nothing Then
  
    'Create a new Trigger
    bNew = True
    Me.Caption = "Create Trigger"
    
    'Load the combos
    For Each objTable In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables
      If Not objTable.SystemObject Then cboProperties(0).ComboItems.Add , , objTable.FormattedID, "table"
    Next objTable
    
    Set objItem = cboProperties(1).ComboItems.Add(, , "BEFORE", "trigger")
    objItem.Selected = True
    cboProperties(1).ComboItems.Add , , "AFTER", "trigger"
    Set objItem = cboProperties(2).ComboItems.Add(, , "ROW", "trigger")
    objItem.Selected = True
    
    If ctx.dbVer >= 7.3 Then
      'First load pg_catalog items, unqualified
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Functions
        If objFunction.Returns = "opaque" Then cboProperties(3).ComboItems.Add , , Mid(objFunction.FormattedID, 12), "function"
      Next objFunction
      'Now load the rest
      For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
        If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
          For Each objFunction In objNamespace.Functions
            If objFunction.Returns = "opaque" Then cboProperties(3).ComboItems.Add , , objFunction.FormattedID, "function"
          Next objFunction
        End If
      Next objNamespace
    Else
      For Each objFunction In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Functions
        If objFunction.Returns = "opaque" Then cboProperties(3).ComboItems.Add , , objFunction.FormattedID, "function"
      Next objFunction
    End If
    
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(1).BackColor = &H80000005
    
    'TODO - Possible bug here. Unlock the Function combo to allow the user to edit the arguments.
    'This should work but on initial testing, PostgreSQL gave an error that function() doesn't
    'exist despite the SQL clearly showng the use of function('arg1').
    cboProperties(3).Locked = False
    cboProperties(3).BackColor = &H80000005
    
  Else
  
    'Display/Edit the specified Trigger.
    Set objTrigger = Trigger
    bNew = False

    'We can rename triggers in 7.3
    If ctx.dbVer >= 7.3 Then
      txtProperties(0).BackColor = &H80000005
      txtProperties(0).Locked = False
    End If
    
    Me.Caption = "Trigger: " & objTrigger.Identifier
    txtProperties(0).Text = objTrigger.Name
    txtProperties(1).Text = objTrigger.OID
    Set objItem = cboProperties(0).ComboItems.Add(, , objTrigger.Table, "table")
    objItem.Selected = True
    Set objItem = cboProperties(1).ComboItems.Add(, , objTrigger.Executes, "trigger")
    objItem.Selected = True
    Set objItem = cboProperties(2).ComboItems.Add(, , objTrigger.ForEach, "trigger")
    objItem.Selected = True
    Set objItem = cboProperties(3).ComboItems.Add(, , objTrigger.TriggerFunction, "trigger")
    objItem.Selected = True
    SetChecks objTrigger.TriggerEvent
    hbxProperties(0).Text = objTrigger.Comment
  End If
  
  'Reset the Tags
  txtProperties(0).Tag = "N"
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.txtProperties_Change"
End Sub

Private Sub SetChecks(szData As String)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.SetChecks(" & QUOTE & szData & QUOTE & ")", etFullDebug

Static bSetting As Boolean

  'bSetting is used to prevent recursion which will occur as setting the checkboxes value triggers
  'the _Click event.
  
  If Not bSetting Then
    bSetting = True
    chkProperties(0).Value = 0
    If InStr(1, UCase(szData), "INSERT") <> 0 Then chkProperties(0).Value = 1
    chkProperties(1).Value = 0
    If InStr(1, UCase(szData), "UPDATE") <> 0 Then chkProperties(1).Value = 1
    chkProperties(2).Value = 0
    If InStr(1, UCase(szData), "DELETE") <> 0 Then chkProperties(2).Value = 1
    bSetting = False
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.SetChecks"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrigger.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objTrigger Is Nothing) Then
    SetChecks objTrigger.TriggerEvent
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrigger.chkProperties_Click"
End Sub

