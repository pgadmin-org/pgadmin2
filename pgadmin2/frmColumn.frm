VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmColumn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Column"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmColumn.frx":0000
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
      TabIndex        =   12
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   13
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
      TabPicture(0)   =   "frmColumn.frx":058A
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
      Tab(0).Control(5)=   "lblProperties(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblProperties(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "hbxProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProperties(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProperties(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtProperties(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkProperties(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtProperties(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkProperties(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Primary key?"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         ToolTipText     =   "Is this column a primary key?"
         Top             =   4185
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         Index           =   5
         Left            =   1935
         TabIndex        =   8
         ToolTipText     =   "A default value for the column. This may be a literal value, user function or niladic function."
         Top             =   3465
         Width           =   3390
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Restrict null values?"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   9
         ToolTipText     =   "Should null values be restricted in this column?"
         Top             =   3870
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "The numeric scale of the column (applicable to numeric columns only)."
         Top             =   3060
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "The defined length of the column."
         Top             =   2655
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "The ordinal position of the column in the table."
         Top             =   1845
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         Index           =   0
         Left            =   1935
         TabIndex        =   1
         ToolTipText     =   "The name of the column."
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
         ToolTipText     =   "The columns OID (Object ID) in the PostgreSQL Database."
         Top             =   1080
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   1680
         Index           =   0
         Left            =   135
         TabIndex        =   11
         ToolTipText     =   "Comments about the column."
         Top             =   4500
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   2963
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
         Index           =   0
         Left            =   1935
         TabIndex        =   3
         ToolTipText     =   "The table that the column will be part of."
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
         TabIndex        =   5
         ToolTipText     =   "The data type of the column."
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
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Default value"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   21
         Top             =   3510
         Width           =   945
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   20
         Top             =   2700
         Width           =   495
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Numeric Precision"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   19
         Top             =   3105
         Width           =   1275
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Data type"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   18
         Top             =   2295
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Ordinal position"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   17
         Top             =   1890
         Width           =   1080
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Top             =   1530
         Width           =   405
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   14
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":05A6
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":0700
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":0C9A
            Key             =   "sequence"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmColumn.frm - Edit/Create a Column

Option Explicit

Dim szDatabase As String
Dim szMode As String
Dim frmCallingForm As Form
Dim objColumn As pgColumn
Dim bNoPrimaryKey As Boolean

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim szOldName As String

  If Not frmCallingForm.Visible Then
    MsgBox "The form that called this form has been destroyed!", vbExclamation, "Error"
    Unload Me
    Exit Sub
  End If
  
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Column name!", vbExclamation, "Error"
    tabProperties.tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must specify a column type!", vbExclamation, "Error"
    tabProperties.tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  
  Select Case szMode
    Case "TA"
      For Each objItem In frmCallingForm.lvProperties(0).ListItems
        If objItem.Text = txtProperties(0).Text Then
          MsgBox "A column with that name already exists!", vbExclamation, "Error"
          tabProperties.tab = 0
          txtProperties(0).SetFocus
          Exit Sub
        End If
      Next objItem
      
      If cboProperties(1).Text = "serial" Then
        Set objItem = frmCallingForm.lvProperties(0).ListItems.Add(, , txtProperties(0).Text, "sequence", "sequence")
      Else
        Set objItem = frmCallingForm.lvProperties(0).ListItems.Add(, , txtProperties(0).Text, "column", "column")
      End If
      objItem.SubItems(1) = txtProperties(2).Text
      objItem.SubItems(2) = cboProperties(1).Text
      If Not txtProperties(3).Locked Then objItem.SubItems(3) = Val(txtProperties(3).Text)
      If Not txtProperties(4).Locked Then objItem.SubItems(3) = Val(txtProperties(3).Text) & ", " & Val(txtProperties(4).Text)
      objItem.SubItems(4) = txtProperties(5).Text
      If chkProperties(0).Value = 1 Then
        objItem.SubItems(5) = "Yes"
      Else
        objItem.SubItems(5) = "No"
      End If
      If chkProperties(1).Value = 1 Then
        objItem.SubItems(6) = "Yes"
      Else
        objItem.SubItems(6) = "No"
      End If
      objItem.SubItems(7) = hbxProperties(0).Text
      
      frmCallingForm.lvProperties(0).Tag = "Y"
      
    Case "MP"
      StartMsg "Updating Column..."
      If txtProperties(0).Tag = "Y" Then
        szOldName = objColumn.Name
        frmMain.svr.Databases(szDatabase).Tables(cboProperties(0).Text).Columns.Rename szOldName, txtProperties(0).Text
        
        'Update the node text
        For Each objNode In frmMain.tv.Nodes
          If (InStr(1, objNode.FullPath, "\" & szDatabase & "\") <> 0) And (InStr(1, objNode.FullPath, "\" & cboProperties(0).Text & "\") <> 0) Then
            If (Left(objNode.Key, 4) = "COL-") And (objNode.Parent.Parent.Text = cboProperties(0).Text) And (objNode.Parent.Parent.Parent.Parent.Text = szDatabase) And (objNode.Text = szOldName) Then
              objNode.Text = txtProperties(0).Text
            End If
          End If
        Next objNode
      End If
      If txtProperties(5).Tag = "Y" Then objColumn.Default = txtProperties(5).Text
      If hbxProperties(0).Tag = "Y" Then objColumn.Comment = hbxProperties(0).Text
  End Select
  
  'Simulate a node click to refresh the ListColumn
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szMD As String, Optional Column As pgColumn, Optional frmCF As Form, Optional bNoPKey As Boolean)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.Initialise(" & QUOTE & szDB & QUOTE & ", " & QUOTE & szMD & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ComboItem
Dim objType As pgType

  szDatabase = szDB
  
  'The mode indicates the way the form is being used:
  'MP = Called from frmMain, viewing the properties of a column object.
  'TA = Called from frmTable (ref: frmCallingForm), adding a new column object.
  
  szMode = szMD
  Set frmCallingForm = frmCF
    
  Select Case szMode
    Case "TA"
  
      'Create a new Column
      Me.Caption = "Create Column"
      bNoPrimaryKey = bNoPKey
      
      'Unlock the edittable fields
      txtProperties(0).BackColor = &H80000005
      txtProperties(0).Locked = False
      cboProperties(1).BackColor = &H80000005
      Set objItem = cboProperties(0).ComboItems.Add(, , frmCallingForm.txtProperties(0).Text, "table", "table")
      objItem.Selected = True
      txtProperties(2).Text = frmCallingForm.lvProperties(0).ListItems.Count + 1
      
      'Populate the Types combo
      cboProperties(1).ComboItems.Add , , "serial", "sequence", "sequence"
      For Each objType In frmMain.svr.Databases(szDatabase).Types
        If Left(objType.Name, 1) <> "_" Then cboProperties(1).ComboItems.Add , , objType.Name, "type", "type"
      Next objType
    
    Case "MP"
  
      'Display/Edit the specified Column.
      Set objColumn = Column
    
      Me.Caption = "Column: " & objColumn.Identifier
      txtProperties(0).Text = objColumn.Name
      txtProperties(1).Text = objColumn.OID
      txtProperties(2).Text = objColumn.Position
      If objColumn.Length = 0 Then
        txtProperties(3).Text = "Variable"
      Else
        txtProperties(3).Text = objColumn.Length
      End If
      If objColumn.DataType = "numeric" Then txtProperties(4).Text = objColumn.NumericScale
      txtProperties(5).Text = objColumn.Default
      Set objItem = cboProperties(0).ComboItems.Add(, , objColumn.Table, "table")
      objItem.Selected = True
      Set objItem = cboProperties(1).ComboItems.Add(, , objColumn.DataType, "type")
      objItem.Selected = True
      chkProperties(0).Value = Bool2Bin(objColumn.NotNull)
      chkProperties(1).Value = Bool2Bin(objColumn.PrimaryKey)
      hbxProperties(0).Text = objColumn.Comment
  End Select

  'Reset the Tags
  txtProperties(0).Tag = "N"
  txtProperties(5).Tag = "N"
  hbxProperties(0).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.txtProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.chkProperties_Click(" & Index & ")", etFullDebug

  If Not (objColumn Is Nothing) Then
    chkProperties(0).Value = Bool2Bin(objColumn.NotNull)
    chkProperties(1).Value = Bool2Bin(objColumn.PrimaryKey)
  ElseIf bNoPrimaryKey Then
    chkProperties(1).Value = 0
  Else
    'Primary Key implicitly implies Not Null
    If chkProperties(1).Value = 1 Then chkProperties(0).Value = 1
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.chkProperties_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering frmColumn.cboProperties_Click(" & Index & ")", etFullDebug

Dim objColumn As pgColumn

  If ((Index = 1) And (szMode = "TA")) Then
     
    'Lock first
    txtProperties(3).BackColor = &H8000000F
    txtProperties(3).Locked = True
    txtProperties(4).BackColor = &H8000000F
    txtProperties(4).Locked = True
    
    'Now unlock based on the data type
    Select Case cboProperties(1).Text
      Case "numeric"
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
        txtProperties(4).BackColor = &H80000005
        txtProperties(4).Locked = False
      Case "char"
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
      Case "varchar"
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
    End Select
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, "frmColumn.cboProperties_Click"
End Sub
