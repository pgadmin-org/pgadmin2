VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmColumn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Column"
   ClientHeight    =   6870
   ClientLeft      =   4005
   ClientTop       =   1920
   ClientWidth     =   5520
   Icon            =   "frmColumn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3285
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4410
      TabIndex        =   12
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
      TabPicture(0)   =   "frmColumn.frx":06C2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblProperties(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblProperties(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProperties(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProperties(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "hbxProperties(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtProperties(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtProperties(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtProperties(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkProperties(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtProperties(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkProperties(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtProperties(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "&Properties 2"
      TabPicture(1)   =   "frmColumn.frx":06DE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtProperties(5)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cboProperties(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblProperties(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblProperties(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   6
         Left            =   1932
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "The defined dimension array. E.g [1][][3]"
         Top             =   2256
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   5
         Left            =   -73044
         Locked          =   -1  'True
         TabIndex        =   20
         ToolTipText     =   $"frmColumn.frx":06FA
         Top             =   720
         Width           =   3390
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Primary key?"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Is this column a primary key?"
         Top             =   4344
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         Height          =   285
         Index           =   4
         Left            =   1935
         TabIndex        =   7
         ToolTipText     =   "A default value for the column. This may be a literal value, user function or niladic function."
         Top             =   3444
         Width           =   3390
      End
      Begin VB.CheckBox chkProperties 
         Alignment       =   1  'Right Justify
         Caption         =   "Restrict null values?"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Should null values be restricted in this column?"
         Top             =   3936
         Width           =   1995
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "The numeric scale of the column (applicable to numeric columns only)."
         Top             =   3036
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "The defined length of the column."
         Top             =   2628
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The ordinal position of the column in the table."
         Top             =   1485
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
      Begin HighlightBox.HBX hbxProperties 
         Height          =   1380
         Index           =   0
         Left            =   132
         TabIndex        =   10
         ToolTipText     =   "Comments about the column."
         Top             =   4800
         Width           =   5196
         _ExtentX        =   9155
         _ExtentY        =   2434
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
         Caption         =   "Comments"
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   2
         ToolTipText     =   "The table that the column will be part of."
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
         TabIndex        =   4
         ToolTipText     =   "The data type of the column."
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
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   2
         Left            =   -73050
         TabIndex        =   21
         ToolTipText     =   "Storage technique for the data type. If specified, must be 'plain', 'external', 'extended', or 'main'; the default is 'plain'."
         Top             =   1125
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Array dimension"
         Height          =   192
         Index           =   9
         Left            =   132
         TabIndex        =   25
         Top             =   2304
         Width           =   1164
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Statistics"
         Height          =   192
         Index           =   1
         Left            =   -74844
         TabIndex        =   23
         Top             =   768
         Width           =   636
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Storage"
         Height          =   192
         Index           =   8
         Left            =   -74856
         TabIndex        =   22
         Top             =   1176
         Width           =   672
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Default value"
         Height          =   192
         Index           =   7
         Left            =   132
         TabIndex        =   19
         Top             =   3492
         Width           =   996
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Length"
         Height          =   192
         Index           =   6
         Left            =   144
         TabIndex        =   18
         Top             =   2688
         Width           =   492
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Numeric Scale"
         Height          =   192
         Index           =   5
         Left            =   132
         TabIndex        =   17
         Top             =   3084
         Width           =   1128
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Data type"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   16
         Top             =   1935
         Width           =   690
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Ordinal position"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   1530
         Width           =   1080
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Table"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   1170
         Width           =   405
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":078F
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":08E9
            Key             =   "type"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":0E83
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":141D
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmColumn.frx":1AEF
            Key             =   "storage"
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
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmColumn.frm - Edit/Create a Column

Option Explicit

Dim szDatabase As String
Dim szNamespace As String
Dim szMode As String
Dim frmCallingForm As Form
Dim objColumn As pgColumn

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim szOldName As String
Dim szTemp As String

  If Not frmCallingForm Is Nothing Then
    If Not frmCallingForm.Visible Then
      MsgBox "The form that called this form has been destroyed!", vbExclamation, "Error"
      Unload Me
      Exit Sub
    End If
  End If
  
  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a Column name!", vbExclamation, "Error"
    tabProperties.Tab = 0
    txtProperties(0).SetFocus
    Exit Sub
  End If
  If cboProperties(1).Text = "" Then
    MsgBox "You must specify a column type!", vbExclamation, "Error"
    tabProperties.Tab = 0
    cboProperties(1).SetFocus
    Exit Sub
  End If
  
  Select Case szMode
    Case "TA"
      For Each objItem In frmCallingForm.lvProperties(0).ListItems
        If objItem.Text = txtProperties(0).Text Then
          MsgBox "A column with that name already exists!", vbExclamation, "Error"
          tabProperties.Tab = 0
          txtProperties(0).SetFocus
          Exit Sub
        End If
      Next objItem
      
      If ((cboProperties(1).Text = "serial") Or (cboProperties(1).Text = "serial8")) Then
        Set objItem = frmCallingForm.lvProperties(0).ListItems.Add(, , txtProperties(0).Text, "sequence", "sequence")
      Else
        Set objItem = frmCallingForm.lvProperties(0).ListItems.Add(, , txtProperties(0).Text, "column", "column")
      End If
      objItem.SubItems(1) = txtProperties(1).Text
      
      'verify if column is array
      szTemp = cboProperties(1).Text
      If Right(szTemp, 2) = "[]" Then
        szTemp = Mid(szTemp, 1, Len(szTemp) - 2) & Replace(txtProperties(6).Text, " ", "")
      End If
      objItem.SubItems(2) = szTemp
      If Not txtProperties(2).Locked Then objItem.SubItems(3) = Val(txtProperties(2).Text)
      If Not txtProperties(3).Locked Then objItem.SubItems(3) = Val(txtProperties(2).Text) & ", " & Val(txtProperties(3).Text)
      objItem.SubItems(4) = txtProperties(4).Text
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
      
      'DJP 2002-01-10 Change column name last of all otherwise things get complex.
      If txtProperties(4).Tag = "Y" Then objColumn.Default = txtProperties(4).Text
      If txtProperties(5).Tag = "Y" Then objColumn.Statistics = Val(txtProperties(5).Text)
      If hbxProperties(0).Tag = "Y" Then objColumn.Comment = hbxProperties(0).Text
      If chkProperties(0).Tag = "Y" Then objColumn.NotNull = Bin2Bool(chkProperties(0).Value)
      If chkProperties(1).Tag = "Y" Then objColumn.PrimaryKey = Bin2Bool(chkProperties(1).Value)
      
      'update storage
      If ctx.dbVer >= 7.3 Then
        If objColumn.Storage <> cboProperties(2).Text Then objColumn.Storage = cboProperties(2).Text
      End If
      
      If txtProperties(0).Tag = "Y" Then
        szOldName = objColumn.Name
        frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(objColumn.Table).Columns.Rename szOldName, txtProperties(0).Text
        
        'Update the node text
        frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Tables(objColumn.Table).Columns(txtProperties(0).Text).Tag.Text = txtProperties(0).Text
      End If
      
      'Simulate a node click to refresh the ListColumn (only do this when updating a column).
      frmMain.tv_NodeClick frmMain.tv.SelectedItem
  End Select
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.cmdOK_Click"
End Sub

Public Sub Initialise(szDB As String, szNS As String, szMD As String, Optional Column As pgColumn, Optional frmCF As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.Initialise(" & QUOTE & szDB & QUOTE & ", " & QUOTE & szMD & QUOTE & ")", etFullDebug

Dim X As Integer
Dim objItem As ComboItem
Dim objDomain As pgDomain
Dim objType As pgType
Dim objNamespace As pgNamespace

  szDatabase = szDB
  szNamespace = szNS
  
  PatchForm Me
  
  'The mode indicates the way the form is being used:
  'MP = Called from frmMain, viewing the properties of a column object.
  'TA = Called from frmTable (ref: frmCallingForm), adding a new column object.
  
  szMode = szMD
  Set frmCallingForm = frmCF
    
  Select Case szMode
    Case "TA"
  
      'Create a new Column
      Me.Caption = "Create Column"
      
      'Unlock the edittable fields
      txtProperties(0).BackColor = &H80000005
      txtProperties(0).Locked = False
      cboProperties(1).BackColor = &H80000005
      hbxProperties(0).Locked = False
      hbxProperties(0).BackColor = &H80000005
      Set objItem = cboProperties(0).ComboItems.Add(, , frmCallingForm.txtProperties(0).Text, "table", "table")
      objItem.Selected = True
      txtProperties(1).Text = frmCallingForm.lvProperties(0).ListItems.Count + 1
      
      'Populate the Types combo
      'Pseudo types
      If ctx.dbVer >= 7.2 Then cboProperties(1).ComboItems.Add , , "serial8", "sequence", "sequence"
      cboProperties(1).ComboItems.Add , , "serial", "sequence", "sequence"
      
      If ctx.dbVer >= 7.3 Then
        'Add pg_catalog items first, unqualified
        For Each objDomain In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Domains
          cboProperties(1).ComboItems.Add , , fmtID(objDomain.Name), "domain", "domain"
        Next objDomain
        For Each objType In frmMain.svr.Databases(szDatabase).Namespaces("pg_catalog").Types
          cboProperties(1).ComboItems.Add , , fmtTypeName(objType), "type", "type"
        Next objType
        'Now add other items
        For Each objNamespace In frmMain.svr.Databases(szDatabase).Namespaces
          If (Not objNamespace.SystemObject) Or (objNamespace.Name = "public") Then
            For Each objDomain In objNamespace.Domains
              cboProperties(1).ComboItems.Add , , objDomain.FormattedID, "domain", "domain"
            Next objDomain
            For Each objType In objNamespace.Types
              cboProperties(1).ComboItems.Add , , fmtTypeName(objType), "type", "type"
            Next objType
          End If
        Next objNamespace
      Else
        For Each objDomain In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Domains
          cboProperties(1).ComboItems.Add , , objDomain.FormattedID, "domain", "domain"
        Next objDomain
        For Each objType In frmMain.svr.Databases(szDatabase).Namespaces(szNamespace).Types
          cboProperties(1).ComboItems.Add , , fmtTypeName(objType), "type", "type"
        Next objType
      End If
    
    Case "MP"
  
      'Display/Edit the specified Column.
      Set objColumn = Column
    
      Me.Caption = "Column: " & objColumn.Identifier
      
      txtProperties(0).Text = objColumn.Name
      txtProperties(1).Text = objColumn.Position
      If objColumn.Length = 0 Then
        txtProperties(2).Text = "Variable"
      Else
        txtProperties(2).Text = objColumn.Length
      End If
      If objColumn.DataType = "numeric" Then txtProperties(3).Text = objColumn.NumericScale
      txtProperties(4).Text = objColumn.Default
      If ctx.dbVer >= 7.3 Then
        Set objItem = cboProperties(0).ComboItems.Add(, , objColumn.Namespace & "." & objColumn.Table, "table")
        objItem.Selected = True
      Else
        Set objItem = cboProperties(0).ComboItems.Add(, , objColumn.Table, "table")
        objItem.Selected = True
      End If
      Set objItem = cboProperties(1).ComboItems.Add(, , objColumn.DataType, "type")
      objItem.Selected = True
      chkProperties(0).Value = Bool2Bin(objColumn.NotNull)
      chkProperties(1).Value = Bool2Bin(objColumn.PrimaryKey)
      hbxProperties(0).Text = objColumn.Comment
      If objColumn.Position > 0 Then
        hbxProperties(0).Locked = False
        hbxProperties(0).BackColor = &H80000005
      End If
      If ctx.dbVer >= 7.2 Then
        txtProperties(5).BackColor = &H80000005
        txtProperties(5).Locked = False
        txtProperties(5).Text = objColumn.Statistics
      End If
      
      'storage
      cboProperties(2).ComboItems.Add , "PLAIN", "PLAIN", "storage"
      cboProperties(2).ComboItems.Add , "EXTERNAL", "EXTERNAL", "storage"
      cboProperties(2).ComboItems.Add , "EXTENDED", "EXTENDED", "storage"
      cboProperties(2).ComboItems.Add , "MAIN", "MAIN", "storage"
      If ctx.dbVer >= 7.3 Then
        cboProperties(2).BackColor = &H80000005
        cboProperties(2).Locked = False
      End If
      cboProperties(2).ComboItems(objColumn.Storage).Selected = True
  End Select

  'Reset the Tags
  txtProperties(0).Tag = "N"
  txtProperties(4).Tag = "N"
  txtProperties(5).Tag = "N"
  hbxProperties(0).Tag = "N"
  chkProperties(0).Tag = "N"
  chkProperties(1).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.Initialise"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.hbxProperties_Change"
End Sub

Private Sub txtProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.txtProperties_Change"
End Sub

Private Sub chkProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.chkProperties_Click(" & Index & ")", etFullDebug

  If ctx.dbVer < 7.3 Then
    If Not (objColumn Is Nothing) Then
      chkProperties(0).Value = Bool2Bin(objColumn.NotNull)
      chkProperties(1).Value = Bool2Bin(objColumn.PrimaryKey)
    Else
      If szMode = "TA" Then
        If Not (frmCallingForm.objTable Is Nothing) Then
          chkProperties(0).Value = 0
          chkProperties(1).Value = 0
        End If
      Else
        chkProperties(0).Value = 0
        chkProperties(1).Value = 0
      End If
    End If
  Else
    If Index = 0 Then chkProperties(0).Tag = "Y"
    If Index = 1 Then chkProperties(1).Tag = "Y"
  End If

  'Primary Key implies Not Null
  If chkProperties(1).Value = 1 Then chkProperties(0).Value = 1
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.chkProperties_Click"
End Sub

Private Sub cboProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmColumn.cboProperties_Click(" & Index & ")", etFullDebug

Dim objColumn As pgColumn

  If ((Index = 1) And (szMode = "TA")) Then
     
    'Lock first
    txtProperties(2).BackColor = &H8000000F
    txtProperties(2).Locked = True
    txtProperties(3).BackColor = &H8000000F
    txtProperties(3).Locked = True
    
    'Now unlock based on the data type
    Select Case cboProperties(1).Text
      Case "numeric"
        txtProperties(2).BackColor = &H80000005
        txtProperties(2).Locked = False
        txtProperties(3).BackColor = &H80000005
        txtProperties(3).Locked = False
      Case "char"
        txtProperties(2).BackColor = &H80000005
        txtProperties(2).Locked = False
      Case "varchar"
        txtProperties(2).BackColor = &H80000005
        txtProperties(2).Locked = False
    End Select
    
    'array column
    If Right(cboProperties(1).Text, 2) = "[]" Then
        txtProperties(6).BackColor = &H80000005
        txtProperties(6).Locked = False
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmColumn.cboProperties_Click"
End Sub
