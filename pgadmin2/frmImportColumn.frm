VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmImportColumn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import column"
   ClientHeight    =   6885
   ClientLeft      =   8505
   ClientTop       =   1755
   ClientWidth     =   5520
   Icon            =   "frmImportColumn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   Begin MSComctlLib.ImageList il 
      Left            =   120
      Top             =   6360
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
            Picture         =   "frmImportColumn.frx":08CA
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportColumn.frx":0F9C
            Key             =   "namespace"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportColumn.frx":1B6E
            Key             =   "column"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   4365
      TabIndex        =   0
      ToolTipText     =   "Add the selected column."
      Top             =   6480
      Width           =   1095
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   6360
      Left            =   25
      TabIndex        =   1
      Top             =   25
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   11218
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Database"
      TabPicture(0)   =   "frmImportColumn.frx":2240
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDetail"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.TreeView tv 
         Height          =   4440
         Left            =   120
         TabIndex        =   2
         Top             =   420
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7832
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "il"
         Appearance      =   1
      End
      Begin VB.Label lblDetail 
         BorderStyle     =   1  'Fixed Single
         Height          =   1290
         Left            =   120
         TabIndex        =   3
         Top             =   4965
         Width           =   5175
      End
   End
   Begin VB.Menu mnuModify 
      Caption         =   "Modify"
      Visible         =   0   'False
      Begin VB.Menu mnuModifyCopyVar 
         Caption         =   "Copy Setting Variable"
      End
      Begin VB.Menu mnuModifyPasteVar 
         Caption         =   "Paste Setting Variable"
      End
   End
End
Attribute VB_Name = "frmImportColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmImportColumn.frm - Import column form other table

Option Explicit
Dim frmCallingForm As Form

Public Sub Initialise(frmCF As Form)
If inIDE Then:  On Error GoTo 0: Else: On Error GoTo Err_Handler:
frmMain.svr.LogEvent "Entering " & App.Title & ":frmImportColumn.Initialise()", etFullDebug
  
Dim objNS As pgNamespace
Dim objTable As pgTable
Dim objColumn As pgColumn
Dim objNode As Node
Dim objNode1 As Node

  'Set the font
  PatchForm Me
  
  Set frmCallingForm = frmCF
  
  StartMsg "Load column table..."
  
  'Load the namespace
  tv.Nodes.Clear
  For Each objNS In frmMain.svr.Databases(ctx.CurrentDB).Namespaces
    If Not objNS.SystemObject And objNS.Tables.Count > 0 Then
      Set objNode = tv.Nodes.Add(, , "NSP+" & GetID, objNS.Identifier, "namespace")
    
      'Load the table
      For Each objTable In objNS.Tables
        Set objNode1 = tv.Nodes.Add(objNode.Key, tvwChild, "TBL+" & GetID, objTable.Identifier, "table")
      
        'Load the column
        For Each objColumn In objTable.Columns
          If Not (objColumn.SystemObject And Not ctx.IncludeSys) Then tv.Nodes.Add objNode1.Key, tvwChild, "COL+" & GetID, objColumn.Identifier, "column"
        Next
      Next
    End If
  Next
  
  EndMsg
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmImportColumn.Initialise"
End Sub

Private Sub cmdAdd_Click()
If inIDE Then:  On Error GoTo 0: Else: On Error GoTo Err_Handler:
frmMain.svr.LogEvent "Entering " & App.Title & ":frmImportColumn.cmdAdd_Click()", etFullDebug

Dim objItem As ListItem
Dim objColumn As pgColumn
  
  If Not frmCallingForm Is Nothing Then
    If Not frmCallingForm.Visible Then
      MsgBox "The form that called this form has been destroyed!", vbSystemModal + vbExclamation, "Error"
      Unload Me
      Exit Sub
    End If
  End If

  If tv.SelectedItem Is Nothing Then
    MsgBox "You must select a column to import!", vbExclamation, "Error"
    Exit Sub
  End If
  If Left(tv.SelectedItem.Key, 3) <> "COL" Then
    MsgBox "You must select a column to import!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'get select colum
  Set objColumn = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(tv.SelectedItem.Parent.Parent).Tables(tv.SelectedItem.Parent).Columns(tv.SelectedItem)
  
  For Each objItem In frmCallingForm.lvProperties(0).ListItems
    If objItem.Text = objColumn.Name Then
      MsgBox "A column with that name already exists!", vbExclamation, "Error"
      Exit Sub
    End If
  Next
      
  'add column
  If ((objColumn.DataType = "serial") Or (objColumn.DataType = "serial8")) Then
    Set objItem = frmCallingForm.lvProperties(0).ListItems.Add(, , objColumn.Name, "sequence", "sequence")
  Else
    Set objItem = frmCallingForm.lvProperties(0).ListItems.Add(, , objColumn.Name, "column", "column")
  End If
  objItem.SubItems(1) = objItem.Index
  objItem.SubItems(2) = objColumn.DataType
  If objColumn.DataType = "numeric" Or objColumn.DataType = "char" Or objColumn.DataType = "varchar" Then
    objItem.SubItems(3) = objColumn.Length
  End If
  objItem.SubItems(4) = objColumn.Default
  If objColumn.NotNull Then
    objItem.SubItems(5) = "Yes"
  Else
    objItem.SubItems(5) = "No"
  End If
  If objColumn.PrimaryKey Then
    objItem.SubItems(6) = "Yes"
  Else
    objItem.SubItems(6) = "No"
  End If
  objItem.SubItems(7) = objColumn.Comment
      
  frmCallingForm.lvProperties(0).Tag = "Y"
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmImportColumn.cmdAdd_Click"
End Sub

'Load detail column
Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
If inIDE Then:  On Error GoTo 0: Else: On Error GoTo Err_Handler:
frmMain.svr.LogEvent "Entering " & App.Title & ":frmImportColumn.tv_NodeClick(" & Node.Text & ")", etFullDebug
  
Dim szTemp As String
Dim objColumn As pgColumn
  
  lblDetail.Caption = ""
  If tv.SelectedItem Is Nothing Then Exit Sub
  If Left(tv.SelectedItem.Key, 3) <> "COL" Then Exit Sub
      
  'get select colum
  Set objColumn = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(tv.SelectedItem.Parent.Parent).Tables(tv.SelectedItem.Parent).Columns(tv.SelectedItem)
  
  szTemp = szTemp & "Type : " & objColumn.DataType & vbCrLf
  szTemp = szTemp & "Length : " & objColumn.Length & vbCrLf
  szTemp = szTemp & "Default : " & objColumn.Default & vbCrLf
  szTemp = szTemp & "Not Null : " & objColumn.NotNull & vbCrLf
  szTemp = szTemp & "Primary Key : " & objColumn.PrimaryKey & vbCrLf
  szTemp = szTemp & "Comment : " & objColumn.Comment & vbCrLf
  lblDetail.Caption = szTemp

  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmImportColumn.tv_NodeClick"
End Sub
