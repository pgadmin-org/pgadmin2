VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASCII Data Export"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   6694
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Options"
      TabPicture(0)   =   "frmExport.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblfileName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraDelimiter"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraQuoting"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBrowse"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFileName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Substitution Map"
      TabPicture(1)   =   "frmExport.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDelete"
      Tab(1).Control(1)=   "cmdAdd"
      Tab(1).Control(2)=   "txtReplace"
      Tab(1).Control(3)=   "txtSearch"
      Tab(1).Control(4)=   "lvSubMap"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Substitution"
         Height          =   330
         Left            =   -72210
         TabIndex        =   21
         Top             =   3330
         Width           =   1680
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Substitution"
         Height          =   330
         Left            =   -74595
         TabIndex        =   20
         Top             =   3330
         Width           =   1680
      End
      Begin VB.TextBox txtReplace 
         Height          =   285
         Left            =   -72525
         TabIndex        =   19
         ToolTipText     =   "Enter a string to replace with."
         Top             =   2970
         Width           =   2265
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   -74910
         TabIndex        =   18
         ToolTipText     =   "Enter a string to search for."
         Top             =   2970
         Width           =   2310
      End
      Begin MSComctlLib.ListView lvSubMap 
         Height          =   2535
         Left            =   -74955
         TabIndex        =   17
         ToolTipText     =   "Lists text substitutions that will be made to the data as it is exported."
         Top             =   360
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   990
         TabIndex        =   15
         Top             =   450
         Width           =   3345
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   4365
         TabIndex        =   14
         Top             =   450
         Width           =   330
      End
      Begin VB.Frame fraQuoting 
         Caption         =   "Quoting"
         Height          =   1215
         Left            =   1035
         TabIndex        =   8
         Top             =   2340
         Width           =   2970
         Begin VB.OptionButton optQuote 
            Alignment       =   1  'Right Justify
            Caption         =   "None"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   13
            ToolTipText     =   "Specify not to enclose values between 2 characters"
            Top             =   270
            Width           =   1455
         End
         Begin VB.OptionButton optQuote 
            Alignment       =   1  'Right Justify
            Caption         =   "Character"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   12
            ToolTipText     =   "Specify a character to 'quote' each column with "
            Top             =   585
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optQuote 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascii Value"
            Height          =   225
            Index           =   2
            Left            =   105
            TabIndex        =   11
            ToolTipText     =   "Specify a decimal Ascii value to 'quote' each column with"
            Top             =   900
            Width           =   1455
         End
         Begin VB.TextBox txtQuoteChar 
            Height          =   285
            Left            =   2205
            TabIndex        =   10
            Text            =   "'"
            ToolTipText     =   "Enter a character to use as a quote mark"
            Top             =   525
            Width           =   645
         End
         Begin VB.TextBox txtQuoteAscii 
            Height          =   285
            Left            =   2205
            TabIndex        =   9
            ToolTipText     =   "Enter a decimal Ascii value to use as a quote mark"
            Top             =   855
            Width           =   645
         End
      End
      Begin VB.Frame fraDelimiter 
         Caption         =   "Delimiter"
         Height          =   1215
         Left            =   1035
         TabIndex        =   2
         Top             =   945
         Width           =   2970
         Begin VB.CheckBox chkTrailing 
            Alignment       =   1  'Right Justify
            Caption         =   "Add trailing delimiter?"
            Height          =   225
            Left            =   105
            TabIndex        =   7
            ToolTipText     =   "Specify whether or not to include a delimiter after the last column (Recommended)"
            Top             =   900
            Value           =   1  'Checked
            Width           =   2715
         End
         Begin VB.OptionButton optDelimiter 
            Alignment       =   1  'Right Justify
            Caption         =   "Character"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   6
            ToolTipText     =   "Specify a character as a delimiter"
            Top             =   270
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optDelimiter 
            Alignment       =   1  'Right Justify
            Caption         =   "Ascii Value"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   5
            ToolTipText     =   "Specify an Ascii value to use as a delimiter"
            Top             =   585
            Width           =   1275
         End
         Begin VB.TextBox txtDelimChar 
            Height          =   285
            Left            =   2205
            TabIndex        =   4
            Text            =   ","
            ToolTipText     =   "Enter the character to use as delimiter"
            Top             =   225
            Width           =   645
         End
         Begin VB.TextBox txtDelimAscii 
            Height          =   285
            Left            =   2205
            TabIndex        =   3
            ToolTipText     =   "Enter the decimal Ascii value to use as the delimiter"
            Top             =   540
            Width           =   645
         End
      End
      Begin VB.Label lblfileName 
         AutoSize        =   -1  'True
         Caption         =   "Export File"
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   495
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   330
      Left            =   3735
      TabIndex        =   0
      Top             =   3915
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   3690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit
Public szDelimiter As String
Public szQuote As String
Public bTrailing As Boolean
Public szFilename As String

Private Sub cmdAdd_Click()
Dim itmX As ListItem
  If txtSearch.Text = "" Then
    MsgBox "You must enter a string to search for!", vbExclamation, "Error"
    txtSearch.SetFocus
    Exit Sub
  End If
  If txtReplace.Text = "" Then
    MsgBox "You must enter a string to replace with!", vbExclamation, "Error"
    txtReplace.SetFocus
    Exit Sub
  End If
  Set itmX = lvSubMap.ListItems.Add(, txtSearch.Text, txtSearch.Text)
  itmX.SubItems(1) = txtReplace.Text
  txtSearch.Text = ""
  txtReplace.Text = ""
End Sub

Private Sub cmdBrowse_Click()
  With CommonDialog1
    .FileName = txtFileName.Text
    .DialogTitle = "Save ASCII Text File"
    .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowSave
  End With
  txtFileName.Text = CommonDialog1.FileName
End Sub

Private Sub cmdDelete_Click()
  If MsgBox("Are you sure you want to remove the selected substitution?", vbQuestion + vbYesNo, "Delete Substitution?") = vbNo Then Exit Sub
  lvSubMap.ListItems.Remove lvSubMap.SelectedItem.Key
End Sub

Private Sub cmdExport_Click()
Dim X As Integer
  If optQuote(0).Value = True Then
    szQuote = ""
  ElseIf optQuote(1).Value = True Then
    szQuote = txtQuoteChar.Text
    If szQuote = "" Then
      MsgBox "You must specify a quoting character!", vbExclamation, "Error"
      Exit Sub
    End If
  ElseIf optQuote(2).Value = True Then
    szQuote = Chr(txtQuoteAscii.Text)
    If szQuote = "" Then
      MsgBox "You must specify a quoting character!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  If optDelimiter(0).Value = True Then
    szDelimiter = txtDelimChar.Text
    If szDelimiter = "" Then
      MsgBox "You must specify a delimiting character!", vbExclamation, "Error"
      Exit Sub
    End If
  ElseIf optDelimiter(1).Value = True Then
    szDelimiter = Chr(txtDelimAscii.Text)
    If szDelimiter = "" Then
      MsgBox "You must specify a delimiting character!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  If chkTrailing.Value = 1 Then
    bTrailing = True
  Else
    bTrailing = False
  End If
  If txtFileName.Text = "" Then
    MsgBox "You must specify a filename!", vbExclamation, "Error"
    Exit Sub
  End If
  If Dir(txtFileName.Text) <> "" Then
    If MsgBox("File exists - overwrite?", vbYesNo + vbQuestion, "Overwrite File") = vbNo Then Exit Sub
  End If
  szFilename = txtFileName.Text
  If optDelimiter(0).Value = True Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter Type", regString, "0"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter Type", regString, "1"
  End If
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter Character", regString, txtDelimChar.Text
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter ASCII", regString, txtDelimAscii.Text
  If chkTrailing.Value = 0 Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Trailing Delimiter", regString, "0"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Trailing Delimiter", regString, "1"
  End If
  If optQuote(0).Value = True Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Type", regString, "0"
  ElseIf optQuote(1).Value = True Then
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Type", regString, "1"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Type", regString, "2"
  End If
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Character", regString, txtQuoteChar.Text
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote ASCII", regString, txtQuoteAscii.Text
  RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Entries", regString, lvSubMap.ListItems.Count
  If lvSubMap.ListItems.Count > 0 Then
    For X = 1 To lvSubMap.ListItems.Count
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Search - " & X, regString, lvSubMap.ListItems(X).Text
      RegWrite HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Replace - " & X, regString, lvSubMap.ListItems(X).SubItems(1)
    Next
  End If
  Me.Hide
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim X As Integer
Dim itmX As ListItem
  If RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter Type", "0") = "0" Then
    optDelimiter(0).Value = True
  Else
    optDelimiter(1).Value = True
  End If
  txtDelimChar.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter Character", ",")
  txtDelimAscii.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Delimiter ASCII", "44")
  If RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Trailing Delimiter", "0") = "0" Then
    chkTrailing.Value = 0
  Else
    chkTrailing.Value = 1
  End If
  If RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Type", "0") = "0" Then
    optQuote(0).Value = True
  ElseIf RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Type", "0") = "1" Then
    optQuote(1).Value = True
  Else
    optQuote(2).Value = True
  End If
  txtQuoteChar.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote Character", Chr(34))
  txtQuoteAscii.Text = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter", "Quote ASCII", "34")
  lvSubMap.ListItems.Clear
  lvSubMap.ColumnHeaders.Add , , "Search for:", lvSubMap.Width / 2
  lvSubMap.ColumnHeaders.Add , , "Replace with:", lvSubMap.Width / 2
  If Val(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Entries", "0")) > 0 Then
    For X = 1 To Val(RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Entries", "0"))
      Set itmX = lvSubMap.ListItems.Add(, RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Search - " & X, ""), RegRead(HKEY_CURRENT_USER, "Software\pgAdmin\ASCII Exporter\Substitution Map", "Search - " & X, ""))
      itmX.SubItems(1) = RegRead(HKEY_CURRENT_USER, "Software\pgAdmin II\ASCII Exporter\Substitution Map", "Replace - " & X, "")
    Next
  End If
  
End Sub

