VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database"
   ClientHeight    =   6885
   ClientLeft      =   2085
   ClientTop       =   1800
   ClientWidth     =   5520
   Icon            =   "frmDatabase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   5520
   Begin MSComctlLib.ImageList il 
      Left            =   90
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":014A
            Key             =   "encoding"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":0A24
            Key             =   "off"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":0E76
            Key             =   "database"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":0FD0
            Key             =   "public"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":112A
            Key             =   "user"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":1284
            Key             =   "group"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":1956
            Key             =   "property"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":2028
            Key             =   "on"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":247A
            Key             =   "info"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":28CC
            Key             =   "error"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":2D1E
            Key             =   "warning"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":3170
            Key             =   "debug"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabase.frx":370A
            Key             =   "log"
         EndProperty
      EndProperty
   End
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
      TabIndex        =   8
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
      TabPicture(0)   =   "frmDatabase.frx":455C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProperties(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProperties(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperties(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblProperties(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblProperties(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblProperties(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboProperties(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "hbxProperties(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProperties(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProperties(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtProperties(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtProperties(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cboProperties(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "&Variables"
      TabPicture(1)   =   "frmDatabase.frx":4578
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboVarValue"
      Tab(1).Control(1)=   "cmdCurrVal"
      Tab(1).Control(2)=   "cboVarName"
      Tab(1).Control(3)=   "txtVarValue"
      Tab(1).Control(4)=   "cmdAddVar"
      Tab(1).Control(5)=   "cmdRemoveVar"
      Tab(1).Control(6)=   "lvProperties(0)"
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(8)=   "Label1"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "&Security"
      TabPicture(2)   =   "frmDatabase.frx":4594
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvProperties(1)"
      Tab(2).Control(1)=   "fraAdd"
      Tab(2).Control(2)=   "cmdAdd"
      Tab(2).Control(3)=   "cmdRemove"
      Tab(2).ControlCount=   4
      Begin MSComctlLib.ImageCombo cboVarValue 
         Height          =   330
         Left            =   -73425
         TabIndex        =   30
         Top             =   5940
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.CommandButton cmdCurrVal 
         Caption         =   "&Show Current Settings"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71520
         TabIndex        =   29
         ToolTipText     =   "Show the current variables and their values."
         Top             =   4995
         Width           =   1830
      End
      Begin MSComctlLib.ImageCombo cboVarName 
         Height          =   330
         Left            =   -73425
         TabIndex        =   28
         Top             =   5520
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73515
         TabIndex        =   26
         ToolTipText     =   "Remove the selected entry."
         Top             =   3900
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74865
         TabIndex        =   25
         ToolTipText     =   "Add the defined entry."
         Top             =   3900
         Width           =   1230
      End
      Begin VB.Frame fraAdd 
         Caption         =   "Define Privilege"
         Height          =   1815
         Left            =   -74865
         TabIndex        =   20
         Top             =   4380
         Width           =   5190
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Temp"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   22
            ToolTipText     =   "Give temp privilege to the selected entity."
            Top             =   1350
            Width           =   1590
         End
         Begin VB.CheckBox chkPrivilege 
            Caption         =   "&Create"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   21
            ToolTipText     =   "Give create privilege to the selected entity."
            Top             =   945
            Width           =   1590
         End
         Begin MSComctlLib.ImageCombo cboEntities 
            Height          =   330
            Left            =   1260
            TabIndex        =   23
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
            Index           =   5
            Left            =   180
            TabIndex        =   24
            Top             =   360
            Width           =   840
         End
      End
      Begin VB.TextBox txtVarValue 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73425
         TabIndex        =   12
         ToolTipText     =   "Enter the name of the variable to add or update."
         Top             =   5940
         Width           =   3750
      End
      Begin VB.CommandButton cmdAddVar 
         Caption         =   "&Add/Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74820
         TabIndex        =   10
         ToolTipText     =   "Add (or update) the specified variable."
         Top             =   4995
         Width           =   1230
      End
      Begin VB.CommandButton cmdRemoveVar 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73515
         TabIndex        =   11
         ToolTipText     =   "Remove the selected variable."
         Top             =   4995
         Width           =   1230
      End
      Begin MSComctlLib.ImageCombo cboProperties 
         Height          =   330
         Index           =   0
         Left            =   1935
         TabIndex        =   4
         ToolTipText     =   "Select or enter the encoding scheme to use."
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
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "An alternate filesystem location in which to store the new database, specified as a string literal."
         Top             =   2295
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "The databases owner."
         Top             =   1485
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "The databases OID (Object ID) in PostgreSQL."
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtProperties 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "The name of the database."
         Top             =   675
         Width           =   3390
      End
      Begin HighlightBox.HBX hbxProperties 
         Height          =   2940
         Index           =   0
         Left            =   135
         TabIndex        =   6
         ToolTipText     =   "Comments about the database."
         Top             =   3060
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   5186
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
      Begin MSComctlLib.ListView lvProperties 
         Height          =   4425
         Index           =   0
         Left            =   -74865
         TabIndex        =   9
         ToolTipText     =   "Lists the configuration variables set for this database."
         Top             =   450
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   7805
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
            Text            =   "Variable"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView lvProperties 
         Height          =   3390
         Index           =   1
         Left            =   -74865
         TabIndex        =   27
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
         Index           =   1
         Left            =   1935
         TabIndex        =   31
         ToolTipText     =   "Select or enter the encoding scheme to use."
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
         Caption         =   "Template"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   2685
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Variable Value"
         Height          =   195
         Left            =   -74820
         TabIndex        =   19
         Top             =   5985
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Variable Name"
         Height          =   195
         Left            =   -74820
         TabIndex        =   18
         Top             =   5580
         Width           =   1140
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "OID"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Owner"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   1530
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Encoding"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   14
         Top             =   1935
         Width           =   675
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   2340
         Width           =   330
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
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmDatabase.frm - Edit/Create a Database

Option Explicit

Dim bNew As Boolean
Dim objDatabase As pgDatabase
Dim szUsers() As String
Dim szVarDropList As String
Const PrefKey = "KEY_"

Private Sub cboVarName_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cboVarName_Click()", etFullDebug

Dim objVardb As VarDb
Dim vData
Dim szImg As String

  txtVarValue.Visible = False
  txtVarValue.Text = ""
  cboVarValue.Visible = False
  cboVarValue.Text = ""
  cboVarValue.ComboItems.Clear
  cboVarValue.Locked = True
  
  objVardb = GetVarDb(cboVarName.Text)
  
  If objVardb.Type = TVDB_BOOLEAN Then
    szImg = GetImageFromVal("ON", TVDB_BOOLEAN)
    cboVarValue.ComboItems.Add , PrefKey & "on", "ON", szImg, szImg
    szImg = GetImageFromVal("OFF", TVDB_BOOLEAN)
    cboVarValue.ComboItems.Add , PrefKey & "off", "OFF", szImg, szImg
    cboVarValue.ComboItems(1).Selected = True
    cboVarValue.Locked = True
    cboVarValue.Visible = True
  ElseIf objVardb.Type = TVDB_FLOAT Or objVardb.Type = TVDB_INTEGR Or objVardb.Type = TVDB_STRING Then
    txtVarValue.Visible = True
  ElseIf objVardb.Type = TVDB_CAST Then
    For Each vData In objVardb.CastValue
      szImg = GetImageFromVal(CStr(vData), TVDB_CAST)
      cboVarValue.ComboItems.Add , PrefKey & LCase(vData), vData, szImg, szImg
    Next
    cboVarValue.ComboItems(1).Selected = True
    cboVarValue.Locked = False
    cboVarValue.Visible = True
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdCancel_Click"
End Sub

Private Sub cmdCancel_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdCancel_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdCancel_Click"
End Sub

Private Sub cmdCurrVal_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.CmdCurrVal_Click()", etFullDebug

Dim rsQuery As New Recordset
Dim objOutputForm As New frmSQLOutput
  
  Set rsQuery = frmMain.svr.Databases(objDatabase.Name).Execute("SELECT name AS " & QUOTE & "Variable Name" & QUOTE & ", setting AS " & QUOTE & "Current Value" & QUOTE & " FROM pg_settings ORDER BY name")
  Load objOutputForm
  objOutputForm.Display rsQuery, objDatabase.Name, Me.Tag
  objOutputForm.Show

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.CmdCurrVal_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdOK_Click()", etFullDebug

Dim objNode As Node
Dim objItem As ListItem
Dim objNewDatabase As pgDatabase
Dim szDropVars() As String
Dim X As Integer
Dim lACL As Long
Dim szEntity As String
Dim vEntity As Variant

  'Check the data
  If txtProperties(0).Text = "" Then
    MsgBox "You must specify a database name!", vbExclamation, "Error"
    txtProperties(0).SetFocus
    Exit Sub
  End If
  
  If bNew Then
    StartMsg "Creating Database..."
    Set objNewDatabase = frmMain.svr.Databases.Add(txtProperties(0).Text, cboProperties(1).Text, _
                                                   txtProperties(3).Text, cboProperties(0).Text, _
                                                   hbxProperties(0).Text)
    
    'Add a new node and update the text on the parent
    Set objNode = frmMain.svr.Databases.Tag
    Set objNewDatabase.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "DAT-" & GetID, txtProperties(0).Text, "database")
    objNode.Text = "Databases (" & objNode.Children & ")"
    
  Else
    StartMsg "Updating Database..."
    
    'Add any vars
    If lvProperties(0).Tag = "Y" Then
      For Each objItem In lvProperties(0).ListItems
        objDatabase.DatabaseVars.AddOrUpdate objItem.Text, objItem.SubItems(1)
      Next objItem
    End If
    
    'Drop any vars
    If Len(szVarDropList) > 3 Then
      szDropVars = Split(szVarDropList, "!|!")
      For X = 0 To UBound(szDropVars)
        If szDropVars(X) <> "" Then objDatabase.DatabaseVars.Remove szDropVars(X)
      Next X
    End If
    
    If hbxProperties(0).Tag = "Y" Then objDatabase.Comment = hbxProperties(0).Text
  End If
  
  'Set the ACL on the Database as required
  If lvProperties(1).Tag = "Y" Then
    'Revoke all from existing entries
    For Each vEntity In szUsers
      If vEntity <> "" Then
        If vEntity = "PUBLIC" Then
          frmMain.svr.Databases(txtProperties(0).Text).Revoke vEntity, aclAll
        ElseIf Left(vEntity, 6) = "GROUP " Then
          frmMain.svr.Databases(txtProperties(0).Text).Revoke "GROUP " & fmtID(Mid(vEntity, 7)), aclAll
        Else
          frmMain.svr.Databases(txtProperties(0).Text).Revoke fmtID(vEntity), aclAll
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
      If InStr(1, objItem.SubItems(1), "Create") <> 0 Then lACL = lACL + aclCreate
      If InStr(1, objItem.SubItems(1), "Temp") <> 0 Then lACL = lACL + aclTemp
      frmMain.svr.Databases(txtProperties(0).Text).Grant szEntity, lACL
    Next objItem
  End If
  
  'Simulate a node click to refresh the ListView
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
    
  EndMsg
  Unload Me
  Exit Sub
  
Err_Handler:
  If Err.Number = 35606 Then Resume Next
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdOK_Click"
  
  'If we error here, refresh the database vars to ensure they are in a consistant state
  If Not (objDatabase Is Nothing) Then
    objDatabase.DatabaseVars.Refresh
    LoadVars
  End If
End Sub

Public Sub Initialise(Optional Database As pgDatabase)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.Initialise()", etFullDebug

Dim X As Integer
Dim objItem As ComboItem
Dim objLItem As ListItem
Dim objUser As pgUser
Dim objGroup As pgGroup
Dim szUserlist As String
Dim szAccesslist As String
Dim szAccess() As String
Dim rsVar As Recordset
Dim objDb As pgDatabase
  
  PatchForm Me
  
  'Unlock the edittable fields
  If ctx.dbVer >= 7.3 Then
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    lvProperties(1).BackColor = &H80000005
    cboEntities.BackColor = &H80000005
    chkPrivilege(0).Enabled = True
    chkPrivilege(1).Enabled = True
    cmdCurrVal.Enabled = True
  End If
  
  If Database Is Nothing Then
  
    'Create a new database
    bNew = True
    Me.Caption = "Create Database"
    
    'Load the Encoding Schemes
    cboProperties(0).Text = "SQL_ASCII"
    Set objItem = cboProperties(0).ComboItems.Add(, , "SQL_ASCII", "encoding", "encoding")
    objItem.Selected = True
    cboProperties(0).ComboItems.Add , , "EUC_JP", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_CN", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_KR", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "EUC_TW", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "UNICODE", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "MULE_INTERNAL", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN1", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN2", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN3", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN4", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "LATIN5", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "KOI8", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "WIN", "encoding", "encoding"
    cboProperties(0).ComboItems.Add , , "ALT", "encoding", "encoding"
   
    'Unlock the edittable fields
    txtProperties(0).BackColor = &H80000005
    txtProperties(0).Locked = False
    cboProperties(0).BackColor = &H80000005
    cboProperties(0).Locked = False
    cboProperties(1).BackColor = &H80000005
    cboProperties(1).Locked = False
    txtProperties(3).BackColor = &H80000005
    txtProperties(3).Locked = False
    hbxProperties(0).BackColor = &H80000005
    hbxProperties(0).Locked = False
    
    'Redim the userlist so it doesn't cause an error later.
    ReDim szUsers(0)

    'Load the Template Database
    For Each objDb In frmMain.svr.Databases
      cboProperties(1).ComboItems.Add , objDb.Name, objDb.Name, "database", "database"
    Next
    cboProperties(1).ComboItems("template0").Selected = True
  Else
    
    'Display/Edit the specified Database.
    Set objDatabase = Database
    bNew = False
    Me.Caption = "Database: " & objDatabase.Identifier
    
    'Unlock the Vars. We only edit these for existing objects as there is no
    ' safe way to create the object & update the vars in one 'transaction'
    If ctx.dbVer >= 7.3 Then
      lvProperties(0).BackColor = &H80000005
      txtVarValue.Enabled = True
      txtVarValue.BackColor = &H80000005
      cmdAddVar.Enabled = True
      cmdRemoveVar.Enabled = True
      cboVarName.Enabled = True
      cboVarName.BackColor = &H80000005
      cboVarValue.Enabled = True
      cboVarValue.BackColor = &H80000005
      
      'load var name
      cboVarName.ComboItems.Clear
      Set rsVar = frmMain.svr.Databases(frmMain.svr.MasterDB).Execute("SELECT name FROM pg_settings ORDER BY name")
      While Not rsVar.EOF
        cboVarName.ComboItems.Add , LCase(rsVar("name")), rsVar("name"), "property"
        rsVar.MoveNext
      Wend
      cboVarName.ComboItems(1).Selected = True
      cboVarName_Click
    End If
    
    If objDatabase.SystemObject Then  'Lock the permissions Add/Remove buttons if it's a system object
      cmdAdd.Enabled = False
      cmdRemove.Enabled = False
    End If
    
    If objDatabase.Status <> statInaccessible Then
      hbxProperties(0).BackColor = &H80000005
      hbxProperties(0).Locked = False
    End If
    txtProperties(0).Text = objDatabase.Name
    txtProperties(1).Text = objDatabase.Oid
    txtProperties(2).Text = objDatabase.Owner
    Set objItem = cboProperties(0).ComboItems.Add(, , objDatabase.ServerEncoding, "encoding", "encoding")
    objItem.Selected = True
    txtProperties(3).Text = objDatabase.Path
    hbxProperties(0).Text = objDatabase.Comment
    
    LoadVars
    
    ParseACL objDatabase.ACL, szUserlist, szAccesslist
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
  For X = 0 To 3
    txtProperties(X).Tag = "N"
  Next X
  hbxProperties(0).Tag = "N"
  lvProperties(0).Tag = "N"
  lvProperties(1).Tag = "N"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.Initialise"
End Sub

Private Sub cmdRemove_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdRemove_Click()", etFullDebug

  If lvProperties(1).SelectedItem Is Nothing Then Exit Sub
  lvProperties(1).ListItems.Remove lvProperties(1).SelectedItem.Index
  lvProperties(1).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdRemove_Click"
End Sub

Private Sub cmdAdd_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdAdd_Click()", etFullDebug

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
  If chkPrivilege(0).Value = 1 Then szAccess = szAccess & "Create, "
  If chkPrivilege(1).Value = 1 Then szAccess = szAccess & "Temp, "
  If Len(szAccess) > 2 Then szAccess = Left(szAccess, Len(szAccess) - 2)
  If szAccess = "" Then szAccess = "None"
  
  Set objItem = lvProperties(1).ListItems.Add(, , cboEntities.SelectedItem.Text, cboEntities.SelectedItem.Image, cboEntities.SelectedItem.Image)
  objItem.SubItems(1) = szAccess
  lvProperties(1).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdAdd_Click"
End Sub

Private Sub LoadVars()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.LoadVars()", etFullDebug

Dim objItem As ListItem
Dim objVar As pgVar
Dim objVardb As VarDb
Dim szImg As String

  lvProperties(0).ListItems.Clear
  If ctx.dbVer >= 7.3 Then
    If Not (objDatabase.DatabaseVars Is Nothing) Then
      For Each objVar In objDatabase.DatabaseVars
        Set objItem = lvProperties(0).ListItems.Add(, , objVar.Name)
        objItem.SubItems(1) = objVar.Value
        
        'get image
        szImg = "property"    'image default
        objVardb = GetVarDb(objVar.Name)
        If objVardb.Type = TVDB_BOOLEAN Or objVardb.Type = TVDB_CAST Then
          szImg = GetImageFromVal(objVar.Value, objVardb.Type)
        End If
        objItem.Icon = szImg
        objItem.SmallIcon = szImg
      Next objVar
    Else
      cmdRemoveVar.Enabled = False
      cmdAddVar.Enabled = False
      cmdCurrVal.Enabled = False
    End If
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.LoadVars"
End Sub

Private Sub cmdRemoveVar_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdRemoveVar_Click()", etFullDebug

  If lvProperties(0).SelectedItem Is Nothing Then
    MsgBox "You must select a variable to remove!", vbExclamation, "Error"
    tabProperties.Tab = 1
    lvProperties(0).SetFocus
    Exit Sub
  End If
  
  If objDatabase Is Nothing Then
    lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
    lvProperties(0).Tag = "Y"
    If lvProperties(0).SelectedItem Is Nothing Then
      cmdRemoveVar.Enabled = False
    End If
  Else
    szVarDropList = szVarDropList & lvProperties(0).SelectedItem.Text & "!|!"
    lvProperties(0).ListItems.Remove lvProperties(0).SelectedItem.Index
    lvProperties(0).Tag = "Y"
    If lvProperties(0).SelectedItem Is Nothing Then
      cmdRemoveVar.Enabled = False
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdRemoveVar_Click"
End Sub

Private Sub cmdAddVar_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.cmdChkAdd_Click()", etFullDebug

Dim objItem As ListItem
Dim szVal As String
Dim szImg As String

  'get value variable
  If txtVarValue.Visible = True Then
    szVal = txtVarValue.Text
  ElseIf cboVarValue.Visible = True Then
    szVal = cboVarValue.Text
  End If

  If Trim(szVal) = "" Then
    MsgBox "You must enter a value for the variable!", vbExclamation, "Error"
    tabProperties.Tab = 1
    txtVarValue.SetFocus
    Exit Sub
  End If
 
  'image default
  szImg = "property"
  
  'Update
  For Each objItem In lvProperties(0).ListItems
    If objItem.Text = cboVarName.SelectedItem.Text Then
      objItem.SubItems(1) = szVal
      lvProperties(0).Tag = "Y"
      
      If cboVarValue.Visible And Not cboVarValue.SelectedItem Is Nothing Then
        If Len(cboVarValue.SelectedItem.Image) > 0 Then
          szImg = cboVarValue.SelectedItem.Image
        End If
      End If
      objItem.Icon = szImg
      objItem.SmallIcon = szImg
      
      cboVarName.ComboItems(1).Selected = True
      cboVarName_Click
      Exit Sub
    End If
  Next objItem
  
  'Or add
  Set objItem = lvProperties(0).ListItems.Add(, , cboVarName.SelectedItem.Text)
  objItem.SubItems(1) = szVal
  lvProperties(0).Tag = "Y"
  
  If cboVarValue.Visible And Not cboVarValue.SelectedItem Is Nothing Then
    If Len(cboVarValue.SelectedItem.Image) > 0 Then
      szImg = cboVarValue.SelectedItem.Image
    End If
  End If
  objItem.Icon = szImg
  objItem.SmallIcon = szImg
  
  cboVarName.ComboItems(1).Selected = True
  cboVarName_Click
  
  Exit Sub
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.cmdAddVar_Click"
End Sub

Private Sub hbxProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.hbxProperties_Change(" & Index & ")", etFullDebug

  hbxProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.hbxProperties_Change"
End Sub

Private Sub lvProperties_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.lvProperties_Click(" & Index & ")", etFullDebug

Dim objVardb As VarDb
Dim vData
Dim szVal As String

  If Index = 0 Then
    If Not (lvProperties(0).SelectedItem Is Nothing) Then
      cboVarName.ComboItems(LCase(lvProperties(0).SelectedItem.Text)).Selected = True
      cboVarName_Click
      
      If txtVarValue.Visible = True Then
        txtVarValue.Text = lvProperties(0).SelectedItem.SubItems(1)
      ElseIf cboVarValue.Visible = True Then
        objVardb = GetVarDb(cboVarName.Text)
        If objVardb.Type = TVDB_BOOLEAN Then
          Select Case UCase(lvProperties(0).SelectedItem.SubItems(1))
            Case "ON", "TRUE", "YES", "1"
              szVal = "ON"
            Case "OFF", "FALSE", "NO", "0"
              szVal = "OFF"
            Case Else
              szVal = "OFF"
          End Select
          cboVarValue.ComboItems(PrefKey & LCase(szVal)).Selected = True
        Else
          szVal = lvProperties(0).SelectedItem.SubItems(1)
        End If
        
        'find value in combo
        For Each vData In objVardb.CastValue
          If UCase(szVal) = UCase(vData) Then
            cboVarValue.ComboItems("KEY_" & LCase(szVal)).Selected = True
            Exit Sub
          End If
        Next
        
        'cast value not present
        'manual insert
        If objVardb.Type = TVDB_CAST Then cboVarValue.Text = szVal
      End If
    End If
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.lvProperties_Click"
End Sub

Private Sub txtProperties_Change(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.txtProperties_Change(" & Index & ")", etFullDebug

  txtProperties(Index).Tag = "Y"
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.txtProperties_Change"
End Sub

Private Sub lvProperties_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.lvProperties_MouseDown(" & Index & "," & Button & "," & Shift & "," & X & "," & y & ")", etFullDebug

  If Button = vbRightButton Then
    mnuModifyPasteVar.Enabled = False
    If ColVarDbBuffer.Count > 0 Then mnuModifyPasteVar.Enabled = True
    PopupMenu mnuModify
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.lvProperties_MouseDown"
End Sub

'copy var setting database
Private Sub mnuModifyCopyVar_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.mnuModifyCopyVar_Click()", etFullDebug

Dim objLv As ListItem

  Set ColVarDbBuffer = New Collection
  For Each objLv In lvProperties(0).ListItems
    ColVarDbBuffer.Add objLv.Text & "|" & objLv.SubItems(1)
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.mnuModifyCopyVar_Click"
End Sub

'paste var setting database
Private Sub mnuModifyPasteVar_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmDatabase.mnuModifyPasteVar_Click()", etFullDebug

Dim vData

  'simulate add/update var
  For Each vData In ColVarDbBuffer
    vData = Split(vData, "|")
    'select variable name
    cboVarName.ComboItems(vData(0)).Selected = True
    cboVarName_Click
    
    'set value
    If cboVarValue.Visible Then
      cboVarValue.ComboItems(PrefKey & LCase(vData(1))).Selected = True
    Else
      txtVarValue.Text = vData(1)
    End If
    
    'add var
    cmdAddVar_Click
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmDatabase.mnuModifyPasteVar_Click"
End Sub

