VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration file pg_hba.conf"
   ClientHeight    =   6540
   ClientLeft      =   660
   ClientTop       =   840
   ClientWidth     =   10875
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10875
   Begin MSComDlg.CommonDialog cdlfrmWizard 
      Left            =   720
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList il 
      Left            =   120
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0BC2
            Key             =   "property"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1294
            Key             =   "database"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":13EE
            Key             =   "user"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1548
            Key             =   "group"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1C1A
            Key             =   "all"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1D74
            Key             =   "key"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAddUpd 
      Height          =   3495
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   10695
      Begin VB.Frame fraDBadd 
         Caption         =   "Manually add a Database"
         Height          =   945
         Left            =   120
         TabIndex        =   35
         Top             =   2400
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtDatabase 
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton cmdManualAddDB 
            Caption         =   "Add"
            Height          =   285
            Left            =   2040
            TabIndex        =   36
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   8400
         TabIndex        =   34
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   9525
         TabIndex        =   33
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Frame fraUG 
         Caption         =   "Manually add a User or Group"
         Height          =   945
         Left            =   3120
         TabIndex        =   28
         Top             =   2400
         Visible         =   0   'False
         Width           =   3615
         Begin VB.OptionButton optUG 
            Caption         =   "Group"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   32
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdManualUsrGroupAdd 
            Caption         =   "Add"
            Height          =   285
            Left            =   2040
            TabIndex        =   31
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtUG 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2655
         End
         Begin VB.OptionButton optUG 
            Caption         =   "User"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CheckBox ChkBckConf 
         Caption         =   "Create backup file before save"
         Height          =   375
         Left            =   4800
         TabIndex        =   27
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Update"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1230
      End
      Begin VB.TextBox txtFileUG 
         Height          =   285
         Left            =   3120
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtFileDatabase 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   1230
      End
      Begin VB.TextBox txtIpAddress 
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   4
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpAddress 
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   5
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpAddress 
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   6
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpAddress 
         Height          =   285
         Index           =   3
         Left            =   6240
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpMask 
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpMask 
         Height          =   285
         Index           =   1
         Left            =   7320
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpMask 
         Height          =   285
         Index           =   2
         Left            =   7800
         TabIndex        =   10
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtIpMask 
         Height          =   285
         Index           =   3
         Left            =   8280
         TabIndex        =   11
         Top             =   960
         Width           =   495
      End
      Begin MSComctlLib.ImageCombo cboType 
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ListView lvDatabase 
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   1931
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvUG 
         Height          =   1095
         Left            =   3120
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1931
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageCombo cboAut 
         Height          =   330
         Left            =   8880
         TabIndex        =   12
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboDatabase 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin MSComctlLib.ImageCombo cboUG 
         Height          =   330
         Left            =   3120
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.Label lbDes 
         Caption         =   "Type"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lbDes 
         Caption         =   "Database"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   20
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lbDes 
         Caption         =   "User/Group"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   19
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lbDes 
         Caption         =   "IP-Address"
         Height          =   195
         Index           =   3
         Left            =   4800
         TabIndex        =   18
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lbDes 
         Caption         =   "IP-Mask"
         Height          =   195
         Index           =   4
         Left            =   6840
         TabIndex        =   17
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lbDes 
         Caption         =   "Authentication Method"
         Height          =   195
         Index           =   5
         Left            =   8880
         TabIndex        =   16
         Top             =   720
         Width           =   1620
      End
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "il"
      SmallIcons      =   "il"
      ColHdrIcons     =   "il"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Database"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "User/Group"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IP-Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IP-Mask"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Authentication Method"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmWizard.frm - Configuration file access

Option Explicit

Dim szPath As String
Dim szConfFileName As String    'Config Filename
Dim szConfFilePath As String    'Path to prepend to config Filename

Private Sub cmdAdd_Click(Index As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdAdd_Click()", etFullDebug

Dim lvItem As ListItem
Dim szDatabase As String
Dim szUG As String
Dim szIpAdrress As String
Dim szIpMask As String
Dim ii As Integer
  
  If Index = 1 And lv.SelectedItem Is Nothing Then
    MsgBox "You must select a line to Update!", vbSystemModal + vbExclamation, "Error"
    Exit Sub
  End If
  
  'Database
  If cboDatabase.SelectedItem.Key = "all" Then
    szDatabase = "all"
  ElseIf cboDatabase.SelectedItem.Key = "file" Then
    szDatabase = "@" & txtFileDatabase.Text
  ElseIf cboDatabase.SelectedItem.Key = "database" Then
    szDatabase = ""
    For Each lvItem In lvDatabase.ListItems
      If lvItem.Checked Then szDatabase = szDatabase & lvItem.Text & ","
    Next
    szDatabase = Trim(szDatabase)
    If Len(szDatabase) = 0 Then
      MsgBox "Select a database !!", vbSystemModal + vbExclamation
      Exit Sub
    End If
    szDatabase = Mid(szDatabase, 1, Len(szDatabase) - 1)
  End If
  
  'user/group
  If cboUG.SelectedItem.Key = "all" Then
    szUG = "all"
  ElseIf cboUG.SelectedItem.Key = "file" Then
    szUG = "@" & txtFileUG.Text
  ElseIf cboUG.SelectedItem.Key = "ug" Then
    szUG = ""
    For Each lvItem In lvUG.ListItems
      If lvItem.Checked Then
        If Left(lvItem.Key, 3) = "USR" Then
          szUG = szUG & lvItem.Text & ","
        Else
          szUG = szUG & "+" & lvItem.Text & ","
        End If
      End If
    Next
    szUG = Trim(szUG)
    If Len(szUG) = 0 Then
      MsgBox "Select a user/group !!", vbSystemModal + vbExclamation
      Exit Sub
    End If
    szUG = Mid(szUG, 1, Len(szUG) - 1)
  End If
  
  'IpAdrress
  ii = Int(Len(txtIpAddress(0).Text) > 0) + Int(Len(txtIpAddress(1).Text) > 0) + _
       Int(Len(txtIpAddress(2).Text) > 0) + Int(Len(txtIpAddress(3).Text) > 0)
  If ii = -4 Then
    'verify if number
    szIpAdrress = ""
    For ii = 0 To 3
      If IsNumeric(txtIpAddress(ii).Text) Then
        If txtIpAddress(ii).Text < 0 Or txtIpAddress(ii).Text > 255 Then
          MsgBox "Error IpAddress not valid (range 0 to 255) !!", vbSystemModal + vbExclamation
          txtIpAddress(ii).SetFocus
          Exit Sub
        End If
        szIpAdrress = szIpAdrress & txtIpAddress(ii).Text & "."
      Else
        MsgBox "Error IpAddress not valid !!", vbSystemModal + vbExclamation
        txtIpAddress(ii).SetFocus
        Exit Sub
      End If
    Next
    szIpAdrress = Trim(szIpAdrress)
    szIpAdrress = Mid(szIpAdrress, 1, Len(szIpAdrress) - 1)
  ElseIf ii = 0 Then
    szIpAdrress = ""
  Else
    MsgBox "Error IpAddress not valid !!", vbSystemModal + vbExclamation
    txtIpAddress(0).SetFocus
    Exit Sub
  End If
  
  'Ipmask
  ii = Int(Len(txtIpMask(0).Text) > 0) + Int(Len(txtIpMask(1).Text) > 0) + _
       Int(Len(txtIpMask(2).Text) > 0) + Int(Len(txtIpMask(3).Text) > 0)
  If ii = -4 Then
    'verify if number
    szIpMask = ""
    For ii = 0 To 3
      If IsNumeric(txtIpMask(ii).Text) Then
        If txtIpMask(ii).Text < 0 Or txtIpMask(ii).Text > 255 Then
          MsgBox "Error IpMask not valid (range 0 to 255) !!", vbSystemModal + vbExclamation
          txtIpMask(ii).SetFocus
          Exit Sub
        End If
        szIpMask = szIpMask & txtIpMask(ii).Text & "."
      Else
        MsgBox "Error IpMask not valid !!", vbSystemModal + vbExclamation
        txtIpMask(ii).SetFocus
        Exit Sub
      End If
    Next
    szIpMask = Trim(szIpMask)
    szIpMask = Mid(szIpMask, 1, Len(szIpMask) - 1)
  ElseIf ii = 0 Then
    szIpMask = ""
  Else
    MsgBox "Error IpMask not valid !!", vbSystemModal + vbExclamation
    txtIpMask(0).SetFocus
    Exit Sub
  End If
  
  'verify if line Exists
  For Each lvItem In lv.ListItems
    If lvItem.Text = cboType.SelectedItem.Key And _
       lvItem.SubItems(1) = szDatabase And _
       lvItem.SubItems(2) = szUG And _
       lvItem.SubItems(3) = szIpAdrress And _
       lvItem.SubItems(4) = szIpMask And _
       lvItem.SubItems(5) = cboAut.SelectedItem.Key Then
      
      MsgBox "Line already Exists !!", vbSystemModal + vbExclamation
      Exit Sub
    End If
  Next
  
  If Index = 0 Then
    'add
    Set lvItem = lv.ListItems.Add(, , cboType.SelectedItem.Key, "property", "property")
    lvItem.SubItems(1) = szDatabase
    lvItem.SubItems(2) = szUG
    lvItem.SubItems(3) = szIpAdrress
    lvItem.SubItems(4) = szIpMask
    lvItem.SubItems(5) = cboAut.SelectedItem.Key
  Else
    'update
    lv.SelectedItem.Text = cboType.SelectedItem.Key
    lv.SelectedItem.SubItems(1) = szDatabase
    lv.SelectedItem.SubItems(2) = szUG
    lv.SelectedItem.SubItems(3) = szIpAdrress
    lv.SelectedItem.SubItems(4) = szIpMask
    lv.SelectedItem.SubItems(5) = cboAut.SelectedItem.Key
  End If
  lv.Tag = "Y"
  
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdAdd_Click"
End Sub

Private Sub cmdManualAddDB_Click()
On Error GoTo Err_Handler

Dim lX, lY As Long

lY = lvDatabase.ListItems.Count
If txtDatabase.Text = "" Then Exit Sub
      
  'Is there anything listed in the listbox ?
  If lY > 0 Then
    For lX = 1 To lY
      If InStr(1, lvDatabase.ListItems(lX).Text, txtDatabase.Text) > 0 Then
        MsgBox "Database '" & lvDatabase.ListItems(lX).Text & "' is already in the list", vbOKOnly + vbCritical, "Error"
        txtDatabase.SelStart = 0
        txtDatabase.SelLength = Len(txtDatabase.Text)
        Exit Sub
      End If
    Next lX
  End If
  
         
  lvDatabase.ListItems.Add , txtDatabase.Text, txtDatabase.Text, "database", "database"
  txtDatabase.Text = ""

Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdManualAddDB"
End Sub


Public Sub cmdManualUsrGroupAdd_Click()
On Error GoTo Err_Handler

Dim lX, lY As Long

lY = lvUG.ListItems.Count
If txtUG.Text = "" Then Exit Sub
      
  'Is there anything listed in the listbox ?
  If lY > 0 Then
    If optUG(0).Value = True Then   'Check Users
      For lX = 1 To lY
        If InStr(1, lvUG.ListItems(lX).Text, txtUG.Text) And InStr(1, lvUG.ListItems(lX).Key, "USR-") > 0 Then
          MsgBox "User '" & lvUG.ListItems(lX).Text & "' is already in the list", vbOKOnly + vbCritical, "Error"
          Exit Sub
        End If
      Next lX
    Else  'Check Groups
      For lX = 1 To lY
        If InStr(1, lvUG.ListItems(lX).Text, txtUG.Text) And InStr(1, lvUG.ListItems(lX).Key, "GRP-") > 0 Then
          MsgBox "Group '" & lvUG.ListItems(lX).Text & "' is already in the list", vbOKOnly + vbCritical, "Error"
          Exit Sub
        End If
      Next lX
    End If
  End If
      
  If optUG(0).Value = True Then
    lvUG.ListItems.Add , "USR-" & GetID, txtUG.Text, "user", "user"
  Else
    lvUG.ListItems.Add , "GRP-" & GetID, txtUG.Text, "group", "group"
  End If

Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdManualUsrGroupAdd_Click"
End Sub


Private Sub cmdOK_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

Dim iFile As Integer
Dim szTemp As String
Dim szData As String
Dim lvItem As ListItem
Dim vData
Dim ii As Integer

  If lv.Tag = "N" Then
    MsgBox "The configuration is unchanged!", vbSystemModal + vbInformation
    Exit Sub
  End If

  If MsgBox("Save new configuration?", vbSystemModal + vbYesNo) = vbNo Then Exit Sub
  'backup file using current date
  If ChkBckConf.Value = 1 Then
    Dim szSplitTmp() As String
    szSplitTmp = Split(szConfFileName, ".")
    MsgBox szConfFilePath & szSplitTmp(0) & Format(Date, "yyyy_mm_dd") & "." & szSplitTmp(1)
    FileCopy szConfFilePath & szConfFileName, szConfFilePath & szSplitTmp(0) & Format(Date, "_yyyy_mm_dd") & "." & szSplitTmp(1)
  End If

  'load file
  iFile = FreeFile
  Open szConfFilePath & szConfFileName For Input As #iFile
  szTemp = Input(LOF(iFile), #iFile)
  Close #iFile
  
  If InStr(szTemp, vbCrLf) > 0 Then
    vData = Split(szTemp, vbCrLf)
  ElseIf InStr(szTemp, vbCr) > 0 Then
    vData = Split(szTemp, vbCr)
  ElseIf InStr(szTemp, vbLf) > 0 Then
    vData = Split(szTemp, vbLf)
  End If
  
  'comment old line and mark
  szData = ""
  For ii = 0 To UBound(vData)
    szTemp = Trim(vData(ii))
    If Len(szTemp) = 0 Or Left(szTemp, 1) = "#" Then
      'comment
      szData = szData & szTemp
    Else
      szData = szData & "# Change on pgAdmin II " & Date & " " & vData(ii)
    End If
    szData = szData & vbCrLf
  Next
  
  'save file
  iFile = FreeFile
  Open szConfFilePath & szConfFileName For Output As #iFile
  szData = szData & "# New Line on pgAdmin II " & Date & vbCrLf
  szData = szData & "# TYPE  DATABASE    USER        IP-ADDRESS        IP-MASK           METHOD" & vbCrLf
  
  For Each lvItem In lv.ListItems
    szData = szData & lvItem.Text & "  " & _
                      lvItem.SubItems(1) & "  " & _
                      lvItem.SubItems(2) & "  " & _
                      lvItem.SubItems(3) & "  " & _
                      lvItem.SubItems(4) & "  " & _
                      lvItem.SubItems(5) & "  "
    szData = szData & vbCrLf
  Next
  Print #iFile, szData
  Close #iFile
  
  MsgBox "Please reread the configuration files (pg_ctl reload)", vbSystemModal + vbInformation
  Unload Me

Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdAdd_Click"
End Sub

Public Sub Initialise()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

Dim objDatabase As pgDatabase
Dim objUser As pgUser
Dim objGroup As pgGroup
Dim lvItem As ListItem

Dim iFile As Integer
Dim ii As Integer
Dim iPos As Integer
Dim iFlag As Integer

Dim szPath As String
Dim szTemp As String
Dim szTempArray() As String
Dim szVal As String
Dim vAuthMethod
Dim vData

Dim InQuote As Boolean

Dim lX As Long
Dim lY As Long
Dim lA As Long
Dim lB As Long

  'Display the file open box to recover the config file
  With cdlfrmWizard
    .Filter = "All Files (*.*)|*.*|Conf Files (*.conf)|*.conf"
    .FilterIndex = 1
    .ShowOpen
  End With

  'Check for the existence of the file
  If Dir(cdlfrmWizard.FileName) = "" Then
    MsgBox "File " & cdlfrmWizard.FileName & " not found ", vbSystemModal + vbCritical, "Error: File not found"
    bRunning = False
    Unload Me
  End If
    
  'type connession
  cboType.ComboItems.Add , "host", "host", "property", "property"
  cboType.ComboItems.Add , "local", "local", "property", "property"
  cboType.ComboItems.Add , "hostssl", "hostssl", "property", "property"
  cboType.ComboItems(1).Selected = True

  'database
  cboDatabase.ComboItems.Add , "all", "All Database", "all", "all"
  cboDatabase.ComboItems.Add , "file", "From file", "property", "property"
  cboDatabase.ComboItems.Add , "database", "Database", "database", "database"
  cboDatabase.ComboItems(1).Selected = True
  cboDatabase_Click
  
  For Each objDatabase In svr.Databases
      'If Not (objDatabase.SystemObject And Not ctx.IncludeSys) And objDatabase.AllowConnections Then
      If Not (objDatabase.SystemObject) And objDatabase.AllowConnections Then
        lvDatabase.ListItems.Add , objDatabase.Name, objDatabase.Name, "database", "database"
      End If
  Next

  'user
  cboUG.ComboItems.Add , "all", "All User", "all", "all"
  cboUG.ComboItems.Add , "file", "From file", "property", "property"
  cboUG.ComboItems.Add , "ug", "User/Group", "user", "user"
  cboUG.ComboItems(1).Selected = True

  'User
  For Each objUser In svr.Users
    lvUG.ListItems.Add , "USR-" & GetID, objUser.Name, "user", "user"
  Next
  
  'Group
  For Each objGroup In svr.Groups
    lvUG.ListItems.Add , "GRP-" & GetID, objGroup.Name, "group", "group"
  Next
  
  
  vAuthMethod = Array("trust", "reject", "md5", "crypt", "password", "krb4", "krb5", "ident", "pam")
  
  lB = UBound(vAuthMethod)
  'Authentication Method
  For lA = 0 To lB
    cboAut.ComboItems.Add , vAuthMethod(lA), vAuthMethod(lA), "key", "key"
  Next lA
  cboAut.ComboItems(1).Selected = True
  
  'read file and load configuration
  iFile = FreeFile
  
  'Get the filename and file path
  szConfFileName = cdlfrmWizard.FileTitle
  
  lX = Len(cdlfrmWizard.FileName)
  For lY = lX To 1 Step -1
    If InStr(lY, cdlfrmWizard.FileName, "\") > 0 Then
      szConfFilePath = Mid$(cdlfrmWizard.FileName, 1, lY)
      Exit For
    End If
  Next lY

  Open szConfFilePath & szConfFileName For Input As #iFile
    szTemp = Input(LOF(iFile), #iFile)
  Close #iFile
  
  If InStr(szTemp, vbCrLf) > 0 Then
    vData = Split(szTemp, vbCrLf)
  ElseIf InStr(szTemp, vbCr) > 0 Then
    vData = Split(szTemp, vbCr)
  ElseIf InStr(szTemp, vbLf) > 0 Then
    vData = Split(szTemp, vbLf)
  End If
  
  
  For ii = 0 To UBound(vData)
    szTemp = Trim(vData(ii))
    If Len(szTemp) = 0 Or Left(szTemp, 1) = "#" Then
      'comment
    Else
      'convert tabs into spaces
      szTemp = Trim(Replace(szTemp, vbTab, " "))
      
      iFlag = 0
      InQuote = False
      szVal = ""
      For iPos = 1 To Len(szTemp)
        szVal = szVal & Mid(szTemp, iPos, 1)
        If Mid(szTemp, iPos, 1) = QUOTE Then
          InQuote = Not InQuote
        Else
          If Not InQuote And (Mid(szTemp, iPos, 1) = " " Or Len(szTemp) = iPos) Then
            szVal = Trim(szVal)
            
            If Len(szVal) > 0 Then
              If iFlag = 0 Then
                'Type connection
                Set lvItem = lv.ListItems.Add(, , szVal, "property", "property")
                iFlag = iFlag + 1
              ElseIf iFlag = 1 Then
                'database
                lvItem.SubItems(iFlag) = szVal
                iFlag = iFlag + 1
              ElseIf iFlag = 2 Then
                'user/group
                lvItem.SubItems(iFlag) = szVal
                iFlag = iFlag + 1
              ElseIf iFlag > 2 Then
                If UBound(Filter(vAuthMethod, LCase(szVal))) >= 0 Then
                  lvItem.SubItems(5) = szVal
                  Exit For
                Else
                  lvItem.SubItems(iFlag) = szVal
                  iFlag = iFlag + 1
                End If
              End If
            End If
            szVal = ""
          End If
        End If
      Next
    End If
  Next
  lv.Tag = "N"
  
   'Recover any other databases from the listview and add them to the lvdatabase listview
   lY = lv.ListItems.Count
    For lX = 1 To lY
      If InStr(1, szTemp, lv.ListItems(lX).SubItems(1)) = 0 Then
        If LCase(lv.ListItems(lX).SubItems(1)) <> "all" Then ' Check it's not All DB's
          szTemp = lv.ListItems(lX).SubItems(1)
          ReDim szTempArray(0)
          szTempArray = Split(szTemp, ",")
          lB = UBound(szTempArray)
          For lA = 0 To lB
            If SearchListview(lvDatabase, szTempArray(lA)) = False Then lvDatabase.ListItems.Add , szTempArray(lA), szTempArray(lA), "database", "database"
          Next lA
        End If
      End If
    Next lX
    lvDatabase.SortOrder = lvwAscending
    lvDatabase.Sorted = True

    'Need to recover the list of users and or groups
    szTemp = ""
    For lX = 1 To lY
      If InStr(1, szTemp, lv.ListItems(lX).SubItems(2)) = 0 Then
        If LCase(lv.ListItems(lX).SubItems(2)) <> "all" Then
          szTemp = szTemp & " " & lv.ListItems(lX).SubItems(2)
            ReDim szTempArray(0)
            szTempArray = Split(szTemp, ",")
            lB = UBound(szTempArray)
            For lA = 0 To lB
              If Mid$(szTempArray(lA), 1, 1) = "+" Then
                lvUG.ListItems.Add , "GRP-" & GetID, Trim(Mid$(szTempArray(lA), 2)), "group", "group"
              Else
                lvUG.ListItems.Add , "USR-" & GetID, Trim(szTempArray(lA)), "user", "user"
              End If
            Next lA
        End If
      End If
    Next lX
    
    'Size the listvire
    AutoSizeColumnLv lv
  
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Initialise"
End Sub
Private Sub cboDatabase_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cboDatabase_Click()", etFullDebug

Dim lvItem As ListItem

  txtFileDatabase.Visible = False
  txtFileDatabase.Text = ""
  lvDatabase.Visible = False
  fraDBadd.Visible = False
  
  For Each lvItem In lvDatabase.ListItems
    lvItem.Selected = False
  Next
  
  Select Case cboDatabase.SelectedItem.Key
    Case "all"
    
    Case "file"
      txtFileDatabase.Visible = True
      
    Case "database"
      lvDatabase.Visible = True
      fraDBadd.Visible = True
      
  End Select
  
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cboDatabase_Click"
End Sub

Private Sub cboUG_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cboUG_Click()", etFullDebug

Dim lvItem As ListItem

  txtFileUG.Visible = False
  txtFileUG.Text = ""
  lvUG.Visible = False
  fraUG.Visible = False
  
  For Each lvItem In lvUG.ListItems
    lvItem.Selected = False
  Next
  
  Select Case cboUG.SelectedItem.Key
    Case "all"
    
    Case "file"
      txtFileUG.Visible = True
      
    Case "ug"
      lvUG.Visible = True
      fraUG.Visible = True
      
  End Select
  
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cboUG_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Exiting " & App.Title & ":frmWizard.Form_Unload", etFullDebug
  
  bRunning = False
  Unload Me
  
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Unload"
End Sub

Private Sub lv_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.lv_Click()", etFullDebug

Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
Dim ii As Integer
Dim bFound As Boolean

  If Not (lv.SelectedItem Is Nothing) Then
    'type
    cboType.ComboItems(LCase(lv.SelectedItem.Text)).Selected = True
    
    'database
    szTemp = lv.SelectedItem.SubItems(1)
    If szTemp = "all" Then
      cboDatabase.ComboItems(szTemp).Selected = True
      cboDatabase_Click
    Else
      If Left(szTemp, 1) = "@" Then
        'from file
        cboDatabase.ComboItems("file").Selected = True
        cboDatabase_Click
        txtFileDatabase = Mid(szTemp, 2)
      Else
        'database
        For Each lvItem In lvDatabase.ListItems
          lvItem.Checked = False
        Next
        cboDatabase.ComboItems("database").Selected = True
        cboDatabase_Click
        vData = Split(szTemp, ",")
        For ii = 0 To UBound(vData)
          bFound = False
          For Each lvItem In lvDatabase.ListItems
            If LCase(vData(ii)) = LCase(lvItem.Text) Then
              lvItem.Checked = True
              bFound = True
              Exit For
            End If
          Next
          If Not bFound Then
            'not found database
            MsgBox "Error in configuration Database" & vbCrLf & vData(ii) & " not found ", vbSystemModal + vbCritical
          End If
        Next
      End If
    End If
  
    'user
    szTemp = lv.SelectedItem.SubItems(2)
    If szTemp = "all" Then
      cboUG.ComboItems(szTemp).Selected = True
      cboUG_Click
    Else
      If Left(szTemp, 1) = "@" Then
        'from file
        cboUG.ComboItems("file").Selected = True
        cboUG_Click
        txtFileUG = Mid(szTemp, 2)
      Else
        'user/group
        cboUG.ComboItems("ug").Selected = True
        cboUG_Click
        vData = Split(szTemp, ",")
        For ii = 0 To UBound(vData)
          bFound = False
          For Each lvItem In lvUG.ListItems
            If Left(vData(ii), 1) = "+" Then
              'group
              If LCase(Mid(vData(ii), 2)) = LCase(lvItem.Text) And Left(lvItem.Key, 3) = "GRP" Then
                lvItem.Checked = True
                bFound = True
                Exit For
              End If
            Else
              'user
              If LCase(vData(ii)) = LCase(lvItem.Text) And Left(lvItem.Key, 3) = "USR" Then
                lvItem.Checked = True
                bFound = True
                Exit For
              End If
            End If
          Next
          
          If Not bFound Then
            'user/group
            MsgBox "Error in configuration User/Group" & vbCrLf & vData(ii) & " not found ", vbSystemModal + vbCritical
          End If
        Next
      End If
    End If
  
    'ip-address
    For ii = 0 To 3
      txtIpAddress(ii).Text = ""
    Next
    szTemp = lv.SelectedItem.SubItems(3)
    If Len(szTemp) > 0 Then
      vData = Split(szTemp, ".")
      For ii = 0 To 3
        txtIpAddress(ii).Text = vData(ii)
      Next
    End If
  
    'ip-mask
    For ii = 0 To 3
      txtIpMask(ii).Text = ""
    Next
    szTemp = lv.SelectedItem.SubItems(4)
    If Len(szTemp) > 0 Then
      vData = Split(szTemp, ".")
      For ii = 0 To 3
        txtIpMask(ii).Text = vData(ii)
      Next
    End If
  
    'Authentication Method
    cboAut.ComboItems(LCase(lv.SelectedItem.SubItems(5))).Selected = True
  End If
  
Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.lv_Click"
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdCancel_Click()", etFullDebug
  
  bRunning = False
  Unload Me

Exit Sub
Err_Handler:
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdCancel_Click"
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdRemove_Click()", etFullDebug

  If lv.SelectedItem Is Nothing Then
    MsgBox "You must select a line to remove!", vbSystemModal + vbExclamation, "Error"
    Exit Sub
  End If
  
  If MsgBox("Confirm remove line?", vbSystemModal + vbYesNo) = vbYes Then
    lv.ListItems.Remove lv.SelectedItem.Index
    lv.Tag = "Y"
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdRemove_Click"
End Sub

