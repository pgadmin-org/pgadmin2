VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmUpgradeWizard 
   Caption         =   "Upgrade Wizard"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   Icon            =   "frmUpgradeWizard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList il 
      Left            =   540
      Top             =   3735
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
            Picture         =   "frmUpgradeWizard.frx":0A02
            Key             =   "upgrade"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpgradeWizard.frx":12DC
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpgradeWizard.frx":1EAE
            Key             =   "unknown"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   330
      Left            =   6480
      TabIndex        =   4
      ToolTipText     =   "Move forward a stage"
      Top             =   3960
      Width           =   960
   End
   Begin VB.PictureBox picStrip 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "frmUpgradeWizard.frx":2D00
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   3
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5445
      TabIndex        =   2
      ToolTipText     =   "Move back a stage"
      Top             =   3960
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   6480
      TabIndex        =   0
      ToolTipText     =   "Return SQL and exit."
      Top             =   3960
      Visible         =   0   'False
      Width           =   960
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   3840
      Left            =   495
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   6773
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   176
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmUpgradeWizard.frx":3ACE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAuto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboFrequency"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtServer"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmUpgradeWizard.frx":3AEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvVersions"
      Tab(1).Control(1)=   "Label1(2)"
      Tab(1).ControlCount=   2
      Begin MSComctlLib.ListView lvVersions 
         Height          =   2400
         Left            =   -74820
         TabIndex        =   12
         Top             =   1215
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   4233
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Software"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Installed."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Available"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Download Site"
            Object.Width           =   4745
         EndProperty
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   2700
         TabIndex        =   10
         Top             =   2610
         Width           =   3570
      End
      Begin VB.ComboBox cboFrequency 
         Height          =   315
         ItemData        =   "frmUpgradeWizard.frx":3B06
         Left            =   4365
         List            =   "frmUpgradeWizard.frx":3B16
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1710
         Width           =   1320
      End
      Begin VB.CheckBox chkAuto 
         Caption         =   "&Automatically run the Upgrade Wizard every "
         Height          =   240
         Left            =   855
         TabIndex        =   7
         Top             =   1755
         Width           =   3480
      End
      Begin VB.Label Label1 
         Caption         =   $"frmUpgradeWizard.frx":3B32
         Height          =   780
         Index           =   2
         Left            =   -74820
         TabIndex        =   11
         Top             =   225
         Width           =   6630
      End
      Begin VB.Label Label2 
         Caption         =   "pgAdmin website address"
         Height          =   240
         Left            =   630
         TabIndex        =   9
         Top             =   2655
         Width           =   2040
      End
      Begin VB.Label Label1 
         Caption         =   "You may also alter the Upgrade Wizard's settings on this page."
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   900
         Width           =   6630
      End
      Begin VB.Label Label1 
         Caption         =   $"frmUpgradeWizard.frx":3C47
         Height          =   645
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   225
         Width           =   6630
      End
   End
End
Attribute VB_Name = "frmUpgradeWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmUpgradeWizard.frm - Check for Upgrades.

Option Explicit
Dim bButtonPress As Boolean
Dim bProgramPress As Boolean

Private Sub lvVersions_DblClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.lvVersions_DblClick()", etFullDebug

Dim hDC As Long

  If Not lvVersions.SelectedItem Is Nothing Then
    hDC = GetDesktopWindow()
    ShellExecute hDC, "Open", lvVersions.SelectedItem.SubItems(3), "", "C:\", SW_SHOWNORMAL
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.lvVersions_DblClick"
End Sub

Private Sub txtServer_Change()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.txtServer_Change()", etFullDebug

  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Server", regString, txtServer.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.txtServer_Change"
End Sub

Private Sub cboFrequency_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.cboFrequency_Click()", etFullDebug

  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Frequency", regString, cboFrequency.Text
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.cboFrequency_Click"
End Sub

Private Sub chkAuto_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.chkAuto_Click()", etFullDebug

  If chkAuto.Value = 1 Then
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Check", regString, "Y"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Check", regString, "N"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.chkAuto_Click"
End Sub

Private Sub cmdNext_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.cmdNext_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 0
      lvVersions.ListItems.Clear
      tabWizard.Tab = 1
      cmdNext.Enabled = False
      cmdNext.Visible = False
      cmdOK.Enabled = True
      cmdOK.Visible = True
      cmdPrevious.Enabled = True
      UpgradeCheck
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.cmdNext_Click"
End Sub

Private Sub cmdPrevious_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.cmdPrevious_Click()", etFullDebug

  bButtonPress = True
  Select Case tabWizard.Tab
    Case 1
      tabWizard.Tab = 0
      cmdNext.Enabled = True
      cmdNext.Visible = True
      cmdOK.Enabled = False
      cmdOK.Visible = False
      cmdPrevious.Enabled = False
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.cmdPrevious_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.cmdOK_Click()", etFullDebug

  Unload Me
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.cmdOK_Click"
End Sub

Private Sub Form_Load()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.Form_Load()", etFullDebug

  PatchForm Me
  
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Check", "Y")) = "Y" Then
    chkAuto.Value = 1
  Else
    chkAuto.Value = 0
  End If
  cboFrequency.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Frequency", "Week")
  txtServer.Text = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Server", "www.pgadmin.org")
  
  'Log the upgrade check. If the user doesn't actually run, assume that they meant to exit
  'and didn't want to be bugged by the wizard. The user can always run it from the menu...
  RegWrite HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Last Check", regString, Format(Date, "yyyy-MM-dd")
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.Form_Load"
End Sub

Private Sub tabWizard_Click(PreviousTab As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.tabWizard_Click(" & PreviousTab & ")", etFullDebug

  If bButtonPress = False And bProgramPress = False Then
    bProgramPress = True
    tabWizard.Tab = PreviousTab
  Else
    bProgramPress = False
  End If
  bButtonPress = False
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.tabWizard_Click"
End Sub

Private Sub UpgradeCheck()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.UpgradeCheck()", etFullDebug

Dim szData As String
Dim szBigBits() As String
Dim szLittleBits() As String
Dim vBigBit As Variant
Dim objItem As ListItem

  StartMsg "Contacting Server..."
  szData = GetVersions
  szData = Replace(szData, vbCr, "")
  szBigBits = Split(szData, vbLf)
  For Each vBigBit In szBigBits
    If (Left(Trim(vBigBit), 1) <> "#") And (Trim(vBigBit) <> "") Then
      szLittleBits = Split(vBigBit, "|")
      If UBound(szLittleBits) = 2 Then
        Select Case UCase(szLittleBits(0))
          Case "POSTGRESQL"
            If frmMain.svr.dbVersion.Major & "." & frmMain.svr.dbVersion.Minor & "." & frmMain.svr.dbVersion.Revision = "0.0.0" Then
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "unknown", "unknown")
            ElseIf VersionGreater(szLittleBits(1), frmMain.svr.dbVersion.Major & "." & frmMain.svr.dbVersion.Minor & "." & frmMain.svr.dbVersion.Revision) Then
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "upgrade", "upgrade")
            Else
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "ok", "ok")
            End If
            objItem.SubItems(1) = frmMain.svr.dbVersion.Major & "." & frmMain.svr.dbVersion.Minor & "." & frmMain.svr.dbVersion.Revision
          Case "PSQLODBC"
            If frmMain.svr.DriverVersion.Major & "." & frmMain.svr.DriverVersion.Minor & "." & frmMain.svr.DriverVersion.Revision = "0.0.0" Then
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "unknown", "unknown")
            ElseIf VersionGreater(szLittleBits(1), frmMain.svr.DriverVersion.Major & "." & frmMain.svr.DriverVersion.Minor & "." & frmMain.svr.DriverVersion.Revision) Then
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "upgrade", "upgrade")
            Else
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "ok", "ok")
            End If
            objItem.SubItems(1) = frmMain.svr.DriverVersion.Major & "." & frmMain.svr.DriverVersion.Minor & "." & frmMain.svr.DriverVersion.Revision
          Case UCase(App.Title)
            If VersionGreater(szLittleBits(1), App.Major & "." & App.Minor & "." & App.Revision) Then
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "upgrade", "upgrade")
            Else
              Set objItem = lvVersions.ListItems.Add(, , szLittleBits(0), "ok", "ok")
            End If
            objItem.SubItems(1) = App.Major & "." & App.Minor & "." & App.Revision
        End Select
        objItem.SubItems(2) = szLittleBits(1)
        objItem.SubItems(3) = szLittleBits(2)
      End If
    End If
  Next vBigBit
  EndMsg
  
  If lvVersions.ListItems.Count = 0 Then
    MsgBox "The list of current software couldn't be downloaded from the server. Please check the server name and your network connection. If problems still persist, please contact the Support Mailing list listed in the helpfile.", vbExclamation, "Error"
    cmdPrevious_Click
    txtServer.SetFocus
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.UpgradeCheck"
End Sub

Private Function VersionGreater(szVer1 As String, szVer2 As String) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.VersionGreater(" & QUOTE & szVer1 & QUOTE & ", " & QUOTE & szVer2 & QUOTE & ")", etFullDebug

Dim szBits1() As String
Dim szBits2() As String

  szBits1 = Split(szVer1, ".")
  szBits2 = Split(szVer2, ".")
  
  If (Val(szBits1(0)) > Val(szBits2(0))) Then
    VersionGreater = True
    Exit Function
  End If
  
  If (Val(szBits1(0)) >= Val(szBits2(0))) And (Val(szBits1(1)) > Val(szBits2(1))) Then
    VersionGreater = True
    Exit Function
  End If
  
  If (Val(szBits1(0)) >= Val(szBits2(0))) And (Val(szBits1(1)) >= Val(szBits2(1))) And (Val(szBits1(2)) > Val(szBits2(2))) Then
    VersionGreater = True
    Exit Function
  End If
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.VersionGreater"
End Function

Private Function GetVersions() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmUpgradeWizard.GetVersions()", etFullDebug

Dim lISession As Long
Dim lIConnect As Long
Dim lHttpOpenReq As Long
Dim szRead As String * 2048
Dim szBuffer As String
Dim lBufferLen As Long
Dim lNumBytes As Long
Dim bDoLoop As Boolean

  lBufferLen = Len(szBuffer)
  lISession = InternetOpen(App.Title & "v" & App.Major & "." & App.Minor & "." & App.Revision, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
  lIConnect = InternetConnect(lISession, txtServer.Text, INTERNET_DEFAULT_HTTP_PORT, vbNullString, vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
  lHttpOpenReq = HttpOpenRequest(lIConnect, "GET", "versions.dat", "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
  HttpSendRequest lHttpOpenReq, vbNullString, 0, 0, 0
  bDoLoop = True
  While bDoLoop
    szRead = vbNullString
    bDoLoop = InternetReadFile(lHttpOpenReq, szRead, Len(szRead), lNumBytes)
    szBuffer = szBuffer & Left$(szRead, lNumBytes)
    If Not CBool(lNumBytes) Then bDoLoop = False
  Wend

  InternetCloseHandle (lHttpOpenReq)
  InternetCloseHandle (lISession)
  InternetCloseHandle (lIConnect)
  GetVersions = szBuffer
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmUpgradeWizard.GetVersions"
End Function
