VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Debugging Tools"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList il 
      Left            =   6885
      Top             =   3735
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
            Picture         =   "frmWizard.frx":08CA
            Key             =   "property"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":0E64
            Key             =   "ping"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1CB6
            Key             =   "dsn"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWizard.frx":1E10
            Key             =   "timeout"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabWizard 
      Height          =   4245
      Left            =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   7488
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " ICMP Ping"
      TabPicture(0)   =   "frmWizard.frx":26EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCount"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdPing"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtHost"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "statframe"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lvResults"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   " ODBC Connect"
      TabPicture(1)   =   "frmWizard.frx":2706
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(5)"
      Tab(1).Control(1)=   "Label1(6)"
      Tab(1).Control(2)=   "Label1(7)"
      Tab(1).Control(3)=   "lvDetails"
      Tab(1).Control(4)=   "txtPWD"
      Tab(1).Control(5)=   "txtUID"
      Tab(1).Control(6)=   "cmdConnect"
      Tab(1).Control(7)=   "cboDatasource"
      Tab(1).ControlCount=   8
      Begin MSComctlLib.ImageCombo cboDatasource 
         Height          =   330
         Left            =   -73515
         TabIndex        =   5
         ToolTipText     =   "Select a datasource to attempt to connect to."
         Top             =   450
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "il"
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Height          =   1050
         Left            =   -69060
         TabIndex        =   8
         ToolTipText     =   "Attempt to connect to the data source."
         Top             =   450
         Width           =   1365
      End
      Begin VB.TextBox txtUID 
         Height          =   285
         Left            =   -73515
         TabIndex        =   6
         ToolTipText     =   "Enter a valid username for this datasource"
         Top             =   855
         Width           =   4290
      End
      Begin VB.TextBox txtPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73515
         PasswordChar    =   "*"
         TabIndex        =   7
         ToolTipText     =   "Enter a valid password for this datasource."
         Top             =   1215
         Width           =   4290
      End
      Begin MSComctlLib.ListView lvResults 
         Height          =   1905
         Left            =   135
         TabIndex        =   3
         Top             =   1170
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
            Text            =   "Host"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Seq"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "RTT"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame statframe 
         Caption         =   "Statistics"
         Height          =   1005
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "Displays Statisitcs on the current ping."
         Top             =   3150
         Width           =   7215
         Begin VB.Label lblAverage 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   2385
            TabIndex        =   22
            Top             =   630
            Width           =   90
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Average RTT (round trip time):"
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   21
            Top             =   630
            Width           =   2175
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Sent :"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   20
            Top             =   315
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Received:"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   19
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "% Loss :"
            Height          =   195
            Index           =   2
            Left            =   5625
            TabIndex        =   18
            Top             =   315
            Width           =   585
         End
         Begin VB.Label lblSent 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   720
            TabIndex        =   17
            Top             =   315
            Width           =   90
         End
         Begin VB.Label lblReceived 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   3780
            TabIndex        =   16
            Top             =   315
            Width           =   90
         End
         Begin VB.Label lblLoss 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   6345
            TabIndex        =   15
            Top             =   315
            Width           =   90
         End
         Begin VB.Label lblMinimum 
            Alignment       =   2  'Center
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   600
         End
         Begin VB.Label lblMaximum 
            Alignment       =   2  'Center
            Height          =   255
            Left            =   3840
            TabIndex        =   13
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1485
         TabIndex        =   0
         ToolTipText     =   "Enter the name or IP address of the host you wish to ping."
         Top             =   450
         Width           =   4305
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "&Ping"
         Height          =   660
         Left            =   5940
         TabIndex        =   1
         ToolTipText     =   "Ping the entered hostname/IP address"
         Top             =   450
         Width           =   1365
      End
      Begin VB.TextBox txtCount 
         Height          =   285
         Left            =   1485
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "5"
         ToolTipText     =   "Select how many times the host should be pinged."
         Top             =   810
         Width           =   555
      End
      Begin MSComctlLib.ListView lvDetails 
         Height          =   2535
         Left            =   -74865
         TabIndex        =   9
         Top             =   1575
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "il"
         SmallIcons      =   "il"
         ColHdrIcons     =   "il"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Property"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   5997
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Datasource"
         Height          =   195
         Index           =   7
         Left            =   -74820
         TabIndex        =   25
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   195
         Index           =   6
         Left            =   -74820
         TabIndex        =   24
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         Height          =   195
         Index           =   5
         Left            =   -74820
         TabIndex        =   23
         Top             =   1260
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Number of Pings"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   855
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Host/IP Address"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   495
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

Public Sub Initialise()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Initialise()", etFullDebug

Dim szDSN As String * 1024
Dim szDRV As String * 1024
Dim szDSNItem As String
Dim lRet As Long
Dim lDSN As Integer
Dim lDRV As Integer
Dim lHenv As Long

  cboDatasource.ComboItems.Clear
  
  'Get the DSNs
  If SQLAllocEnv(lHenv) <> -1 Then
    Do Until lRet <> SQL_SUCCESS
      szDSN = Space(1024)
      szDRV = Space(1024)
      lRet = SQLDataSources(lHenv, SQL_FD_FETCH_NEXT, szDSN, 1024, lDSN, szDRV, 1024, lDRV)
      szDSNItem = Left(szDSN, lDSN)
      If Trim(szDSNItem) <> "" Then cboDatasource.ComboItems.Add , , szDSNItem, "dsn", "dsn"
    Loop
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Load"
End Sub

Private Sub cmdConnect_Click()
On Error GoTo Cleanup
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdConnect_Click()", etFullDebug

Dim objItem As ListItem
Dim iSize As Integer
Dim lEnv As Long
Dim lDBC As Long
Dim lRet As Long
Dim szResult As String * 1024
Dim szConnect As String

  lvDetails.ListItems.Clear
  
  If cboDatasource.Text = "" Then
    MsgBox "You must select a Datasource to connect to!", vbExclamation, "Error"
    cboDatasource.SetFocus
    Exit Sub
  End If
  
  StartMsg "Connecting to " & cboDatasource.Text
  
  szConnect = "DSN=" & cboDatasource.Text
  If txtUID.Text <> "" Then szConnect = szConnect & ";UID=" & txtUID.Text
  If txtPWD.Text <> "" Then szConnect = szConnect & ";PWD=" & txtPWD.Text
  
  'Initialise the ODBC subsystem
  If SQLAllocEnv(lEnv) <> 0 Then
    EndMsg
    MsgBox "Couldn't initialise the ODBC subsystem!", vbExclamation, "Error"
    svr.LogEvent "Couldn't initialise the ODBC subsystem!", etMiniDebug
    Exit Sub
  End If

  'Allocate space for the connection object
  If SQLAllocConnect(lEnv, lDBC) <> 0 Then
    EndMsg
    MsgBox "Couldn't allocate memory for the ODBC connection!", vbExclamation, "Error"
    svr.LogEvent "Couldn't allocate memory for the ODBC connection!", etMiniDebug
    GoTo Cleanup
  End If

  'Connect
  lRet = SQLDriverConnect(lDBC, Me.hWnd, szConnect, Len(szConnect), szResult, Len(szResult), iSize, SQL_DRIVER_COMPLETE_REQUIRED)
  If (lRet <> SQL_SUCCESS) And (lRet <> SQL_SUCCESS_WITH_INFO) Then
    EndMsg
    MsgBox "An ODBC error occured whilst connecting!" & vbCrLf & "Please check the connection details and try again.", vbExclamation, "Error"
    svr.LogEvent "An ODBC error occured whilst connecting!" & vbCrLf & "Please check the connection details and try again.", etMiniDebug
    GoTo Cleanup
  End If
  

  'Get some details...
  szResult = ""
  SQLGetInfoString lDBC, SQL_DRIVER_NAME, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Driver Name", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)
  
  szResult = ""
  SQLGetInfoString lDBC, SQL_DRIVER_VER, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Driver Version", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)

  szResult = ""
  SQLGetInfoString lDBC, SQL_DRIVER_ODBC_VER, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Driver ODBC Version", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)

  szResult = ""
  SQLGetInfoString lDBC, SQL_DATA_SOURCE_NAME, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Datasource Name", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)
  
  szResult = ""
  SQLGetInfoString lDBC, SQL_SERVER_NAME, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Server Name", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)
  
  szResult = ""
  SQLGetInfoString lDBC, SQL_DBMS_NAME, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "DBMS Name", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)
  
  szResult = ""
  SQLGetInfoString lDBC, SQL_DBMS_VER, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "DBMS Version", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)
  
  szResult = ""
  SQLGetInfoString lDBC, SQL_USER_NAME, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Username", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)

  szResult = ""
  SQLGetInfoString lDBC, SQL_DATA_SOURCE_READ_ONLY, szResult, Len(szResult), iSize
  Set objItem = lvDetails.ListItems.Add(, , "Read Only?", "property", "property")
  objItem.SubItems(1) = Left(szResult, iSize)
  
Cleanup:
  On Error Resume Next
  EndMsg
  If lDBC <> 0 Then SQLDisconnect lDBC
  SQLFreeConnect lDBC
  If lEnv <> 0 Then SQLFreeEnv lEnv
End Sub

Private Sub cmdPing_Click()
'On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.cmdPing_Click()", etFullDebug

Dim szTriptime As String
Dim objOpts As ICMP_OPTIONS
Dim objEcho As ICMP_ECHO_REPLY
Dim objItem As ListItem
Dim lPort As Long
Dim lAddress As Long
Dim lSent As Long
Dim lReceived As Long
Dim lTotalRTT As Long
Dim sStart As Single
Dim szAddress As String
Dim szHostIP As String
Dim szHostName As String
Dim szData As String
Dim szRTT As String


  StartMsg "Pinging " & txtHost.Text & "..."
  lvResults.ListItems.Clear
  lblSent.Caption = 0
  lblReceived.Caption = 0
  lblLoss.Caption = 0
  lblAverage.Caption = 0
   
  If NetInitialise <> "" Then
    ReDim objPingStatistics(0)

    szAddress = GetIPFromHostName(txtHost.Text)
    If szAddress = "" Then
      EndMsg
      MsgBox "Couldn't resolve the hostname '" & txtHost.Text & "'!", vbExclamation, "Error"
      svr.LogEvent "Couldn't resolve the hostname '" & txtHost.Text & "'!", etMiniDebug
      txtHost.SetFocus
      txtHost.SelStart = 0
      txtHost.SelLength = Len(txtHost.Text)
      Exit Sub
    End If
    lAddress = inet_addr(szAddress)
    lPort = IcmpCreateFile()
      
    For lSent = 1 To CInt(Val(txtCount.Text))
     
      If lPort <> 0 Then
        objOpts.TTL = 100
        szData = String(59, "a") & "SEQ" & Format(lSent, "00")
      
        If IcmpSendEcho(lPort, lAddress, szData, Len(szData), objOpts, objEcho, Len(objEcho) + 8, 2400) = 1 Then
          szHostIP = AddrToIP(objEcho.Address)
          szHostName = GetHostByAddress(objEcho.Address)
        End If
      
        Select Case objEcho.RoundTripTime
          Case Is < 10: szRTT = "<10 ms"
          Case Is > 1200: szRTT = "*"
          Case Else: szRTT = objEcho.RoundTripTime & " ms"
        End Select
       
        If szRTT = "*" Then
          Set objItem = lvResults.ListItems.Add(, , "Timeout!", "timeout", "timeout")
        Else
          Set objItem = lvResults.ListItems.Add(, , szHostName, "ping", "ping")
          objItem.SubItems(1) = szHostIP
          objItem.SubItems(2) = Val(Mid(objEcho.ReturnedData, InStr(objEcho.ReturnedData, "SEQ") + 3, 2))
          objItem.SubItems(3) = szRTT
          lReceived = lReceived + 1
          lTotalRTT = lTotalRTT + objEcho.RoundTripTime
        End If
        
        'Display stats
        lblSent.Caption = lSent
        lblReceived.Caption = lReceived
        lblLoss.Caption = ((lSent - lReceived) / lSent) * 100
        If lTotalRTT > 0 And lReceived > 0 Then
          lblAverage.Caption = lTotalRTT / lReceived & " ms"
        Else
          lblAverage.Caption = "0" & " ms"
        End If
  
        'Pause for a second
        If lSent < Val(txtCount.Text) Then
          sStart = Timer
          While Timer < sStart + 1
            DoEvents
          Wend
        End If
      
      Else 'Couldn't open port
        EndMsg
        MsgBox "Couldn't open an ICMP port!", vbExclamation, "Error"
        svr.LogEvent "Couldn't open an ICMP port!", etMiniDebug
        Exit Sub
      End If
      
    Next lSent
  Else
    EndMsg
    MsgBox "Couldn't initialise Winsock!", vbExclamation, "Error"
    svr.LogEvent "Couldn't initialise Winsock!", etMiniDebug
  End If
  NetShutDown
  EndMsg
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.cmdPing_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":frmWizard.Form_Unload()", etFullDebug

  bRunning = False

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmWizard.Form_Unload"
End Sub
