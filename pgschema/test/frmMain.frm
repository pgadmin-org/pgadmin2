VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "pgSchema Test Project"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSQL 
      BackColor       =   &H8000000F&
      Height          =   3570
      Left            =   3420
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   4185
      Width           =   7575
   End
   Begin MSComctlLib.ImageList ilTV 
      Left            =   2970
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "property"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27B2
            Key             =   "server"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F64
            Key             =   "servers"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7716
            Key             =   "groups"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FF0
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88CA
            Key             =   "users"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91A4
            Key             =   "user"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A7E
            Key             =   "databases"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A358
            Key             =   "database"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkIncludeSystem 
      Caption         =   "Include System Objects"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   810
      Value           =   1  'Checked
      Width           =   2760
   End
   Begin MSComctlLib.ListView lvInfo 
      Height          =   3030
      Left            =   3420
      TabIndex        =   10
      Top             =   1125
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   5345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   6585
      Left            =   0
      TabIndex        =   9
      Top             =   1125
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   11615
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   915
      Left            =   7380
      TabIndex        =   8
      Top             =   135
      Width           =   2265
   End
   Begin VB.TextBox txtPWD 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4635
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   450
      Width           =   2625
   End
   Begin VB.TextBox txtUID 
      Height          =   285
      Left            =   4635
      TabIndex        =   4
      Text            =   "postgres"
      Top             =   135
      Width           =   2625
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "5432"
      Top             =   450
      Width           =   870
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   135
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   3645
      TabIndex        =   7
      Top             =   495
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      Height          =   195
      Index           =   2
      Left            =   3645
      TabIndex        =   5
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   495
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   180
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001, Dave Page


Option Explicit

Dim WithEvents svr As pgServer
Attribute svr.VB_VarHelpID = -1

Private Function GetID() As String
Static lID As Long
  lID = lID + 1
  GetID = Format(lID, "00000000")
End Function

Private Sub cmdConnect_Click()
Dim szID As String
Dim nodX As Node

  'Create a new server object, and connect it.
  Set svr = New pgServer
  svr.Logfile = "C:\pgSchema.log"
  svr.LogLevel = llsql
  If chkIncludeSystem.Value = 1 Then
    svr.IncludeSys = True
  Else
    svr.IncludeSys = False
  End If
  svr.Connect txtServer, txtPort, txtUID, txtPWD
 
  'Start populating the treeview.
  tv.Nodes.Clear
  tv.Nodes.Add , , "SVR-" & GetID, svr.Server
End Sub

Private Sub svr_EventLog(EventLevel As pgSchema.LogLevel, EventMessage As String)
  If EventLevel < llsql Then Debug.Print EventLevel & " " & EventMessage
End Sub

Private Sub tvServer(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "DAT+" & GetID, "Databases (" & svr.Databases.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "GRP+" & GetID, "Groups (" & svr.Groups.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "USR+" & GetID, "Users (" & svr.Users.Count & ")"
  End If
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Hostname")
  lvItem.SubItems(1) = svr.Server & ""
  Set lvItem = lvInfo.ListItems.Add(, , "Port")
  lvItem.SubItems(1) = svr.Port
  Set lvItem = lvInfo.ListItems.Add(, , "Username")
  lvItem.SubItems(1) = svr.Username
  Set lvItem = lvInfo.ListItems.Add(, , "DBMS")
  lvItem.SubItems(1) = svr.dbVersion.Description
End Sub

Private Sub tvDatabases(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim dat As pgDatabase
  If Node.Children = 0 Then
    For Each dat In svr.Databases
      tv.Nodes.Add Node.Key, tvwChild, "DAT-" & GetID, dat.Identifier
    Next dat
  End If
  lvInfo.ColumnHeaders.Add , , "Database", 2000
  lvInfo.ColumnHeaders.Add , , "Owner", 1000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 3100
  For Each dat In svr.Databases
    Set lvItem = lvInfo.ListItems.Add(, , dat.Name)
    lvItem.SubItems(1) = dat.Owner
    lvItem.SubItems(2) = Replace(dat.Comment, vbCrLf, " ")
  Next dat
End Sub

Private Sub tvDatabase(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "SCH+" & GetID, "Schemas (" & svr.Databases(Node.Text).Namespaces.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "LNG+" & GetID, "Languages (" & svr.Databases(Node.Text).Languages.Count & ")"
  End If
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "Path")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Path
  Set lvItem = lvInfo.ListItems.Add(, , "Encoding")
  lvItem.SubItems(1) = svr.Databases(Node.Text).ServerEncoding
  Set lvItem = lvInfo.ListItems.Add(, , "System Database?")
  If svr.Databases(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = svr.Databases(Node.Text).Comment
  txtSQL.Text = svr.Databases(Node.Text).SQL
End Sub

Private Sub tvSchemas(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim nsp As pgNamespace
  If Node.Children = 0 Then
    For Each nsp In svr.Databases(Node.Parent.Text).Namespaces
      tv.Nodes.Add Node.Key, tvwChild, "SCH-" & GetID, nsp.Identifier
    Next nsp
  End If
  lvInfo.ColumnHeaders.Add , , "Schema", 2000
  lvInfo.ColumnHeaders.Add , , "Owner", 1000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 3100
  For Each nsp In svr.Databases(Node.Parent.Text).Namespaces
    Set lvItem = lvInfo.ListItems.Add(, , nsp.Name)
    lvItem.SubItems(1) = nsp.Owner
    lvItem.SubItems(2) = Replace(nsp.Comment, vbCrLf, " ")
  Next nsp
End Sub

Private Sub tvSchema(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "AGG+" & GetID, "Aggregates (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Aggregates.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "DOM+" & GetID, "Domains (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Domains.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "FNC+" & GetID, "Functions (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Functions.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "OPR+" & GetID, "Operators (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Operators.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "SEQ+" & GetID, "Sequences (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Sequences.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "TBL+" & GetID, "Tables (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Tables.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "TYP+" & GetID, "Types (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Types.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "VIE+" & GetID, "Views (" & svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Views.Count & ")"
  End If
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "ACL")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).ACL
  Set lvItem = lvInfo.ListItems.Add(, , "System Schema?")
  If svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).Comment
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Text).Namespaces(Node.Text).SQL
End Sub

Private Sub tvGroups(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
Dim grp As pgGroup
  If Node.Children = 0 Then
    For Each grp In svr.Groups
      tv.Nodes.Add Node.Key, tvwChild, "GRP-" & GetID, grp.Identifier, "group"
    Next grp
  End If
  lvInfo.ColumnHeaders.Add , , "Group", 2000
  lvInfo.ColumnHeaders.Add , , "Group ID", 1000
  lvInfo.ColumnHeaders.Add , , "Members", lvInfo.Width - 3100
  For Each grp In svr.Groups
    Set lvItem = lvInfo.ListItems.Add(, , grp.Name, "group", "group")
    lvItem.SubItems(1) = grp.ID
    szTemp = ""
    For Each vData In grp.Members
      szTemp = szTemp & vData & ", "
    Next vData
    If Len(szTemp) > 2 Then lvItem.SubItems(2) = Left(szTemp, Len(szTemp) - 2)
  Next grp
End Sub

Private Sub tvGroup(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Groups(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Groups(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Group ID")
  lvItem.SubItems(1) = svr.Groups(Node.Text).ID
  Set lvItem = lvInfo.ListItems.Add(, , "Member Count")
  lvItem.SubItems(1) = svr.Groups(Node.Text).Members.Count
  Set lvItem = lvInfo.ListItems.Add(, , "Members")
  For Each vData In svr.Groups(Node.Text).Members
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then lvItem.SubItems(1) = Left(szTemp, Len(szTemp) - 2)
End Sub

Private Sub tvUser(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim usr As pgUser
  lvInfo.ColumnHeaders.Add , , "Username", 2000
  lvInfo.ColumnHeaders.Add , , "User ID", 1500
  lvInfo.ColumnHeaders.Add , , "Account Expires", lvInfo.Width - 3600
  For Each usr In svr.Users
    Set lvItem = lvInfo.ListItems.Add(, , usr.Identifier, "user", "user")
    lvItem.SubItems(1) = usr.ID
    lvItem.SubItems(2) = usr.AccountExpires
  Next usr
End Sub

Private Sub tvAggregates(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim agg As pgAggregate
  If Node.Children = 0 Then
    For Each agg In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Aggregates
      tv.Nodes.Add Node.Key, tvwChild, "AGG-" & GetID, agg.Identifier
    Next agg
  End If
  lvInfo.ColumnHeaders.Add , , "Aggregate", 2000
  lvInfo.ColumnHeaders.Add , , "Input Type", 1000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 3100
  For Each agg In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Aggregates
    Set lvItem = lvInfo.ListItems.Add(, , agg.Name)
    lvItem.SubItems(1) = agg.InputType
    lvItem.SubItems(2) = Replace(agg.Comment, vbCrLf, " ")
  Next agg
End Sub

Private Sub tvAggregate(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "Input Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).InputType
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "State Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).StateType
  Set lvItem = lvInfo.ListItems.Add(, , "State Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).StateFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Final Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).FinalType
  Set lvItem = lvInfo.ListItems.Add(, , "Final Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).FinalFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Initial Condition")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).InitialCondition
  Set lvItem = lvInfo.ListItems.Add(, , "System Aggregate?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).Comment, vbCrLf, " ")
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Aggregates(Node.Text).SQL
End Sub

Private Sub tvDomains(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim dom As pgDomain
  If Node.Children = 0 Then
    For Each dom In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Domains
      tv.Nodes.Add Node.Key, tvwChild, "DOM-" & GetID, dom.Identifier
    Next dom
  End If
  lvInfo.ColumnHeaders.Add , , "Domain", 2000
  lvInfo.ColumnHeaders.Add , , "Base Type", 1000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 3100
  For Each dom In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Domains
    Set lvItem = lvInfo.ListItems.Add(, , dom.Name)
    lvItem.SubItems(1) = dom.BaseType
    lvItem.SubItems(2) = Replace(dom.Comment, vbCrLf, " ")
  Next dom
End Sub

Private Sub tvDomain(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "Input Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).BaseType
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "Default")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).Default
  Set lvItem = lvInfo.ListItems.Add(, , "Length")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).Length
  Set lvItem = lvInfo.ListItems.Add(, , "Numeric Scale")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).NumericScale
  Set lvItem = lvInfo.ListItems.Add(, , "Not Null?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).NotNull Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "System Domain?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).Comment, vbCrLf, " ")
    txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Domains(Node.Text).SQL
End Sub

Private Sub tvFunctions(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
Dim fnc As pgFunction
  If Node.Children = 0 Then
    For Each fnc In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Functions
      tv.Nodes.Add Node.Key, tvwChild, "FNC-" & GetID, fnc.Identifier
    Next fnc
  End If
  lvInfo.ColumnHeaders.Add , , "Function", 2000
  lvInfo.ColumnHeaders.Add , , "Arguments", 1500
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 3600
  For Each fnc In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Functions
    Set lvItem = lvInfo.ListItems.Add(, , fnc.Name)
    szTemp = ""
    For Each vData In fnc.Arguments
      szTemp = szTemp & vData & ", "
    Next vData
    If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
    lvItem.SubItems(1) = szTemp
    lvItem.SubItems(2) = Replace(fnc.Comment, vbCrLf, " ")
  Next fnc
End Sub

Private Sub tvFunction(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "Argument Count")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Arguments.Count
  Set lvItem = lvInfo.ListItems.Add(, , "Arguments")
  szTemp = ""
  For Each vData In svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Arguments
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  Set lvItem = lvInfo.ListItems.Add(, , "Returns")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Returns
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "Language")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Language
  Set lvItem = lvInfo.ListItems.Add(, , "Source")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Source
  Set lvItem = lvInfo.ListItems.Add(, , "Cachable?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Cachable Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Strict?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Strict Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "System Function?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).Comment, vbCrLf, " ")
    txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Functions(Node.Text).SQL
End Sub

Private Sub tvLanguages(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim lng As pgLanguage
  If Node.Children = 0 Then
    For Each lng In svr.Databases(Node.Parent.Text).Languages
      tv.Nodes.Add Node.Key, tvwChild, "LNG-" & GetID, lng.Identifier
    Next lng
  End If
  lvInfo.ColumnHeaders.Add , , "Language", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each lng In svr.Databases(Node.Parent.Text).Languages
    Set lvItem = lvInfo.ListItems.Add(, , lng.Name)
    lvItem.SubItems(1) = Replace(lng.Comment, vbCrLf, " ")
  Next lng
End Sub

Private Sub tvLanguage(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Handler")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Handler
  Set lvItem = lvInfo.ListItems.Add(, , "Trusted?")
  If svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Trusted Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "System Language?")
  If svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).Comment, vbCrLf, " ")
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Text).Languages(Node.Text).SQL
End Sub

Private Sub tvOperators(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim opr As pgOperator
  If Node.Children = 0 Then
    For Each opr In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Operators
      tv.Nodes.Add Node.Key, tvwChild, "OPR-" & GetID, opr.Identifier
    Next opr
  End If
  lvInfo.ColumnHeaders.Add , , "Operator", 2000
  lvInfo.ColumnHeaders.Add , , "Left Type", 1000
  lvInfo.ColumnHeaders.Add , , "Right Type", 1000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 4100
  For Each opr In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Operators
    Set lvItem = lvInfo.ListItems.Add(, , opr.Name)
    lvItem.SubItems(1) = opr.LeftOperandType
    lvItem.SubItems(2) = opr.RightOperandType
    lvItem.SubItems(3) = Replace(opr.Comment, vbCrLf, " ")
  Next opr
End Sub

Private Sub tvOperator(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "Left Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).LeftOperandType
  Set lvItem = lvInfo.ListItems.Add(, , "Right Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).RightOperandType
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "Operator Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).OperatorFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Join Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).JoinFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Restrict Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).RestrictFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Result Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).ResultType
  Set lvItem = lvInfo.ListItems.Add(, , "Commutator")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).Commutator
  Set lvItem = lvInfo.ListItems.Add(, , "Negator")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).Negator
  Set lvItem = lvInfo.ListItems.Add(, , "Kind")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).Kind
  Set lvItem = lvInfo.ListItems.Add(, , "Left Sort Operator")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).LeftTypeSortOperator
  Set lvItem = lvInfo.ListItems.Add(, , "Right Sort Operator")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).rightTypeSortOperator
  Set lvItem = lvInfo.ListItems.Add(, , "Hash Joins?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).HashJoins Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "System Operator?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).Comment, vbCrLf, " ")
    txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Operators(Node.Text).SQL
End Sub

Private Sub tvSequences(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim seq As pgSequence
  If Node.Children = 0 Then
    For Each seq In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Sequences
      tv.Nodes.Add Node.Key, tvwChild, "SEQ-" & GetID, seq.Identifier
    Next seq
  End If
  lvInfo.ColumnHeaders.Add , , "Sequence", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each seq In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Sequences
    Set lvItem = lvInfo.ListItems.Add(, , seq.Name)
    lvItem.SubItems(1) = Replace(seq.Comment, vbCrLf, " ")
  Next seq
End Sub

Private Sub tvSequence(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "ACL")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).ACL
  Set lvItem = lvInfo.ListItems.Add(, , "Last Value")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).LastValue
  Set lvItem = lvInfo.ListItems.Add(, , "Minimum")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Minimum
  Set lvItem = lvInfo.ListItems.Add(, , "Maximum")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).maximum
  Set lvItem = lvInfo.ListItems.Add(, , "Increment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Increment
  Set lvItem = lvInfo.ListItems.Add(, , "Cache")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Cache
  Set lvItem = lvInfo.ListItems.Add(, , "Cycled?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Cycled Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "System Sequence?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).Comment, vbCrLf, " ")
    txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Sequences(Node.Text).SQL
End Sub

Private Sub tvTables(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim tbl As pgTable
  If Node.Children = 0 Then
    For Each tbl In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Tables
      tv.Nodes.Add Node.Key, tvwChild, "TBL-" & GetID, tbl.Identifier
    Next tbl
  End If
  lvInfo.ColumnHeaders.Add , , "Table", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each tbl In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Tables
    Set lvItem = lvInfo.ListItems.Add(, , tbl.Name)
    lvItem.SubItems(1) = tbl.Comment
  Next tbl
End Sub

Private Sub tvTable(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  If Node.Children = 0 Then
    tv.Nodes.Add Node.Key, tvwChild, "CHK+" & GetID, "Checks (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Checks.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "COL+" & GetID, "Columns (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Columns.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "FKY+" & GetID, "Foreign Keys (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).ForeignKeys.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "IND+" & GetID, "Indexes (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).indexes.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "RUL+" & GetID, "Rules (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).rules.Count & ")"
    tv.Nodes.Add Node.Key, tvwChild, "TRG+" & GetID, "Triggers (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Triggers.Count & ")"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "ACL")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).ACL
  Set lvItem = lvInfo.ListItems.Add(, , "Rows")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Rows
  Set lvItem = lvInfo.ListItems.Add(, , "Inherited Tables Count")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).InheritedTables.Count
  Set lvItem = lvInfo.ListItems.Add(, , "Inherited Tables")
  For Each vData In svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).InheritedTables
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  lvItem.SubItems(1) = szTemp
  Set lvItem = lvInfo.ListItems.Add(, , "System Table?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).Comment, vbCrLf, " ")
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).SQL
End Sub

Private Sub tvChecks(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim chk As pgCheck
  If Node.Children = 0 Then
    For Each chk In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks
      tv.Nodes.Add Node.Key, tvwChild, "CHK-" & GetID, chk.Identifier
    Next chk
  End If
  lvInfo.ColumnHeaders.Add , , "Check", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each chk In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Checks
    Set lvItem = lvInfo.ListItems.Add(, , chk.Name)
    lvItem.SubItems(1) = chk.Definition
  Next chk
End Sub

Private Sub tvCheck(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "Definition")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text).Definition
  Set lvItem = lvInfo.ListItems.Add(, , "System Check?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Checks(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
End Sub

Private Sub tvColumns(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim col As pgColumn
  If Node.Children = 0 Then
    For Each col In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns
     tv.Nodes.Add Node.Key, tvwChild, "COL-" & GetID, col.Identifier
    Next col
  End If
  lvInfo.ColumnHeaders.Add , , "Column", 2000
  lvInfo.ColumnHeaders.Add , , "Type", 1000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 3100
  For Each col In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Columns
    Set lvItem = lvInfo.ListItems.Add(, , col.Name)
    lvItem.SubItems(1) = col.DataType
    lvItem.SubItems(2) = col.Comment
  Next col
End Sub

Private Sub tvColumn(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Position")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Position
  Set lvItem = lvInfo.ListItems.Add(, , "Data Type")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).DataType
  Set lvItem = lvInfo.ListItems.Add(, , "Size")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Length
  Set lvItem = lvInfo.ListItems.Add(, , "Scale")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).NumericScale
  Set lvItem = lvInfo.ListItems.Add(, , "Default")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Default
  Set lvItem = lvInfo.ListItems.Add(, , "Restrict Nulls?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).NotNull Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "System Column?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Columns(Node.Text).Comment
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Tables(Node.Text).SQL
End Sub

Private Sub tvForeignKeys(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim fky As pgForeignKey
    If Node.Children = 0 Then
      For Each fky In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys
      tv.Nodes.Add Node.Key, tvwChild, "FKY-" & GetID, fky.Identifier
    Next fky
  End If
  lvInfo.ColumnHeaders.Add , , "Foreign Key", 2000
  lvInfo.ColumnHeaders.Add , , "References", lvInfo.Width - 2100
  For Each fky In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).ForeignKeys
    Set lvItem = lvInfo.ListItems.Add(, , fky.Name)
    lvItem.SubItems(1) = fky.ReferencedTable
  Next fky
End Sub

Private Sub tvForeignKey(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  If Node.Children = 0 Then tv.Nodes.Add Node.Key, tvwChild, "REL+" & GetID, "Relationships (" & svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).Relationships.Count & ")"
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "References")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).ReferencedTable
  Set lvItem = lvInfo.ListItems.Add(, , "System Foreign Key?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).ForeignKeys(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
End Sub

Private Sub tvRelationships(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim rel As pgRelationship
  lvInfo.ColumnHeaders.Add , , "Local Column", 2000
  lvInfo.ColumnHeaders.Add , , "Referenced Column", lvInfo.Width - 2600
  For Each rel In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Parent.Text).ForeignKeys(Node.Parent.Text).Relationships
    Set lvItem = lvInfo.ListItems.Add(, , rel.LocalColumn)
    lvItem.SubItems(1) = rel.ReferencedColumn
  Next rel
End Sub

Private Sub tvIndexes(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim ind As pgIndex
  If Node.Children = 0 Then
    For Each ind In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).indexes
      tv.Nodes.Add Node.Key, tvwChild, "IND-" & GetID, ind.Identifier
    Next ind
  End If
  lvInfo.ColumnHeaders.Add , , "Index", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each ind In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).indexes
    Set lvItem = lvInfo.ListItems.Add(, , ind.Name)
    lvItem.SubItems(1) = ind.Comment
  Next ind
End Sub

Private Sub tvIndex(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Unique?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).Unique Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Primary?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).Primary Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Column Count")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).IndexedColumns.Count
  For Each vData In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).IndexedColumns
    szTemp = szTemp & vData & ", "
  Next vData
  If Len(szTemp) > 2 Then szTemp = Left(szTemp, Len(szTemp) - 2)
  Set lvItem = lvInfo.ListItems.Add(, , "Columns")
  lvItem.SubItems(1) = szTemp
  Set lvItem = lvInfo.ListItems.Add(, , "System Index?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).Comment
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).indexes(Node.Text).SQL
End Sub

Private Sub tvRules(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim rul As pgRule
  If Node.Children = 0 Then
    For Each rul In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).rules
      tv.Nodes.Add Node.Key, tvwChild, "RUL-" & GetID, rul.Identifier
    Next rul
  End If
  lvInfo.ColumnHeaders.Add , , "Rule", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each rul In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).rules
    Set lvItem = lvInfo.ListItems.Add(, , rul.Name)
    lvItem.SubItems(1) = rul.Comment
  Next rul
End Sub

Private Sub tvRule(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).rules(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).rules(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Definition")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).rules(Node.Text).Definition
  Set lvItem = lvInfo.ListItems.Add(, , "System Rule?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).rules(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).rules(Node.Text).Comment
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).rules(Node.Text).SQL
End Sub

Private Sub tvTriggers(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim trg As pgTrigger
  If Node.Children = 0 Then
    For Each trg In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers
      tv.Nodes.Add Node.Key, tvwChild, "TRG-" & GetID, trg.Identifier
    Next trg
  End If
  lvInfo.ColumnHeaders.Add , , "Trigger", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each trg In svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Text).Tables(Node.Parent.Text).Triggers
    Set lvItem = lvInfo.ListItems.Add(, , trg.Name)
    lvItem.SubItems(1) = trg.Comment
  Next trg
End Sub

Private Sub tvTrigger(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Executes")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).Executes
  Set lvItem = lvInfo.ListItems.Add(, , "Event")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).TriggerEvent
  Set lvItem = lvInfo.ListItems.Add(, , "For Each")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).ForEach
  Set lvItem = lvInfo.ListItems.Add(, , "Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).TriggerFunction
  Set lvItem = lvInfo.ListItems.Add(, , "System Trigger?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).Comment
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Parent.Parent.Text).Tables(Node.Parent.Parent.Text).Triggers(Node.Text).SQL
End Sub

Private Sub tvTypes(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim typ As pgType
  If Node.Children = 0 Then
    For Each typ In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Types
      tv.Nodes.Add Node.Key, tvwChild, "TYP-" & GetID, typ.Identifier
    Next typ
  End If
  lvInfo.ColumnHeaders.Add , , "Type", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each typ In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Types
    Set lvItem = lvInfo.ListItems.Add(, , typ.Name)
    lvItem.SubItems(1) = typ.Comment
  Next typ
End Sub

Private Sub tvType(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "Input Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).InputFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Output Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).OutputFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Internal Length")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).InternalLength
  Set lvItem = lvInfo.ListItems.Add(, , "External Length")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).ExternalLength
  Set lvItem = lvInfo.ListItems.Add(, , "Default")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Default
  Set lvItem = lvInfo.ListItems.Add(, , "Element")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Element
  Set lvItem = lvInfo.ListItems.Add(, , "Delimiter")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).delimiter
  Set lvItem = lvInfo.ListItems.Add(, , "Send Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).SendFunction
  Set lvItem = lvInfo.ListItems.Add(, , "Receive Function")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).ReceiveFunction
    Set lvItem = lvInfo.ListItems.Add(, , "Passed by Value?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).PassedByValue Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Alignment")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Alignment
  Set lvItem = lvInfo.ListItems.Add(, , "Storage")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Storage
  Set lvItem = lvInfo.ListItems.Add(, , "System Type?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).Comment, vbCrLf, " ")
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Types(Node.Text).SQL
End Sub

Private Sub tvViews(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  Dim vie As pgView
  If Node.Children = 0 Then
    For Each vie In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Views
      tv.Nodes.Add Node.Key, tvwChild, "VIE-" & GetID, vie.Identifier
    Next vie
  End If
  lvInfo.ColumnHeaders.Add , , "View", 2000
  lvInfo.ColumnHeaders.Add , , "Comment", lvInfo.Width - 2100
  For Each vie In svr.Databases(Node.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Text).Views
    Set lvItem = lvInfo.ListItems.Add(, , vie.Name)
    lvItem.SubItems(1) = vie.Comment
  Next vie
End Sub

Private Sub tvView(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
  lvInfo.ColumnHeaders.Add , , "Property", 2000
  lvInfo.ColumnHeaders.Add , , "Value", lvInfo.Width - 2100
  Set lvItem = lvInfo.ListItems.Add(, , "Name")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).Name
  Set lvItem = lvInfo.ListItems.Add(, , "OID")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).OID
  Set lvItem = lvInfo.ListItems.Add(, , "Owner")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).Owner
  Set lvItem = lvInfo.ListItems.Add(, , "ACL")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).ACL
  Set lvItem = lvInfo.ListItems.Add(, , "Definition")
  lvItem.SubItems(1) = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).Definition
  Set lvItem = lvInfo.ListItems.Add(, , "System View?")
  If svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).SystemObject Then
    lvItem.SubItems(1) = "Yes"
  Else
    lvItem.SubItems(1) = "No"
  End If
  Set lvItem = lvInfo.ListItems.Add(, , "Comment")
  lvItem.SubItems(1) = Replace(svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).Comment, vbCrLf, " ")
  txtSQL.Text = svr.Databases(Node.Parent.Parent.Parent.Parent.Text).Namespaces(Node.Parent.Parent.Text).Views(Node.Text).SQL
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
Dim lvItem As ListItem
Dim szTemp As String
Dim vData As Variant

  lvInfo.ColumnHeaders.Clear
  lvInfo.ListItems.Clear
  
  Select Case Left(Node.Key, 4)

    Case "SVR-" 'Server
      tvServer Node
    
    Case "DAT+" 'Databases
      tvDatabases Node
        
    Case "DAT-" 'Database
      tvDatabase Node
      
    Case "SCH+" 'Schemas
      tvSchemas Node
        
    Case "SCH-" 'Schema
      tvSchema Node
      
    Case "GRP+" 'Groups
      tvGroups Node
      
    Case "GRP-" 'Group
      tvGroup Node
      
    Case "USR+" 'Users
      tvUser Node
      
    Case "AGG+" 'Aggregates
      tvAggregates Node
      
    Case "AGG-" 'Aggregate
      tvAggregate Node
      
    Case "DOM+" 'Domains
      tvDomains Node
      
    Case "DOM-" 'Domains
      tvDomain Node
      
    Case "FNC+" 'Functions
      tvFunctions Node
      
    Case "FNC-" 'Function
      tvFunction Node
      
    Case "LNG+" 'Languages
      tvLanguages Node

    Case "LNG-" 'Language
      tvLanguage Node
      
    Case "OPR+" 'Operators
      tvOperators Node
      
    Case "OPR-" 'Operator
      tvOperator Node
      
    Case "SEQ+" 'Sequences
      tvSequences Node

    Case "SEQ-" 'Sequence
      tvSequence Node
      
    Case "TBL+" 'Tables
      tvTables Node
      
    Case "TBL-" 'Table
      tvTable Node
      
    Case "CHK+" 'Checks
      tvChecks Node
      
    Case "CHK-" 'Check
      tvCheck Node
    
    Case "COL+" 'Columns
      tvColumns Node
      
    Case "COL-" 'Column
      tvColumn Node
      
    Case "FKY+" 'Foreign Keys
      tvForeignKeys Node
      
    Case "FKY-" 'Foreign Key
      tvForeignKey Node
      
    Case "REL+" 'Relationships
      tvRelationships Node
      
    Case "IND+" 'Indexes
      tvIndexes Node
      
    Case "IND-" 'Index
      tvIndex Node

    Case "RUL+" 'Rules
      tvRules Node
  
    Case "RUL-" 'Rule
      tvRule Node
      
    Case "TRG+" 'Triggers
      tvTriggers Node
      
    Case "TRG-" 'Trigger
      tvTrigger Node
      
    Case "TYP+" 'Types
      tvTypes Node

    Case "TYP-" 'Type
      tvType Node
      
    Case "VIE+" 'Views
      tvViews Node
      
    Case "VIE-" 'View
      tvView Node
      
  End Select
  Node.Expanded = Not Node.Expanded
End Sub
