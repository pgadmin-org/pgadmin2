VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Object"
   ClientHeight    =   2112
   ClientLeft      =   7908
   ClientTop       =   1980
   ClientWidth     =   2868
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2112
   ScaleWidth      =   2868
   Begin MSComctlLib.ImageList il 
      Left            =   0
      Top             =   1440
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":0000
            Key             =   "aggregate"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":06D2
            Key             =   "check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":0DA4
            Key             =   "column"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":1476
            Key             =   "function"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":1B48
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":221A
            Key             =   "index"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":27B4
            Key             =   "indexcolumn"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":2E86
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":3558
            Key             =   "language"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":3C2A
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":42FC
            Key             =   "property"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":4896
            Key             =   "relationship"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":49F0
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":50C2
            Key             =   "server"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":521C
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":58EE
            Key             =   "table"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":5FC0
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":6692
            Key             =   "type"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":6D64
            Key             =   "user"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":6EBE
            Key             =   "view"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":7590
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":7C62
            Key             =   "namespace"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":8834
            Key             =   "cast"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":9406
            Key             =   "conversion"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClone.frx":9CE0
            Key             =   "operatorclass"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOption 
      Caption         =   "Paste"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton optPaste 
         Caption         =   "Structure and data"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optPaste 
         Caption         =   "Structure only"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClone 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "The name of the column."
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Enter a name for the new object."
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmClone.frm - Clone object database

Option Explicit

Private ObjDbClone

Public Sub Initialise(ByVal ObjClone)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmClone.Initialise(" & QUOTE & ObjClone.ObjectType & QUOTE & ")", etFullDebug

  PatchForm Me
  
  Me.Icon = il.ListImages(LCase(ObjClone.ObjectType)).Picture
  
  Set ObjDbClone = ObjClone
  txtNewName.Text = ObjClone.Name
  
  'verify copy data table
  Select Case ObjClone.ObjectType
    Case "Table"
      If ObjDbClone.Database <> ctx.CurrentDB Then
        optPaste(1).Enabled = False
        MsgBox §§TrasLang§§("Data can only be copied within the same database!"), vbExclamation, §§TrasLang§§("Error")
      End If
      
    Case "Cast"
      'cast not have a name object
      optPaste(1).Enabled = False
      txtNewName.Locked = True
      txtNewName.BackColor = &H8000000F
    
    Case Else
      optPaste(1).Enabled = False
  
  End Select
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmClone.Initialise"
End Sub

Private Sub cmdClone_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmClone.cmdClone_Click", etFullDebug

Dim objTmp
Dim objNode As Node
Dim szArguments As String
Dim vData
Dim szTemp As String

  If Len(Trim(txtNewName.Text)) = 0 Then
    MsgBox §§TrasLang§§("The name you have entered is not valid!"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If
    
  'verify if name Exists
  Select Case ObjDbClone.ObjectType
    Case "Domain", "Table", "View", "Function", "Aggregate", "Operator", "Type", "Conversion", "OperatorClass"
      
      Select Case ObjDbClone.ObjectType
        Case "OperatorClass"
          szTemp = "OperatorsClass"
          
        Case Else
          szTemp = ObjDbClone.ObjectType & "s"
      End Select
      
      Set objTmp = CallByName(frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS), szTemp, VbGet)
    
    Case "Group", "User"
      Set objTmp = CallByName(frmMain.svr, ObjDbClone.ObjectType & "s", VbGet)
      
    Case "Cast"
      Set objTmp = CallByName(frmMain.svr.Databases(ctx.CurrentDB), ObjDbClone.ObjectType & "s", VbGet)
  
  End Select
  
  If objTmp.Exists(txtNewName.Text) Then
    MsgBox §§TrasLang§§("An object named ") & txtNewName.Text & §§TrasLang§§(" already exists"), vbExclamation, §§TrasLang§§("Error")
    Exit Sub
  End If
  
  StartMsg §§TrasLang§§("Copying ") & ObjDbClone.ObjectType & "..."
  'create new object
  Select Case ObjDbClone.ObjectType
    Case "Type"
      Set objTmp = CloneType(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Types.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "TYP-" & GetID, txtNewName.Text, "type")
      objNode.Text = §§TrasLang§§("Types (") & objNode.Children & ")"
    
    Case "Conversion"
      Set objTmp = CloneConversion(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Conversions.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "CNV-" & GetID, txtNewName.Text, "conversion")
      objNode.Text = §§TrasLang§§("Conversions (") & objNode.Children & ")"

    Case "Cast"
      Set objTmp = CloneCast(ctx.CurrentDB)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Casts.Tag
      frmMain.tv.Nodes.Add objNode.Key, tvwChild, "CST-" & GetID, objTmp.Identifier, "cast"
      objNode.Text = §§TrasLang§§("Casts (") & objNode.Children & ")"
    
    Case "Domain"
      Set objTmp = CloneDomain(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Domains.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "DOM-" & GetID, txtNewName.Text, "domain")
      objNode.Text = §§TrasLang§§("Domains (") & objNode.Children & ")"
      
    Case "Operator"
      Set objTmp = CloneOperator(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Operators.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "OPR-" & GetID, txtNewName.Text & " (" & ObjDbClone.LeftOperandType & ", " & ObjDbClone.RightOperandType & ")", "operator")
      objNode.Text = §§TrasLang§§("Operators (") & objNode.Children & ")"

    Case "OperatorClass"
      Set objTmp = CloneOperatorClass(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).OperatorsClass.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "OPC-" & GetID, txtNewName.Text & " (" & ObjDbClone.AccessMethod & ")", "operatorclass")
      objNode.Text = §§TrasLang§§("Operators Class (") & objNode.Children & ")"
    
    Case "Aggregate"
      Set objTmp = CloneAggregate(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
      
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Aggregates.Tag
      If ObjDbClone.InputType = "ANY" Then
        frmMain.tv.Nodes.Add objNode.Key, tvwChild, "AGG-" & GetID, txtNewName.Text & " opaque", "aggregate"
      Else
        frmMain.tv.Nodes.Add objNode.Key, tvwChild, "AGG-" & GetID, txtNewName.Text & " " & ObjDbClone.InputType, "aggregate"
      End If
      objNode.Text = §§TrasLang§§("Aggregates (") & objNode.Children & ")"
    
    Case "Function"
      Set objTmp = CloneFunction(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Get the identifier/arguments in case we need it
      For Each vData In ObjDbClone.Arguments
        szArguments = szArguments & vData & ", "
      Next
      If Len(szArguments) > 2 Then szArguments = Left(szArguments, Len(szArguments) - 2)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Functions.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "FNC-" & GetID, txtNewName.Text & "(" & szArguments & ")", "function")
      objNode.Text = §§TrasLang§§("Functions (") & objNode.Children & ")"
    
    Case "Table"
      Set objTmp = CloneTable(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS, optPaste(1).Value)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Tables.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "TBL-" & GetID, txtNewName.Text, "table")
      objNode.Text = §§TrasLang§§("Tables (") & objNode.Children & ")"
      
      MsgBox §§TrasLang§§("Please verify the checks, foreign keys, rules and triggers!"), vbInformation
    
    Case "View"
      Set objTmp = CloneView(txtNewName.Text, ctx.CurrentDB, ctx.CurrentNS)
    
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Databases(ctx.CurrentDB).Namespaces(ctx.CurrentNS).Views.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "VIE-" & GetID, txtNewName.Text, "view")
      objNode.Text = §§TrasLang§§("Views (") & objNode.Children & ")"
    
    Case "Group"
      Set objTmp = CloneGroup(txtNewName.Text)

      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Groups.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "GRP-" & GetID, txtNewName.Text, "group")
      objNode.Text = §§TrasLang§§("Groups (") & frmMain.svr.Groups.Count & ")"
  
    Case "User"
      Set objTmp = CloneUser(txtNewName.Text)
      
      'Add a new node and update the text on the parent
      Set objNode = frmMain.svr.Users.Tag
      Set objTmp.Tag = frmMain.tv.Nodes.Add(objNode.Key, tvwChild, "USR-" & GetID, txtNewName.Text, "user")
      objNode.Text = §§TrasLang§§("Users (") & frmMain.svr.Users.Count & ")"
      
      MsgBox §§TrasLang§§("The password for the new user is blank!"), vbInformation
  
  End Select
  
  'Simulate a node click to refresh
  frmMain.tv_NodeClick frmMain.tv.SelectedItem
  
  EndMsg
  Unload Me
  
  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmClone.cmdClone_Click"
End Sub

