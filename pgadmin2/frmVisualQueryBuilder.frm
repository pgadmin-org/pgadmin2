VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{44F33AC4-8757-4330-B063-18608617F23E}#12.4#0"; "HighlightBox.ocx"
Begin VB.Form frmVisualQueryBuilder 
   Caption         =   "Visual Query Builder"
   ClientHeight    =   6492
   ClientLeft      =   2004
   ClientTop       =   2088
   ClientWidth     =   9132
   Icon            =   "frmVisualQueryBuilder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6492
   ScaleWidth      =   9132
   Visible         =   0   'False
   Begin HighlightBox.HBX txtSQL 
      Height          =   1272
      Left            =   4200
      TabIndex        =   7
      Top             =   60
      Width           =   4872
      _ExtentX        =   8594
      _ExtentY        =   2244
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SQL Query"
   End
   Begin VB.Frame fraTypeQuery 
      Caption         =   "Type Query"
      Height          =   1332
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   1272
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Insert"
         Height          =   192
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1020
         Width           =   1032
      End
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Delete"
         Height          =   192
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   780
         Width           =   1032
      End
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Update"
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1032
      End
      Begin VB.OptionButton OptTypeQuery 
         Caption         =   "&Select"
         Height          =   192
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1032
      End
   End
   Begin MSComctlLib.ImageList il 
      Left            =   1920
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":0BC2
            Key             =   "table"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":1294
            Key             =   "namespace"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":1E66
            Key             =   "view"
         EndProperty
      EndProperty
   End
   Begin pgAdmin2.RelationObj RelQuery 
      Height          =   5052
      Left            =   60
      TabIndex        =   0
      Top             =   1380
      Width           =   9012
      _extentx        =   17590
      _extenty        =   8911
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1260
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Add table to selection"
      Top             =   60
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   2223
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "il"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Menu mnuRelQuery 
      Caption         =   "Relation Query"
      Visible         =   0   'False
      Begin VB.Menu mnuRelQueryViewSql 
         Caption         =   "View SQL"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRetunQuery 
         Caption         =   "Retun Query"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmVisualQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmVisualQueryBuilder.frm - Visual Query Builder

Option Explicit

Dim szDB As String
Dim WithEvents FGridVQB As MSFlexGridLib.MSFlexGrid
Attribute FGridVQB.VB_VarHelpID = -1
Dim WithEvents objFGrid As ClsSuperFGrid
Attribute objFGrid.VB_VarHelpID = -1
Dim frmCallingForm As Form

Public Sub Initialise(Database As String, frmCF As Form)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Initialise(" & QUOTE & Database & QUOTE & ")", etFullDebug

Dim objNS As pgNamespace
Dim objTable As pgTable
Dim NodeNs As Node
Dim ii As Integer
  
  PatchForm Me
  
  szDB = Database
  Set frmCallingForm = frmCF
  txtSQL.Wordlist = ctx.AutoHighlight
  
  'load structure
  tv.Nodes.Clear
  For Each objNS In frmMain.svr.Databases(szDB).Namespaces
    If Not (objNS.SystemObject And Not ctx.IncludeSys) Then
      If objNS.Tables.Count(Not ctx.IncludeSys) > 0 Then
        Set NodeNs = tv.Nodes.Add(, , "NSP-" & GetID, objNS.Identifier, "namespace")
        For Each objTable In objNS.Tables
          If Not (objTable.SystemObject And Not ctx.IncludeSys) Then
            tv.Nodes.Add NodeNs.Key, tvwChild, "TBL-" & GetID, objTable.Identifier, "table"
          End If
        Next
      End If
    End If
  Next
  
  With RelQuery
    .MenuActionGrid(1).Caption = "Delete"
    .MenuActionGridEnable = True
  End With
  
  Set FGridVQB = RelQuery.GetGridCompose
  With FGridVQB
    .Visible = True
    .FixedRows = 1
    .FixedCols = 1
    .Rows = 6
    .Cols = 64
    .Height = Me.TextHeight("0") * (.Rows + 4)
    .RowHeight(0) = 100
    .TextMatrix(1, 0) = "Table"
    .TextMatrix(2, 0) = "Column"
    .TextMatrix(3, 0) = "Order"
    .TextMatrix(4, 0) = "Visible"
    .TextMatrix(5, 0) = "Where"
    .HighLight = flexHighlightNever
  End With
  AutoSizeColumnFGrid FGridVQB

  'fix grid using super grid
  Set objFGrid = New ClsSuperFGrid
  Set objFGrid.FlexGrid = FGridVQB
  For ii = 1 To FGridVQB.Cols - 1
    With objFGrid
      .AddAction ii, 3, TAFG_COMBO, "|0|1|Ascending|1|0|Discending|2|0"
      .AddAction ii, 4, TAFG_CHECK, "UNCHECKED"
      .AddAction ii, 5, TAFG_TEXT
      With .FlexGrid
        .Col = ii
        .Row = 5
        .CellAlignment = flexAlignLeftCenter
      End With
    End With
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Initialise"
End Sub

Private Sub Form_Resize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Form_Resize()", etFullDebug
  
  If Me.WindowState <> vbMinimized Then
    tv.Left = RelQuery.Left
    RelQuery.Width = Me.ScaleWidth - 100
    
    If txtSQL.Top = 0 Then
      txtSQL.Width = Me.ScaleWidth
      txtSQL.Height = Me.ScaleHeight
    Else
      If Me.ScaleWidth - txtSQL.Left > 0 Then txtSQL.Width = Me.ScaleWidth - txtSQL.Left - 50
    End If
    RelQuery.Height = Me.ScaleHeight - tv.Top - tv.Height - 100
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Form_Resize"
End Sub

Private Sub mnuFileExit_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileExit_Click()", etFullDebug

  Unload Me
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileExit_Click"
End Sub

Private Sub mnuFileRetunQuery_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuFileRetunQuery_Click()", etFullDebug

  If Not frmCallingForm Is Nothing Then
    If Not frmCallingForm.Visible Then
      MsgBox "The form that called this form has been destroyed!", vbExclamation, "Error"
      Exit Sub
    End If
  End If
  
  mnuRelQueryViewSql_Click
  frmCallingForm.txtSQL.Text = txtSQL.Text
  frmCallingForm.ZOrder
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuFileRetunQuery_Click"
End Sub

Private Sub RelQuery_MenuActionGridCompose(Index As Integer, Col As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_MenuActionGridCompose(" & Index & "," & Col & ")", etFullDebug

Dim iCol As Integer

  Select Case Index
    
    Case 1    'remove action
      objFGrid.RemoveColumn FGridVQB.ColSel
      For iCol = 1 To FGridVQB.Cols - 1
        If Len(FGridVQB.TextMatrix(1, iCol)) = 0 Then
          objFGrid.RemoveAction iCol, 1, TAFG_COMBO
          objFGrid.RemoveAction iCol, 2, TAFG_COMBO
        End If
      Next
  
  End Select
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_RemoveElementinGridCompose"
End Sub

Private Sub RelQuery_RenameRelation(OldName As String, NewName As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_RenameRelation(" & QUOTE & OldName & QUOTE & "," & QUOTE & NewName & QUOTE & ")", etFullDebug

Dim iCol As Integer
  
  'change name in flexgrid
  For iCol = 1 To FGridVQB.Cols - 1
    If FGridVQB.TextMatrix(1, iCol) = OldName Then
      FGridVQB.TextMatrix(1, iCol) = NewName
    End If
  Next
  
  RebildComboTable
  AutoSizeColumnFGrid FGridVQB
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_RenameRelation"
End Sub

'load table relation
Private Sub tv_DblClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.tv_DblClick()", etFullDebug

Dim szTemp As String
Dim objColumn As pgColumn
Dim tmpCol As Collection
Dim szTable As String
Dim szNamespace As String

  If tv.SelectedItem Is Nothing Then Exit Sub
  If Left(tv.SelectedItem.Key, 3) <> "TBL" Then Exit Sub
  
  szTable = tv.SelectedItem.Text
  szNamespace = tv.SelectedItem.Parent.Text
  
  'load filed table
  szTemp = ""
  For Each objColumn In frmMain.svr.Databases(szDB).Namespaces(szNamespace).Tables(szTable).Columns
    If Not (objColumn.SystemObject And Not ctx.IncludeSys) Then
      If Len(szTemp) > 0 Then szTemp = szTemp & "|"
      szTemp = szTemp & objColumn.Name
    End If
  Next
  
  'add new table in relation
  RelQuery.AddElement szTable, szNamespace & "." & szTable, "Schema: " & szNamespace & "  Table: " & szTable, Split(szTemp, "|")

  RebildComboTable
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.tv_DblClick"
End Sub

'rebuild any combo table in flexgrid
Private Sub RebildComboTable()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RebildComboTable()", etFullDebug

Dim iCol As Integer
Dim colRel As Collection
Dim szTemp As String
Dim vData
  
  'create string table
  Set colRel = RelQuery.GetRelation
  szTemp = ""
  For Each vData In colRel
    vData = Split(vData, ",")
    If Len(szTemp) > 0 Then szTemp = szTemp & "|"
    szTemp = szTemp & vData(0) & "|0|0"
  Next
  
  'remove and add combo table
  For iCol = 1 To FGridVQB.Cols - 1
    If Len(FGridVQB.TextMatrix(1, iCol)) > 0 Then
      objFGrid.RemoveAction iCol, 1, TAFG_COMBO
      objFGrid.AddAction iCol, 1, TAFG_COMBO, szTemp
    End If
  Next
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RebildComboTable"
End Sub

Private Sub RelQuery_AddElementInGridCompose(Col As Integer, Name As String, Tag As String, Element As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.tv_DblClick(" & Col & "," & QUOTE & Name & QUOTE & "," & QUOTE & Tag & QUOTE & "," & QUOTE & Element & QUOTE & ")", etFullDebug
  
Dim objTable As pgTable
Dim szTemp As String
Dim vData
Dim ii As Integer
Dim colRel As Collection
  
  If Col <= 0 Then Exit Sub
  With FGridVQB
    If Len(.TextMatrix(1, Col)) > 0 Then
      'add new column
      objFGrid.AddColumn Col
    End If
    .TextMatrix(1, Col) = Name
    .TextMatrix(2, Col) = Element
  End With
  objFGrid.SetCheckBoxes Col, 4, True

  LoadComboColumn Name, Col

  'load combo table
  Set colRel = RelQuery.GetRelation
  szTemp = ""
  For Each vData In colRel
    vData = Split(vData, ",")
    If Len(szTemp) > 0 Then szTemp = szTemp & "|"
    szTemp = szTemp & vData(0) & "|0|0"
  Next
  objFGrid.RemoveAction Col, 1, TAFG_COMBO
  objFGrid.AddAction Col, 1, TAFG_COMBO, szTemp
  AutoSizeColumnFGrid FGridVQB
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_AddElementInGridCompose"
End Sub

Private Sub LoadComboColumn(Name As String, Col As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.LoadComboColumn(" & QUOTE & Name & QUOTE & "," & Col & ")", etFullDebug

Dim objColumn As pgColumn
Dim szTemp As String
Dim bFound As Boolean

Dim szNamespace As String
Dim szTable As String
Dim colRel As Collection
Dim iCol As Integer
Dim vData

  'create string table
  Set colRel = RelQuery.GetRelation
  szTemp = ""
  For Each vData In colRel
    vData = Split(vData, ",")
    If vData(0) = Name Then
      vData = Split(vData(1), ".")
      szNamespace = vData(0)
      szTable = vData(1)
      Exit For
    End If
    If Len(szTemp) > 0 Then szTemp = szTemp & "|"
    szTemp = szTemp & vData(0) & "|0|0"
  Next

  'load combo column
  bFound = False
  szTemp = "*|0|0"
  For Each objColumn In frmMain.svr.Databases(szDB).Namespaces(szNamespace).Tables(szTable).Columns
    If Not (objColumn.SystemObject And Not ctx.IncludeSys) Then
      szTemp = szTemp & "|" & objColumn.Name & "|0|0"
      If objColumn.Name = FGridVQB.TextMatrix(2, Col) Then bFound = True
    End If
  Next

  'if not found column insert default * (all)
  If Not bFound Then FGridVQB.TextMatrix(2, Col) = "*"

  objFGrid.RemoveAction Col, 2, TAFG_COMBO
  objFGrid.AddAction Col, 2, TAFG_COMBO, szTemp
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.LoadComboColumn"
End Sub

Private Sub objFGrid_ComboChange(Col As Integer, Row As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.objFGrid_ComboChange(" & Col & "," & Row & ")", etFullDebug

  If Row = 1 Then
    'change table reload column
    LoadComboColumn FGridVQB.TextMatrix(1, Col), Col
  End If
  AutoSizeColumnFGrid FGridVQB

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.objFGrid_ComboChange"
End Sub

'remove relation from flexgrid
Private Sub RelQuery_RemoveRelation(Name As String, Tag As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_RemoveRelation(" & QUOTE & Name & QUOTE & "," & QUOTE & Tag & QUOTE & ")", etFullDebug

Dim iCol As Integer
Dim vData

  vData = Split(Tag, ".")
  For iCol = 1 To FGridVQB.Cols - 1
    If FGridVQB.TextMatrix(1, iCol) = Name Then
      RelQuery_MenuActionGridCompose 1, iCol
    End If
  Next
  
  RebildComboTable
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.objFGrid_ComboChange"
End Sub

Private Sub RelQuery_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.RelQuery_MouseUp(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug

  If Button = vbRightButton Then
    PopupMenu mnuRelQuery
  End If
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.RelQuery_MouseUp"
End Sub

Private Sub mnuRelQueryViewSql_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.mnuRelQueryViewSql_Click()", etFullDebug

Dim vData, vData1
Dim ii As Integer
Dim szWhere As String
Dim szWhereFlexGrid As String
Dim szFrom As String
Dim szSelect As String
Dim szOrder As String
Dim colForm As New Collection
Dim iCol As Integer
Dim szSQL As String
    
  'create select/order/where
  szSelect = ""
  szOrder = ""
  szWhereFlexGrid = ""
  For iCol = 1 To FGridVQB.Cols - 1
    If Len(FGridVQB.TextMatrix(1, iCol)) > 0 Then
      'select
      If objFGrid.IsChecked(iCol, 4) Then
        If Len(szSelect) > 0 Then szSelect = szSelect & ", "
        szSelect = szSelect & FGridVQB.TextMatrix(1, iCol) & "." & FGridVQB.TextMatrix(2, iCol)
      End If
      
      'order
      If Len(FGridVQB.TextMatrix(3, iCol)) > 0 Then
        If Len(szOrder) > 0 Then szOrder = szOrder & ", "
        szOrder = szOrder & FGridVQB.TextMatrix(1, iCol) & "." & FGridVQB.TextMatrix(2, iCol) & " " & Left(FGridVQB.TextMatrix(3, iCol), 3)
      End If
      
      'where
      If Len(FGridVQB.TextMatrix(5, iCol)) > 0 Then
        If Len(szWhereFlexGrid) > 0 Then szWhereFlexGrid = szWhereFlexGrid & " AND "
        szWhereFlexGrid = szWhereFlexGrid & "(" & FGridVQB.TextMatrix(5, iCol) & ")"
      End If
    End If
  Next
    
  'create from
  szFrom = ""
  Set colForm = RelQuery.GetRelation
  For Each vData In colForm
    If Len(szFrom) > 0 Then szFrom = szFrom & ", "
    vData1 = Split(vData, ",")
    szFrom = szFrom & " " & vData1(1) & " AS " & vData1(0)
  Next
  
  'create where
  szWhere = ""
  vData = Split(RelQuery.GetJoinRelation, "|")
  For ii = 0 To ((UBound(vData) + 1) / 6) - 1
    If Len(szWhere) > 0 Then szWhere = szWhere & " AND "
    szWhere = szWhere & vData(ii * 3) & "." & vData(ii * 3 + 2)
    szWhere = szWhere & " = "
    szWhere = szWhere & vData(ii * 3 + 3) & "." & vData(ii * 3 + 5)
  Next
    
  'create sql command
  If Len(szSelect) > 0 Then szSQL = "SELECT " & szSelect & vbCrLf
  If Len(szFrom) > 0 Then szSQL = szSQL & "FROM " & szFrom & vbCrLf
  If Len(szWhere) > 0 Then szSQL = szSQL & "WHERE (" & szWhere & ")" & vbCrLf
  If Len(szWhereFlexGrid) > 0 Then
    szSQL = szSQL & IIf(Len(szWhere) > 0, " AND ", " WHERE ")
    szSQL = szSQL & "(" & szWhereFlexGrid & ")" & vbCrLf
  End If
  If Len(szOrder) > 0 Then szSQL = szSQL & "ORDER BY " & szOrder & vbCrLf
  
  txtSQL.Text = szSQL
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.mnuRelQueryViewSql_Click"
End Sub
