VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVisualQueryBuilder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Query Builder"
   ClientHeight    =   7332
   ClientLeft      =   2028
   ClientTop       =   1812
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7332
   ScaleWidth      =   10080
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid FGridVQB 
      Height          =   1932
      Left            =   2880
      TabIndex        =   5
      Top             =   5340
      Width           =   7152
      _ExtentX        =   12615
      _ExtentY        =   3408
      _Version        =   393216
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
   End
   Begin VB.ListBox lstTable 
      Height          =   1584
      Index           =   0
      Left            =   2100
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   1392
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarH 
      Height          =   252
      Left            =   60
      TabIndex        =   2
      Top             =   5088
      Width           =   9972
      _ExtentX        =   17590
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.Frame fraTable 
      Caption         =   "Title"
      Height          =   2196
      Index           =   0
      Left            =   192
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1836
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1920
      Left            =   48
      TabIndex        =   0
      Top             =   5376
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   3387
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "il"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList il 
      Left            =   144
      Top             =   4464
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":0000
            Key             =   "aggregate"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":06D2
            Key             =   "check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":0DA4
            Key             =   "column"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":1476
            Key             =   "function"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":1B48
            Key             =   "group"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":221A
            Key             =   "index"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":27B4
            Key             =   "indexcolumn"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":2E86
            Key             =   "foreignkey"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":3558
            Key             =   "language"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":3C2A
            Key             =   "operator"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":42FC
            Key             =   "property"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":4896
            Key             =   "relationship"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":49F0
            Key             =   "rule"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":50C2
            Key             =   "server"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":521C
            Key             =   "sequence"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":58EE
            Key             =   "table"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":5FC0
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":6692
            Key             =   "type"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":6D64
            Key             =   "user"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":6EBE
            Key             =   "view"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":7590
            Key             =   "hiproperty"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":7B2A
            Key             =   "database"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":7C84
            Key             =   "closeddatabase"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":7DDE
            Key             =   "baddatabase"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":7F38
            Key             =   "statistics"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":8B0A
            Key             =   "domain"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":91DC
            Key             =   "namespace"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":9DAE
            Key             =   "cast"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVisualQueryBuilder.frx":A980
            Key             =   "conversion"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarV 
      Height          =   5052
      Left            =   9780
      TabIndex        =   4
      Top             =   0
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   8911
      _Version        =   393216
      Orientation     =   1179648
   End
   Begin VB.Image Img1 
      Height          =   252
      Index           =   0
      Left            =   768
      Stretch         =   -1  'True
      Top             =   3744
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image Img2 
      Height          =   252
      Index           =   0
      Left            =   1056
      Stretch         =   -1  'True
      Top             =   3744
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgDrag 
      Height          =   204
      Left            =   768
      Top             =   3456
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Image imgJoinLeft 
      Height          =   252
      Left            =   288
      Picture         =   "frmVisualQueryBuilder.frx":B25A
      Stretch         =   -1  'True
      Top             =   3648
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgNorm 
      Height          =   192
      Left            =   288
      Picture         =   "frmVisualQueryBuilder.frx":B69C
      Top             =   3888
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgJoinRight 
      Height          =   252
      Left            =   288
      Picture         =   "frmVisualQueryBuilder.frx":B7E6
      Stretch         =   -1  'True
      Top             =   3408
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mnuActionDelete 
         Caption         =   "Delete"
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
Dim iNumberTable As Integer
Dim iCurDragNumber As Integer
Dim szDragItem As String
Dim sOldX As Single
Dim sOldY As Single
Dim iCurrScrollV As Integer
Dim iCurrScrollH As Integer
Dim iCurObjAction As Integer

Private Enum VQBTypeJoin
  VQBTJ_NORMAL
  VQBTJ_LEFT
  VQBTJ_RIGTH
End Enum

Private Type VQBJoin1
  TypeJ As VQBTypeJoin
  Index As Integer
  ColumnIndex As Integer
  RectTable As RECT
  RectImage As RECT
End Type

Private Type VQBJoin
  Join1 As VQBJoin1
  Join2 As VQBJoin1
End Type

Private Type VQBTable
  Name As String
  RealName As String
End Type


Dim JoinVQB() As VQBJoin
Dim DataVQB() As VQBTable

Public Sub Initialise(Database As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.Initialise(" & QUOTE & Database & QUOTE & ")", etFullDebug

Dim objNS As pgNamespace
Dim objTable As pgTable
Dim NodeNs As Node
  
  szDB = Database
  
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

  'Initialise Structure
  ReDim DataVQB(0) As VQBTable
  ReDim JoinVQB(0) As VQBJoin

  FGridVQB.FixedRows = 1
  FGridVQB.FixedCols = 0
  FGridVQB.Rows = 5
  FGridVQB.Cols = 64

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.Initialise"
End Sub

'load new table
Private Sub LoadTable(Namespaces As String, Table As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.LoadTable(" & QUOTE & Namespaces & QUOTE & "," & QUOTE & Table & QUOTE & ")", etFullDebug

Dim objColumn As pgColumn

  iNumberTable = iNumberTable + 1
  
  ReDim Preserve DataVQB(iNumberTable) As VQBTable
  With DataVQB(iNumberTable)
    .Name = frmMain.svr.Databases(szDB).Namespaces(Namespaces).Tables(Table).FormattedID
    .RealName = frmMain.svr.Databases(szDB).Namespaces(Namespaces).Tables(Table).Name
  End With
  
  Load fraTable(iNumberTable)
  With fraTable(iNumberTable)
    .Left = 1000 'fraTable(iNumberTable - 1).Left + 1.2 * fraTable(iNumberTable - 1).Width
    .Caption = DataVQB(iNumberTable).Name
    .Visible = True
  End With
  
  Load lstTable(iNumberTable)
  With lstTable(iNumberTable)
    .Left = fraTable(iNumberTable).Left
    .Top = fraTable(iNumberTable).Top + fraTable(iNumberTable).Height - .Height
    .Width = fraTable(iNumberTable).Width
    .ZOrder
    .Visible = True
  End With
  
  'load filed table
  lstTable(iNumberTable).Clear
  lstTable(iNumberTable).AddItem "*"
  For Each objColumn In frmMain.svr.Databases(szDB).Namespaces(Namespaces).Tables(Table).Columns
    If Not (objColumn.SystemObject And Not ctx.IncludeSys) Then
      lstTable(iNumberTable).AddItem objColumn.Name
    End If
  Next
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.LoadTable"
End Sub


Private Sub FGridVQB_DragDrop(Source As Control, X As Single, Y As Single)
Dim iRow As Integer
Dim dColWidth As Double

  If szDragItem = "imgDrag" Then
    dColWidth = 0
    For iRow = 0 To FGridVQB.Cols - 1
      dColWidth = dColWidth + FGridVQB.ColWidth(iRow)
      If dColWidth >= X Then Exit For
    Next
    
    FGridVQB.TextMatrix(1, iRow) = DataVQB(iCurDragNumber).Name
    FGridVQB.TextMatrix(2, iRow) = lstTable(iCurDragNumber).Text
  End If
End Sub

Private Sub Form_Activate()
  DrawLines
End Sub

Private Sub fraTable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    iCurDragNumber = Index
    szDragItem = fraTable(Index).Name
    sOldX = X
    sOldY = fraTable(Index).Top
    fraTable(Index).Drag vbBeginDrag
  ElseIf Button = vbRightButton Then
    iCurObjAction = Index
    PopupMenu mnuAction
  End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
  If szDragItem = "fraTable" Then
    With fraTable(iCurDragNumber)
      .Top = Y
      If .Top + fraTable(iCurDragNumber).Height > FlatScrollBarH.Top Then
        .Top = FlatScrollBarH.Top - fraTable(iCurDragNumber).Height - 10
      End If
      
      .Left = X - sOldX
      If .Left + fraTable(iCurDragNumber).Width > FlatScrollBarV.Left Then
        .Left = FlatScrollBarV.Left - fraTable(iCurDragNumber).Width - 10
      End If
    End With
    
    With lstTable(iCurDragNumber)
      .Left = fraTable(iCurDragNumber).Left
      .Top = fraTable(iCurDragNumber).Top + fraTable(iCurDragNumber).Height - .Height
    End With
    DrawJoins
  End If
End Sub

Private Sub DrawJoins()
Dim ii As Integer
Dim sDiff As Single
  
  For ii = 1 To UBound(JoinVQB)
    With JoinVQB(ii)
      With .Join1
        sDiff = .RectTable.Left - lstTable(.Index).Left
        If sDiff <> 0 Then
          .RectImage.Left = .RectImage.Left - sDiff
          .RectTable.Left = lstTable(.Index).Left
        End If
        sDiff = .RectTable.Top - lstTable(.Index).Top
        If sDiff <> 0 Then
          .RectImage.Top = .RectImage.Top - sDiff
          .RectTable.Top = lstTable(.Index).Top
        End If
      End With
      
      With .Join2
        sDiff = .RectTable.Left - lstTable(.Index).Left
        If sDiff <> 0 Then
          .RectImage.Left = .RectImage.Left - sDiff
          .RectTable.Left = lstTable(.Index).Left
        End If
        sDiff = .RectTable.Top - lstTable(.Index).Top
        If sDiff <> 0 Then
          .RectImage.Top = .RectImage.Top - sDiff
          .RectTable.Top = lstTable(.Index).Top
        End If
      End With
      
      If .Join2.RectImage.Left > .Join1.RectImage.Left Then
        If .Join1.RectImage.Left < fraTable(.Join1.Index).Left Then
          .Join1.RectImage.Left = fraTable(.Join1.Index).Left + fraTable(.Join1.Index).Width
        End If
        .Join2.RectImage.Left = fraTable(.Join2.Index).Left - 252
      ElseIf .Join1.RectImage.Left > .Join2.RectImage.Left Then
        If .Join2.RectImage.Left < fraTable(.Join2.Index).Left Then
          .Join2.RectImage.Left = fraTable(.Join2.Index).Left + fraTable(.Join2.Index).Width
        End If
        .Join1.RectImage.Left = fraTable(.Join1.Index).Left - 252
      End If
    End With
    SetPositionImg (ii)
  Next
  DrawLines
End Sub

Private Sub SetPositionImg(JoinNumber)
  'set new position image
  With Img1(JoinNumber)
    .Left = JoinVQB(JoinNumber).Join1.RectImage.Left
    .Top = JoinVQB(JoinNumber).Join1.RectImage.Top
  End With
  With Img2(JoinNumber)
    .Left = JoinVQB(JoinNumber).Join2.RectImage.Left
    .Top = JoinVQB(JoinNumber).Join2.RectImage.Top
  End With
End Sub


Private Sub lstTable_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
Dim ii As Integer
    
  If Index <> iCurDragNumber Then
    ii = Fix(((Y + 240) / 210)) - 1 + lstTable(Index).TopIndex
    lstTable(Index).ListIndex = ii
  End If
End Sub

Private Sub lstTable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    szDragItem = "imgDrag"
    iCurDragNumber = Index
    sOldY = Y + lstTable(Index).Top
    sOldX = lstTable(Index).Left + lstTable(Index).Width
    With imgDrag
      .Top = sOldY - .Height / 2
      .Left = fraTable(Index).Left + X - .Width / 2
      .Drag vbBeginDrag
      .Visible = True
    End With
  End If
End Sub

Private Sub lstTable_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Dim iJoinNum As Integer
Dim bJoin1 As Boolean
Dim bJoin2 As Boolean
  
  If szDragItem = "imgDrag" And Index <> iCurDragNumber Then
    If sOldX > X + lstTable(Index).Left Then
      X = lstTable(Index).Left + lstTable(Index).Width
      sOldX = sOldX - lstTable(iCurDragNumber).Width - 252
    Else
      X = lstTable(Index).Left - 252
    End If
    Y = Y + lstTable(Index).Top - 100

    'verify if join exists
    For iJoinNum = 0 To UBound(JoinVQB)
      With JoinVQB(iJoinNum)
        With .Join1
          bJoin1 = ((.Index = Index Or .Index = iCurDragNumber) And _
                    (.ColumnIndex = lstTable(.Index).ListIndex))
        End With
        With .Join2
          bJoin2 = ((.Index = Index Or .Index = iCurDragNumber) And _
                    (.ColumnIndex = lstTable(.Index).ListIndex))
        End With
        If bJoin1 And bJoin2 Then
          MsgBox "Join already exists", vbApplicationModal + vbExclamation
          Exit Sub
        End If
      End With
    Next
    
    'create new join
    iJoinNum = UBound(JoinVQB) + 1
    ReDim Preserve JoinVQB(iJoinNum) As VQBJoin
    With JoinVQB(iJoinNum)
      With .Join1
        .TypeJ = VQBTJ_NORMAL
        .Index = Index
        .ColumnIndex = lstTable(.Index).ListIndex
        
        .RectTable.Left = lstTable(.Index).Left
        .RectTable.Top = lstTable(.Index).Top
        
        With .RectImage
          .Left = X
          .Top = Y
        End With
      End With
      
      With .Join2
        .TypeJ = VQBTJ_NORMAL
        .Index = iCurDragNumber
        .ColumnIndex = lstTable(.Index).ListIndex
        
        .RectTable.Left = lstTable(.Index).Left
        .RectTable.Top = lstTable(.Index).Top
        
        With .RectImage
          .Left = sOldX
          .Top = sOldY
        End With
      End With
    End With
    
    'set image join
    Load Img1(iJoinNum)
    With Img1(iJoinNum)
      .Picture = imgNorm.Picture
      .Visible = True
      .ZOrder
    End With
    
    Load Img2(iJoinNum)
    With Img2(iJoinNum)
      .Picture = imgNorm.Picture
      .Visible = True
      .ZOrder
    End With
  
    SetPositionImg (iJoinNum)
    DrawLines
  End If
End Sub

Private Sub DrawLines()
Dim ii As Integer

  Me.Refresh
  For ii = 1 To UBound(JoinVQB)
    If Img2(ii).Left > Img1(ii).Left Then
      Me.Line (Img2(ii).Left, Img2(ii).Top + (Img2(ii).Height / 2))-(Img1(ii).Left + Img1(ii).Width, Img1(ii).Top + (Img1(ii).Height / 2))
    Else
      Me.Line (Img2(ii).Left + Img2(ii).Width, Img2(ii).Top + (Img2(ii).Height / 2))-(Img1(ii).Left, Img1(ii).Top + (Img1(ii).Height / 2))
    End If
  Next
End Sub

Private Sub lstTable_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  iCurDragNumber = -1
End Sub

Private Sub mnuActionDelete_Click()
Dim ii As Integer
Dim iPos As Integer
Dim tmpJoin() As VQBJoin
  
Exit Sub
  Unload fraTable(iCurObjAction)
  Unload lstTable(iCurObjAction)
  If UBound(JoinVQB) > 0 Then
    ReDim tmpJoin(0) As VQBJoin
    iPos = ii
    For ii = 1 To UBound(JoinVQB)
      If JoinVQB(ii).Join1.Index = iCurObjAction Or JoinVQB(ii).Join2.Index = iCurObjAction Then
        Unload Img1(ii)
        Unload Img2(ii)
      Else
        ReDim Preserve tmpJoin(iPos) As VQBJoin
        tmpJoin(iPos) = JoinVQB(ii)
        iPos = iPos + 1
      End If
    Next
    JoinVQB = tmpJoin
    DrawJoins
  End If
End Sub

Private Sub tv_DblClick()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.tv_DblClick()", etFullDebug
  
  If tv.SelectedItem Is Nothing Then Exit Sub
  If Left(tv.SelectedItem.Key, 3) = "TBL" Then
    LoadTable tv.SelectedItem.Parent, tv.SelectedItem.Text
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.tv_DblClick"
End Sub

Private Sub FlatScrollBarH_Scroll()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.FlatScrollBarH_Scroll()", etFullDebug

Dim objCtrl As Control
Dim ii As Integer

  For Each objCtrl In Controls
    If objCtrl.Name = "fraTable" Or objCtrl.Name = "lstTable" Then
      With FlatScrollBarH
        If .Value > iCurrScrollH Then
          objCtrl.Left = objCtrl.Left - .Value + iCurrScrollH
        Else
          objCtrl.Left = objCtrl.Left + iCurrScrollH - .Value
        End If
      End With
    End If
  Next

  For ii = 1 To UBound(JoinVQB)
    If FlatScrollBarH.Value > iCurrScrollH Then
      Img1(ii).Left = Img1(ii).Left - FlatScrollBarH.Value + iCurrScrollH
    Else
      Img1(ii).Left = Img1(ii).Left + iCurrScrollH - FlatScrollBarH.Value
    End If
    If FlatScrollBarH.Value > iCurrScrollH Then
      Img2(ii).Left = Img2(ii).Left - FlatScrollBarH.Value + iCurrScrollH
    Else
      Img2(ii).Left = Img2(ii).Left + iCurrScrollH - FlatScrollBarH.Value
    End If
  Next
  DrawLines
  iCurrScrollH = FlatScrollBarH.Value
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.FlatScrollBarH_Scroll"
End Sub
Private Sub FlatScrollBarV_Scroll()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmVisualQueryBuilder.FlatScrollBarV_Scroll()", etFullDebug

Dim objCtrl As Control
Dim ii As Integer

  For Each objCtrl In Controls
    If objCtrl.Name = "fraTable" Or objCtrl.Name = "lstTable" Then
      With FlatScrollBarV
        If .Value > iCurrScrollV Then
          objCtrl.Top = objCtrl.Top - .Value + iCurrScrollV
        Else
          objCtrl.Top = objCtrl.Top + iCurrScrollV - .Value
        End If
      End With
    End If
  Next

  For ii = 1 To UBound(JoinVQB)
    If FlatScrollBarV.Value > iCurrScrollV Then
      Img1(ii).Top = Img1(ii).Top - FlatScrollBarV.Value + iCurrScrollV
    Else
      Img1(ii).Top = Img1(ii).Top + iCurrScrollV - FlatScrollBarV.Value
    End If
    If FlatScrollBarV.Value > iCurrScrollV Then
      Img2(ii).Top = Img2(ii).Top - FlatScrollBarV.Value + iCurrScrollV
    Else
      Img2(ii).Top = Img2(ii).Top + iCurrScrollV - FlatScrollBarV.Value
    End If
  Next
  DrawLines
  iCurrScrollV = FlatScrollBarV.Value
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmVisualQueryBuilder.FlatScrollBarV_Scroll"
End Sub

Private Sub lstTable_Scroll(Index As Integer)
Dim ii As Integer

  For ii = 1 To UBound(JoinVQB)
    With JoinVQB(ii)
      If .Join1.Index = Index Then
        DrawLines
        Exit For
      ElseIf .Join2.Index = Index Then
        DrawLines
        Exit For
      End If
    End With
  Next
  DrawLines
End Sub

