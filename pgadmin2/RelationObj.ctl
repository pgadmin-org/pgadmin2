VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RelationObj 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3444
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4476
   ControlContainer=   -1  'True
   ScaleHeight     =   3444
   ScaleWidth      =   4476
   Begin VB.HScrollBar HScroll 
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   3180
      Width           =   4032
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3192
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   252
   End
   Begin VB.Frame fraRelation 
      Caption         =   "Title"
      Height          =   2376
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   1836
   End
   Begin VB.ListBox lstDataRelation 
      Height          =   1968
      Index           =   0
      Left            =   1908
      TabIndex        =   2
      Top             =   420
      Visible         =   0   'False
      Width           =   1392
   End
   Begin MSFlexGridLib.MSFlexGrid FGridCompose 
      Height          =   612
      Left            =   1200
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2712
      _ExtentX        =   4784
      _ExtentY        =   1080
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Image imgDrag 
      Height          =   204
      Left            =   108
      Top             =   2520
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Image ImgJoin2 
      Height          =   132
      Index           =   0
      Left            =   480
      Top             =   2808
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image ImgJoin1 
      Height          =   132
      Index           =   0
      Left            =   108
      Top             =   2808
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Image imgNorm 
      Height          =   60
      Left            =   120
      Picture         =   "RelationObj.ctx":0000
      Top             =   3000
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Menu mnuActionGrid 
      Caption         =   "ActionGrid"
      Visible         =   0   'False
      Begin VB.Menu mnuActionGridAction 
         Caption         =   "Action1"
         Index           =   1
      End
      Begin VB.Menu mnuActionGridAction 
         Caption         =   "Action2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionGridAction 
         Caption         =   "Action3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionGridAction 
         Caption         =   "Action4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuActionGridAction 
         Caption         =   "Action5"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuActionJoin 
      Caption         =   "ActionJoin"
      Begin VB.Menu mnuActionJoinDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuActionRelation 
      Caption         =   "ActionRelation"
      Begin VB.Menu mnuActionRelationRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuActionRelationDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "RelationObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' RelationObj.ctl - Control for create relation

Option Explicit

Dim iCurDragNumber As Integer
Dim szDragItem As String
Dim sOldX As Single
Dim sOldY As Single
Dim iJoinIndexAction As Integer
Dim iRelationIndexAction As Integer

Private Type RelationJoinDetail
  Index As Integer
  ColumnIndex As Integer
  RectRelation As RECT
  RectImage As RECT
  InitailRectImage As RECT
End Type

Private Type RelationJoin
  Enable As Boolean
  Join1 As RelationJoinDetail
  Join2 As RelationJoinDetail
End Type

Private Type RelationData
  Name As String
  Tag As String
  Enable As Boolean
End Type

Dim JoinRelation() As RelationJoin
Dim DataRelation() As RelationData

Event AddElementInGridCompose(Col As Integer, Name As String, Tag As String, Element As String)
Event MenuActionGridCompose(Index As Integer, Col As Integer)
Event RemoveRelation(Name As String, Tag As String)
Event RenameRelation(OldName As String, NewName As String)
Event Click()
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Property Let MenuActionGridEnable(bData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.Property Let MenuActionGridEnable(" & bData & ")", etFullDebug
  
  mnuActionGrid.Enabled = bData
  Exit Property

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.MenuActionGridEnable"
End Property
Public Property Get MenuActionGridEnable() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.Property Get MenuActionGridEnable()", etFullDebug
  
  MenuActionGridEnable = mnuActionGrid.Enabled
  Exit Property

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.MenuActionGridEnable"
End Property

Public Property Get MenuActionGrid(Index As Integer) As Menu
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.Property Get MenuActionGrid(" & Index & ")", etFullDebug
  
  Set MenuActionGrid = mnuActionGridAction(Index)
  Exit Property

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.MenuActionGrid"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_MouseMove(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug
  
  RaiseEvent MouseMove(Button, Shift, x, y)
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_MouseMove"
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_MouseUp(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug
  
  RaiseEvent MouseUp(Button, Shift, x, y)
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_MouseUp"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_MouseDown(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug
  
  RaiseEvent MouseDown(Button, Shift, x, y)
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_MouseDown"
End Sub

Private Sub UserControl_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_Click()", etFullDebug
  
  RaiseEvent Click
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_Click"
End Sub

Private Sub FGridCompose_DragDrop(Source As Control, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.FGridCompose_DragDrop(" & QUOTE & Source.Name & QUOTE & "," & x & "," & y & ")", etFullDebug

Dim iCol As Integer
Dim dColWidth As Double
  
  If szDragItem = "imgDrag" Then
    'find column drag
    dColWidth = 0
    For iCol = 0 To FGridCompose.Cols - 1
      dColWidth = dColWidth + FGridCompose.ColWidth(iCol)
      If dColWidth >= x Then Exit For
    Next
    RaiseEvent AddElementInGridCompose(iCol, DataRelation(iCurDragNumber).Name, DataRelation(iCurDragNumber).Tag, lstDataRelation(iCurDragNumber).Text)
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.FGridCompose_DragDrop"
End Sub

Private Sub UserControl_DragDrop(Source As Control, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_DragDrop(" & QUOTE & Source.Name & QUOTE & "," & x & "," & y & ")", etFullDebug
  
  If szDragItem = "fraRelation" Then
    With fraRelation(iCurDragNumber)
      .Top = y
      
      If .Top + fraRelation(iCurDragNumber).Height > HScroll.Top Then
        .Top = HScroll.Top - fraRelation(iCurDragNumber).Height - 10
      End If
      
      .Left = x - sOldX
      If .Left + fraRelation(iCurDragNumber).Width > VScroll.Left Then
        .Left = VScroll.Left - fraRelation(iCurDragNumber).Width - 10
      End If
    End With
    
    With lstDataRelation(iCurDragNumber)
      .Left = fraRelation(iCurDragNumber).Left
      .Top = fraRelation(iCurDragNumber).Top + fraRelation(iCurDragNumber).Height - .Height
    End With
    DrawJoins
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_DragDrop"
End Sub

Private Sub UserControl_GotFocus()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_Initialize()", etFullDebug

  DrawLines
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_Initialize"
End Sub

Private Sub UserControl_Initialize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
If Not frmMain.svr Is Nothing Then frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_Initialize()", etFullDebug
  
  ReDim JoinRelation(0) As RelationJoin
  ReDim DataRelation(0) As RelationData
  FGridCompose.Visible = False
  
  mnuActionGrid.Enabled = False
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_Initialize"
End Sub

Private Sub UserControl_Resize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
If Not frmMain.svr Is Nothing Then frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.UserControl_Resize()", etFullDebug
  
  If UserControl.Width < 1000 Then UserControl.Width = 1000
  If UserControl.Height < 1000 Then UserControl.Height = 1000

  With VScroll
    .Top = 0
    .Left = UserControl.Width - .Width - 50
    .Height = UserControl.Height - HScroll.Height - 50
    If FGridCompose.Visible Then .Height = .Height - FGridCompose.Height
  End With
  With HScroll
    .Left = 0
    .Width = UserControl.Width - VScroll.Width - 50
    .Top = UserControl.Height - .Height - 50
    If FGridCompose.Visible Then .Top = .Top - FGridCompose.Height
  End With
  
  With FGridCompose
    If .Visible Then
      .Left = 0
      .Width = UserControl.Width - 50
      .Top = HScroll.Top + HScroll.Height
    End If
  End With
  DrawJoins
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.UserControl_Resize"
End Sub

Public Sub AddElement(Name As String, Tag As String, ToolTipText As String, DataRel As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.AddElement(" & QUOTE & Name & QUOTE & "," & QUOTE & Tag & QUOTE & "," & QUOTE & ToolTipText & QUOTE & "," & QUOTE & ")", etFullDebug

Dim iNumRelation As Integer
Dim szName As String
Dim ii As Integer
  
  szName = Name
  'verify valid name relation
  If Not IsValidName(szName) Then
    MsgBox "Relation name not valid!", vbApplicationModal + vbExclamation
    Exit Sub
  End If
  
LoopName:
  'verify if relation exists
  If RelationExists(szName) Then
    szName = InputBox("Change name", "Insert alternate name", szName)
    If Len(szName) = 0 Then Exit Sub
    GoTo LoopName
  End If
  
  'add new relation
  iNumRelation = UBound(DataRelation) + 1
  ReDim Preserve DataRelation(iNumRelation) As RelationData
  With DataRelation(iNumRelation)
    .Name = szName
    .Tag = Tag
    .Enable = True
  End With
  
  Load fraRelation(iNumRelation)
  With fraRelation(iNumRelation)
    .Left = fraRelation(iNumRelation - 1).Left + 1.2 * fraRelation(iNumRelation - 1).Width
    .Caption = DataRelation(iNumRelation).Name
    .ToolTipText = ToolTipText
    .Visible = True
  End With
  
  Load lstDataRelation(iNumRelation)
  With lstDataRelation(iNumRelation)
    .Left = fraRelation(iNumRelation).Left
    .Top = fraRelation(iNumRelation).Top + fraRelation(iNumRelation).Height - .Height
    .Width = fraRelation(iNumRelation).Width
    .ZOrder
    .Visible = True
  End With
  
  'load filed table
  lstDataRelation(iNumRelation).Clear
  For ii = LBound(DataRel) To UBound(DataRel)
    lstDataRelation(iNumRelation).AddItem DataRel(ii)
  Next
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.AddElement"
End Sub

Private Sub fraRelation_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.fraRelation_MouseDown(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug
  
  If Button = vbLeftButton Then
    iCurDragNumber = Index
    szDragItem = fraRelation(Index).Name
    sOldX = x
    sOldY = fraRelation(Index).Top
    fraRelation(Index).Drag vbBeginDrag
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.fraRelation_MouseDown"
End Sub

Private Sub DrawJoins()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
If Not frmMain.svr Is Nothing Then frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.DrawJoins()", etFullDebug

Dim ii As Integer
Dim sDiff As Single
  
  For ii = 1 To UBound(JoinRelation)
    If JoinRelation(ii).Enable Then
      With JoinRelation(ii)
        With .Join1
          sDiff = .RectRelation.Left - lstDataRelation(.Index).Left
          If sDiff <> 0 Then
            .RectImage.Left = .RectImage.Left - sDiff
            .InitailRectImage.Left = .InitailRectImage.Left - sDiff
            .RectRelation.Left = lstDataRelation(.Index).Left
          End If
          sDiff = .RectRelation.Top - lstDataRelation(.Index).Top
          If sDiff <> 0 Then
            .RectImage.Top = .RectImage.Top - sDiff
            .InitailRectImage.Top = .InitailRectImage.Top - sDiff
            .RectRelation.Top = lstDataRelation(.Index).Top
          End If
        End With
        
        With .Join2
          sDiff = .RectRelation.Left - lstDataRelation(.Index).Left
          If sDiff <> 0 Then
            .RectImage.Left = .RectImage.Left - sDiff
            .InitailRectImage.Left = .InitailRectImage.Left - sDiff
            .RectRelation.Left = lstDataRelation(.Index).Left
          End If
          sDiff = .RectRelation.Top - lstDataRelation(.Index).Top
          If sDiff <> 0 Then
            .RectImage.Top = .RectImage.Top - sDiff
            .InitailRectImage.Top = .InitailRectImage.Top - sDiff
            .RectRelation.Top = lstDataRelation(.Index).Top
          End If
        End With
        
        If .Join2.RectImage.Left > .Join1.RectImage.Left Then
          If .Join1.RectImage.Left < fraRelation(.Join1.Index).Left Then
            .Join1.RectImage.Left = fraRelation(.Join1.Index).Left + fraRelation(.Join1.Index).Width
          End If
          .Join2.RectImage.Left = fraRelation(.Join2.Index).Left - 252
        ElseIf .Join1.RectImage.Left > .Join2.RectImage.Left Then
          If .Join2.RectImage.Left < fraRelation(.Join2.Index).Left Then
            .Join2.RectImage.Left = fraRelation(.Join2.Index).Left + fraRelation(.Join2.Index).Width
          End If
          .Join1.RectImage.Left = fraRelation(.Join1.Index).Left - 252
        End If
      End With
      SetPositionImg (ii)
    End If
  Next
  DrawLines
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.DrawJoins"
End Sub

Private Sub SetPositionImg(JoinNumber)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.SetPositionImg(" & JoinNumber & ")", etFullDebug
  
  'set new position image
  If JoinRelation(JoinNumber).Enable Then
    With ImgJoin1(JoinNumber)
      .Left = JoinRelation(JoinNumber).Join1.RectImage.Left
      .Top = JoinRelation(JoinNumber).Join1.RectImage.Top
    End With
    With ImgJoin2(JoinNumber)
      .Left = JoinRelation(JoinNumber).Join2.RectImage.Left
      .Top = JoinRelation(JoinNumber).Join2.RectImage.Top
    End With
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.SetPositionImg"
End Sub

Private Sub DrawLines()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
If Not frmMain.svr Is Nothing Then frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.DrawLines()", etFullDebug

Dim ii As Integer

  UserControl.Refresh
  For ii = 1 To UBound(JoinRelation)
    If JoinRelation(ii).Enable Then
      If ImgJoin2(ii).Left > ImgJoin1(ii).Left Then
        UserControl.Line (ImgJoin2(ii).Left, ImgJoin2(ii).Top + (ImgJoin2(ii).Height / 2))-(ImgJoin1(ii).Left + ImgJoin1(ii).Width, ImgJoin1(ii).Top + (ImgJoin1(ii).Height / 2))
      Else
        UserControl.Line (ImgJoin2(ii).Left + ImgJoin2(ii).Width, ImgJoin2(ii).Top + (ImgJoin2(ii).Height / 2))-(ImgJoin1(ii).Left, ImgJoin1(ii).Top + (ImgJoin1(ii).Height / 2))
      End If
    End If
  Next
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.DrawLines"
End Sub

 
Private Sub lstDataRelation_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.lstDataRelation_DragDrop(" & Index & "," & QUOTE & Source.Name & QUOTE & "," & x & "," & y & ")", etFullDebug

Dim iJoinNum As Integer
Dim bJoin1 As Boolean
Dim bJoin2 As Boolean
  
  If szDragItem = "imgDrag" And Index <> iCurDragNumber Then
    If sOldX > x + lstDataRelation(Index).Left Then
      x = lstDataRelation(Index).Left + lstDataRelation(Index).Width
      sOldX = sOldX - lstDataRelation(iCurDragNumber).Width - 252
    Else
      x = lstDataRelation(Index).Left - 252
    End If
    y = y + lstDataRelation(Index).Top - 100

    'verify if join exists
    For iJoinNum = 0 To UBound(JoinRelation)
      If JoinRelation(iJoinNum).Enable Then
        With JoinRelation(iJoinNum)
          With .Join1
            bJoin1 = ((.Index = Index Or .Index = iCurDragNumber) And _
                      (.ColumnIndex = lstDataRelation(.Index).ListIndex))
          End With
          With .Join2
            bJoin2 = ((.Index = Index Or .Index = iCurDragNumber) And _
                      (.ColumnIndex = lstDataRelation(.Index).ListIndex))
          End With
          If bJoin1 And bJoin2 Then
            MsgBox "Join already exists", vbApplicationModal + vbExclamation
            Exit Sub
          End If
        End With
      End If
    Next
    
    'create new join
    iJoinNum = UBound(JoinRelation) + 1
    ReDim Preserve JoinRelation(iJoinNum) As RelationJoin
    With JoinRelation(iJoinNum)
      .Enable = True
      With .Join1
        .Index = Index
        .ColumnIndex = lstDataRelation(.Index).ListIndex
        .RectRelation.Left = lstDataRelation(.Index).Left
        .RectRelation.Top = lstDataRelation(.Index).Top
        .RectImage.Left = x
        .RectImage.Top = y
        .InitailRectImage = .RectImage
      End With
      
      With .Join2
        .Index = iCurDragNumber
        .ColumnIndex = lstDataRelation(.Index).ListIndex
        .RectRelation.Left = lstDataRelation(.Index).Left
        .RectRelation.Top = lstDataRelation(.Index).Top
        .RectImage.Left = sOldX
        .RectImage.Top = sOldY
        .InitailRectImage = .RectImage
      End With
    End With
    
    'set image join
    Load ImgJoin1(iJoinNum)
    With ImgJoin1(iJoinNum)
      .Picture = imgNorm.Picture
      .Visible = True
      .ZOrder
    End With
    
    Load ImgJoin2(iJoinNum)
    With ImgJoin2(iJoinNum)
      .Picture = imgNorm.Picture
      .Visible = True
      .ZOrder
    End With
  
    SetPositionImg (iJoinNum)
    DrawLines
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.lstDataRelation_DragDrop"
End Sub

Private Sub lstDataRelation_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.lstDataRelation_DragOver(" & Index & "," & QUOTE & Source.Name & QUOTE & "," & x & "," & y & ")", etFullDebug

Dim iPos As Integer
  
  If Index <> iCurDragNumber Then
    iPos = Fix(((y + 240) / 210)) - 1 + lstDataRelation(Index).TopIndex
    If iPos < lstDataRelation(Index).ListCount Then lstDataRelation(Index).ListIndex = iPos
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.lstDataRelation_DragOver"
End Sub

Private Sub lstDataRelation_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.lstDataRelation_MouseMove(" & Index & "," & Button & "," & x & "," & y & ")", etFullDebug
  
  If Button = vbLeftButton Then
    szDragItem = "imgDrag"
    iCurDragNumber = Index
    sOldY = y + lstDataRelation(Index).Top
    sOldX = lstDataRelation(Index).Left + lstDataRelation(Index).Width
    With imgDrag
      .Top = sOldY - .Height / 2
      .Left = lstDataRelation(Index).Left + x - .Width / 2
      .Drag vbBeginDrag
      .Visible = True
    End With
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.lstDataRelation_MouseMove"
End Sub

Private Sub lstDataRelation_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.lstDataRelation_MouseUp(" & Index & "," & Button & "," & x & "," & y & ")", etFullDebug
  
  iCurDragNumber = -1
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.lstDataRelation_MouseUp"
End Sub

Private Sub lstDataRelation_Scroll(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.lstDataRelation_Scroll(" & Index & ")", etFullDebug

Dim ii As Integer

  For ii = 1 To UBound(JoinRelation)
    With JoinRelation(ii)
      If .Enable Then
        If .Join1.Index = Index Then
          If .Join1.InitailRectImage.Top - (210 * lstDataRelation(Index).TopIndex) > lstDataRelation(Index).Top - 240 Then
            .Join1.RectImage.Top = .Join1.InitailRectImage.Top - (210 * (lstDataRelation(Index).TopIndex))
            If .Join1.RectImage.Top < lstDataRelation(Index).Top Then .Join1.RectImage.Top = lstDataRelation(Index).Top
          End If
        ElseIf .Join2.Index = Index Then
          If .Join2.InitailRectImage.Top - (210 * lstDataRelation(Index).TopIndex) > lstDataRelation(Index).Top - 240 Then
            .Join2.RectImage.Top = .Join2.InitailRectImage.Top - (210 * (lstDataRelation(Index).TopIndex))
            If .Join2.RectImage.Top < lstDataRelation(Index).Top Then .Join2.RectImage.Top = lstDataRelation(Index).Top
          End If
        End If
      End If
    End With
  Next
  DrawJoins
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.lstDataRelation_Scroll"
End Sub

Private Sub UserControl_Show()
Dim objCrlt As Control
  
  For Each objCrlt In UserControl.ContainedControls
    If TypeOf objCrlt Is VB.Frame Then
      If objCrlt.Name = "fraContainer" Then
        objCrlt.Left = 0
        objCrlt.ZOrder
        Exit For
      End If
    End If
  Next
End Sub

Private Sub VScroll_Scroll()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.VScroll_Scroll()", etFullDebug

Dim objCtrl As Control
Dim ii As Integer
Static iCurrScrollV As Integer


  For Each objCtrl In Controls
    If objCtrl.Name = "fraRelation" Or objCtrl.Name = "lstDataRelation" Then
      With VScroll
        If .Value > iCurrScrollV Then
          objCtrl.Top = objCtrl.Top - .Value + iCurrScrollV
        Else
          objCtrl.Top = objCtrl.Top + iCurrScrollV - .Value
        End If
      End With
    End If
  Next

  For ii = 1 To UBound(JoinRelation)
    If JoinRelation(ii).Enable Then
      If VScroll.Value > iCurrScrollV Then
        ImgJoin1(ii).Top = ImgJoin1(ii).Top - VScroll.Value + iCurrScrollV
      Else
        ImgJoin1(ii).Top = ImgJoin1(ii).Top + iCurrScrollV - VScroll.Value
      End If
      If VScroll.Value > iCurrScrollV Then
        ImgJoin2(ii).Top = ImgJoin2(ii).Top - VScroll.Value + iCurrScrollV
      Else
        ImgJoin2(ii).Top = ImgJoin2(ii).Top + iCurrScrollV - VScroll.Value
      End If
    End If
  Next
  DrawLines
  iCurrScrollV = VScroll.Value
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.VScroll_Scroll"
End Sub

Private Sub HScroll_Scroll()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.HScroll_Scroll()", etFullDebug

Dim objCtrl As Control
Dim ii As Integer
Static iCurrScrollH As Integer

  For Each objCtrl In Controls
    If objCtrl.Name = "fraRelation" Or objCtrl.Name = "lstDataRelation" Then
      With HScroll
        If .Value > iCurrScrollH Then
          objCtrl.Left = objCtrl.Left - .Value + iCurrScrollH
        Else
          objCtrl.Left = objCtrl.Left + iCurrScrollH - .Value
        End If
      End With
    End If
  Next

  For ii = 1 To UBound(JoinRelation)
    If JoinRelation(ii).Enable Then
      If HScroll.Value > iCurrScrollH Then
        ImgJoin1(ii).Left = ImgJoin1(ii).Left - HScroll.Value + iCurrScrollH
      Else
        ImgJoin1(ii).Left = ImgJoin1(ii).Left + iCurrScrollH - HScroll.Value
      End If
      If HScroll.Value > iCurrScrollH Then
        ImgJoin2(ii).Left = ImgJoin2(ii).Left - HScroll.Value + iCurrScrollH
      Else
        ImgJoin2(ii).Left = ImgJoin2(ii).Left + iCurrScrollH - HScroll.Value
      End If
    End If
  Next
  DrawLines
  iCurrScrollH = HScroll.Value
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.HScroll_Scroll"
End Sub

'return grid compose object
Public Property Get GetGridCompose() As MSFlexGrid
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.Property Get GetGridCompose()", etFullDebug
  
  Set GetGridCompose = FGridCompose
  Exit Property

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.GetGridCompose"
End Property

Public Property Get Controls()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.Property Get Controls()", etFullDebug
  
  Set Controls = UserControl.Controls
  Exit Property

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.Controls"
End Property

Public Function TextWidth(Str As String) As Single
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.TextWidth(" & QUOTE & Str & QUOTE & ")", etFullDebug
  
  TextWidth = UserControl.TextWidth(Str)
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.TextWidth"
End Function

Private Sub FGridCompose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.FGridCompose_MouseUp(" & Button & "," & Shift & "," & x & "," & y & ")", etFullDebug
Dim ii As Integer
Dim iCol As Integer

  If Button = vbRightButton Then
    iCol = 0
    For ii = 0 To FGridCompose.Cols - 1
      If x <= FGridCompose.ColPos(ii) Then
        iCol = ii - 1
        Exit For
      End If
    Next

    If iCol > 0 Then
      FGridCompose.ColSel = iCol
      PopupMenu mnuActionGrid
    End If
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.FGridCompose_MouseUp"
End Sub

Private Sub mnuActionGridAction_Click(Index As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.mnuActionGridAction_Click(" & Index & ")", etFullDebug
  
  If FGridCompose.ColSel > 0 Then RaiseEvent MenuActionGridCompose(Index, FGridCompose.ColSel)
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.mnuActionGridAction_Click"
End Sub

Private Sub ImgJoin1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.ImgJoin1_MouseUp(" & Index & "," & Button & "," & x & "," & y & ")", etFullDebug

  If Button = vbRightButton Then
    iJoinIndexAction = Index
    PopupMenu mnuActionJoin
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.ImgJoin1_MouseUp"
End Sub

Private Sub ImgJoin2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.ImgJoin2_MouseUp(" & Index & "," & Button & "," & x & "," & y & ")", etFullDebug

  If Button = vbRightButton Then
    iJoinIndexAction = Index
    PopupMenu mnuActionJoin
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.ImgJoin2_MouseUp"
End Sub

Private Sub mnuActionJoinDelete_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.mnuActionJoinDelete_Click()", etFullDebug

  If MsgBox("Are you sure you wish to drop the join?", vbSystemModal + vbYesNo + vbQuestion, "Drop join") = vbNo Then Exit Sub
  JoinRelation(iJoinIndexAction).Enable = False
  ImgJoin1(iJoinIndexAction).Visible = False
  ImgJoin2(iJoinIndexAction).Visible = False
  DrawJoins
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.mnuActionJoinDelete_Click"
End Sub

Private Sub fraRelation_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.fraRelation_MouseUp(" & Index & "," & Button & "," & x & "," & y & ")", etFullDebug

  If Button = vbRightButton Then
    iRelationIndexAction = Index
    PopupMenu mnuActionRelation
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.fraRelation_MouseUp"
End Sub

Private Sub mnuActionRelationDelete_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.mnuActionJoinRelationDelete_Click()", etFullDebug

Dim ii As Integer

  If MsgBox("Are you sure you wish to drop the relation?", vbSystemModal + vbYesNo + vbQuestion, "Drop relation") = vbNo Then Exit Sub
  DataRelation(iRelationIndexAction).Enable = False
  fraRelation(iRelationIndexAction).Visible = False
  lstDataRelation(iRelationIndexAction).Visible = False
  
  'disable join relation
  For ii = 1 To UBound(JoinRelation)
    With JoinRelation(ii)
      If .Enable Then
        If .Join1.Index = iRelationIndexAction Or .Join2.Index = iRelationIndexAction Then
          JoinRelation(ii).Enable = False
          ImgJoin1(ii).Visible = False
          ImgJoin2(ii).Visible = False
        End If
      End If
    End With
  Next
  DrawJoins
  RaiseEvent RemoveRelation(DataRelation(iRelationIndexAction).Name, DataRelation(iRelationIndexAction).Tag)
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.mnuActionJoinRelationDelete_Click"
End Sub

'rename relation
Private Sub mnuActionRelationRename_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.mnuActionRelationRename_Click()", etFullDebug

Dim ii As Integer
Dim szTemp As String
Dim szOldName As String
  
  szOldName = DataRelation(iRelationIndexAction).Name
  Do
    szTemp = InputBox("Insert new name for relation " & DataRelation(iRelationIndexAction).Name, "Rename relation", DataRelation(iRelationIndexAction).Name)
    If Len(szTemp) = 0 Then Exit Sub
  
    'verify valid name relation
    If Not IsValidName(szTemp) Then
      MsgBox "New name not valid!", vbApplicationModal + vbExclamation
      Exit Sub
    End If
  
    If szTemp = DataRelation(iRelationIndexAction).Name Then
      MsgBox "Old name and new name is equal!", vbApplicationModal + vbExclamation
      Exit Sub
    End If
  Loop While RelationExists(szTemp)   'verify if relation exists
  
  If MsgBox("Are you sure you wish to rename the relation?", vbSystemModal + vbYesNo + vbQuestion, "Drop relation") = vbNo Then Exit Sub
  DataRelation(iRelationIndexAction).Name = szTemp
  fraRelation(iRelationIndexAction).Caption = szTemp
  
  'disable join relation
  For ii = 1 To UBound(JoinRelation)
    With JoinRelation(ii)
      If .Enable Then
        If .Join1.Index = iRelationIndexAction Or .Join2.Index = iRelationIndexAction Then
          JoinRelation(ii).Enable = False
          ImgJoin1(ii).Visible = False
          ImgJoin2(ii).Visible = False
        End If
      End If
    End With
  Next
  DrawJoins
  
  RaiseEvent RenameRelation(szOldName, DataRelation(iRelationIndexAction).Name)
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.mnuActionRelationRename_Click"
End Sub

'verify if relation exists
Private Function RelationExists(Name As String) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.RelationExists(" & QUOTE & Name & QUOTE & ")", etFullDebug
  
Dim ii As Integer
  
  RelationExists = False
  For ii = 0 To UBound(DataRelation)
    With DataRelation(ii)
      If .Enable And .Name = Name Then
        RelationExists = True
        Exit For
      End If
    End With
  Next
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.RelationExists"
End Function

'verify if name is valid
Private Function IsValidName(Name As String) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.IsValidName(" & QUOTE & Name & QUOTE & ")", etFullDebug

Dim ii As Integer
Dim iChr As Integer
  
  IsValidName = True
  For ii = 1 To Len(Name)
    iChr = Asc(LCase(Mid(Name, ii, 1)))
    If (iChr >= 48 And iChr <= 57) Or _
       (iChr >= 97 And iChr <= 122) Or _
       iChr = 95 Then
    Else
      IsValidName = False
      Exit For
    End If
  Next
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.IsValidName"
End Function

'return the relation object
Public Function GetRelation() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.GetRelation()", etFullDebug

Dim tmpCol As New Collection
Dim ii As Integer

  'verify if relation exists
  For ii = 0 To UBound(DataRelation)
    With DataRelation(ii)
      If .Enable Then tmpCol.Add .Name & "," & .Tag
    End With
  Next
  Set GetRelation = tmpCol
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.GetRelation"
End Function

'return the join relation object
Public Function GetJoinRelation() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":RelationObj.GetJoinRelation()", etFullDebug

Dim ii As Integer
Dim szTemp As String

  szTemp = ""
  For ii = 1 To UBound(JoinRelation)
    With JoinRelation(ii)
      If .Enable Then
        If Len(szTemp) > 0 Then szTemp = szTemp & "|"
        szTemp = szTemp & DataRelation(.Join1.Index).Name & "|" & DataRelation(.Join1.Index).Tag
        szTemp = szTemp & "|" & lstDataRelation(.Join1.Index).List(.Join1.ColumnIndex)
        szTemp = szTemp & "|" & DataRelation(.Join2.Index).Name & "|" & DataRelation(.Join2.Index).Tag
        szTemp = szTemp & "|" & lstDataRelation(.Join2.Index).List(.Join2.ColumnIndex)
      End If
    End With
  Next
  GetJoinRelation = szTemp
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":RelationObj.GetJoinRelation"
End Function