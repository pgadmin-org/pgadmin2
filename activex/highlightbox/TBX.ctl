VERSION 5.00
Begin VB.UserControl TBX 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   945
   ScaleWidth      =   2865
   ToolboxBitmap   =   "TBX.ctx":0000
   Begin VB.TextBox txtString 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   180
      Width           =   2850
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000009&
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2475
   End
   Begin VB.Image imgUp 
      Height          =   150
      Left            =   2655
      Picture         =   "TBX.ctx":00FA
      Top             =   30
      Width           =   150
   End
   Begin VB.Image imgDown 
      Height          =   150
      Left            =   2655
      Picture         =   "TBX.ctx":0454
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Shape shpBar 
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   0
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "TBX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' HBX - Auto Highlighting Expanding text box
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Default Property Values:
Const m_def_ControlBarVisible = True
Const m_def_MaximisedHeight = 0
Const m_def_MaximisedWidth = 0

Dim m_ControlBarVisible As Boolean
Dim m_MaximisedWidth As Long
Dim m_MaximisedHeight As Long

Dim bMaximised As Boolean
Dim LastTop As Long
Dim LastLeft As Long
Dim MinimisedHeight As Long
Dim MinimisedWidth As Long

'Event Declarations:
Event Click()
Event DblClick()
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private Sub txtString_Change()
  RaiseEvent Change
End Sub

Private Sub imgUp_Click()
On Error Resume Next
Dim lHeight As Long
Dim lWidth As Long

  MinimisedHeight = UserControl.Height
  MinimisedWidth = UserControl.Width
  LastTop = UserControl.Extender.Top
  LastLeft = UserControl.Extender.Left
  
  imgUp.Visible = False
  imgDown.Visible = True

  If m_MaximisedHeight = 0 Then
    lHeight = UserControl.Extender.Container.ScaleHeight
    If lHeight = 0 Then lHeight = UserControl.Extender.Container.Height
    UserControl.Height = lHeight
    UserControl.Extender.Top = 0
  Else
    If m_MaximisedHeight > MinimisedHeight Then UserControl.Height = m_MaximisedHeight
  End If
  
  If m_MaximisedWidth = 0 Then
    lWidth = UserControl.Extender.Container.ScaleWidth
    If lWidth = 0 Then lWidth = UserControl.Extender.Container.Width
    UserControl.Width = lWidth
    UserControl.Extender.Left = 0
  Else
    If m_MaximisedWidth > MinimisedWidth Then UserControl.Width = m_MaximisedWidth
  End If

  UserControl.Extender.ZOrder 0
  bMaximised = True
End Sub

Private Sub imgDown_Click()
  imgDown.Visible = False
  imgUp.Visible = True
  UserControl.Height = MinimisedHeight
  UserControl.Width = MinimisedWidth
  UserControl.Extender.Top = LastTop
  UserControl.Extender.Left = LastLeft
  bMaximised = False
End Sub

Private Sub txtString_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
  If UserControl.Width < 10 Then UserControl.Width = 10
  If UserControl.Height < 500 Then UserControl.Height = 500
  
  shpBar.Width = UserControl.Width
  If UserControl.BorderStyle = 0 Then
    imgDown.Left = UserControl.Width - 200
    imgUp.Left = UserControl.Width - 200
    txtString.Width = UserControl.Width - 60
    If m_ControlBarVisible = True Then
      txtString.Top = 0 + shpBar.Height
      txtString.Height = (UserControl.Height - shpBar.Height)
    Else
      txtString.Top = 0
      txtString.Height = UserControl.Height
    End If
  Else
    imgDown.Left = UserControl.Width - 250
    imgUp.Left = UserControl.Width - 250
    txtString.Width = UserControl.Width - 50
    If m_ControlBarVisible = True Then
      txtString.Top = 0 + shpBar.Height
      txtString.Height = (UserControl.Height - shpBar.Height) - 60
    Else
      txtString.Top = 0
      txtString.Height = UserControl.Height - 60
    End If
  End If
End Sub

Public Sub Minimise()
  If bMaximised Then imgDown_Click
End Sub

Public Sub Maximise()
 If Not bMaximised Then imgUp_Click
End Sub

Private Sub txtString_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub txtString_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtString_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub txtString_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub txtString_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Public Property Get MaximisedWidth() As Variant
  MaximisedWidth = m_MaximisedWidth
End Property

Public Property Let MaximisedWidth(ByVal New_MaximisedWidth As Variant)
  m_MaximisedWidth = New_MaximisedWidth
  PropertyChanged "MaximisedWidth"
End Property

Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font

  m_ControlBarVisible = m_def_ControlBarVisible
  m_MaximisedWidth = m_def_MaximisedWidth
  m_MaximisedHeight = m_def_MaximisedHeight
  UserControl.BorderStyle = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  LastTop = UserControl.Extender.Top
  LastLeft = UserControl.Extender.Left
  txtString.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  txtString.Enabled = PropBag.ReadProperty("Enabled", True)
  Set txtString.Font = PropBag.ReadProperty("Font", Ambient.Font)
  txtString.Locked = PropBag.ReadProperty("Locked", False)
  m_MaximisedHeight = PropBag.ReadProperty("MaximisedHeight", m_def_MaximisedHeight)
  m_MaximisedWidth = PropBag.ReadProperty("MaximisedWidth", m_def_MaximisedWidth)
  lblCaption.Caption = PropBag.ReadProperty("Caption", "")
  txtString.MaxLength = PropBag.ReadProperty("MaxLength", 0)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  m_ControlBarVisible = PropBag.ReadProperty("ControlBarVisible", m_def_ControlBarVisible)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", txtString.BackColor, &H80000005)
  Call PropBag.WriteProperty("Enabled", txtString.Enabled, True)
  Call PropBag.WriteProperty("Font", txtString.Font, Ambient.Font)
  Call PropBag.WriteProperty("Locked", txtString.Locked, False)
  Call PropBag.WriteProperty("MaximisedHeight", m_MaximisedHeight, m_def_MaximisedHeight)
  Call PropBag.WriteProperty("MaximisedWidth", m_MaximisedWidth, m_def_MaximisedWidth)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("MaxLength", txtString.MaxLength, 0)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
  Call PropBag.WriteProperty("ControlBarVisible", m_ControlBarVisible, m_def_ControlBarVisible)
End Sub

Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
  UserControl.BorderStyle = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Public Property Get Locked() As Boolean
  Locked = txtString.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  txtString.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
  MaxLength = txtString.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  txtString.MaxLength() = New_MaxLength
  PropertyChanged "MaxLength"
End Property

Public Property Get MaximisedHeight() As Variant
  MaximisedHeight = m_MaximisedHeight
End Property

Public Property Let MaximisedHeight(ByVal New_MaximisedHeight As Variant)
  m_MaximisedHeight = New_MaximisedHeight
  PropertyChanged "MaximisedHeight"
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = txtString.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  txtString.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
  Enabled = txtString.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  txtString.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
  Set Font = txtString.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set txtString.Font = New_Font
  PropertyChanged "Font"
End Property

Public Property Get Text() As String
  Text = txtString.Text
End Property

Public Property Let Text(ByVal New_Text As String)
  txtString.Text() = New_Text
  PropertyChanged "Text"
End Property

Public Property Get Caption() As String
  Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  lblCaption.Caption() = New_Caption
  PropertyChanged "Caption"
End Property

Public Property Get ControlBarVisible() As Boolean
  ControlBarVisible = m_ControlBarVisible
End Property

Public Property Let ControlBarVisible(ByVal New_ControlBarVisible As Boolean)
  m_ControlBarVisible = New_ControlBarVisible
  lblCaption.Visible = m_ControlBarVisible
  imgUp.Visible = m_ControlBarVisible
  imgDown.Visible = False
  shpBar.Visible = m_ControlBarVisible
  UserControl_Resize
  PropertyChanged "ControlBarVisible"
End Property

Public Property Get SelStart() As Long
  SelStart = txtString.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  txtString.SelStart = New_SelStart
  PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Long
  SelLength = txtString.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  txtString.SelLength = New_SelLength
  PropertyChanged "SelLength"
End Property

Public Property Get Maximised() As Boolean
  Maximised = bMaximised
End Property




