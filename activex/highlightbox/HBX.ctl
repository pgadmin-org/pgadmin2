VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.UserControl HBX 
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   PropertyPages   =   "HBX.ctx":0000
   ScaleHeight     =   1065
   ScaleWidth      =   2865
   ToolboxBitmap   =   "HBX.ctx":0019
   Begin RichTextLib.RichTextBox rtbString 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   180
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"HBX.ctx":0113
   End
   Begin VB.Image imgUp 
      Height          =   150
      Left            =   2655
      Picture         =   "HBX.ctx":0195
      Top             =   30
      Width           =   150
   End
   Begin VB.Image imgDown 
      Height          =   150
      Left            =   2655
      Picture         =   "HBX.ctx":04EF
      Top             =   30
      Visible         =   0   'False
      Width           =   150
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
   Begin VB.Shape shpBar 
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   0
      Top             =   0
      Width           =   2865
   End
End
Attribute VB_Name = "HBX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' HBX - Auto Highlighting Expanding text box
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Default Property Values:
Const m_def_ControlBarVisible = True
Const m_def_MaximisedHeight = 0
Const m_def_MaximisedWidth = 0
Const m_def_Wordlist = ""
Const m_def_AutoColour = True

Const DELIMCHARS = " []{}()'"""

Dim m_ControlBarVisible As Boolean
Dim m_MaximisedWidth As Long
Dim m_MaximisedHeight As Long
Dim m_Wordlist As String
Dim m_AutoColour As Boolean

Dim bTextPropertySet As Boolean
Dim bMaximised As Boolean
Dim LastTop As Long
Dim LastLeft As Long
Dim MinimisedHeight As Long
Dim MinimisedWidth As Long
Dim szRTBFontinfo As String
Dim szRTBWordinfo As String
Dim szRTBColours As String
Dim WordCache As Collection

'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the control is clicked."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the control is double clicked."
Event Change()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when a key is pressed down."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when a key is pressed."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when a key is released."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the mouse button is down."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the mouse is moved."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the mouse button is released."

Public Type WordStyle
  szColour As String
  szRTFstring As String
  szString As String
  bBold As Boolean
  bItalic As Boolean
End Type

Type T_RGB
  R As Integer
  G As Integer
  B As Integer
End Type

Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long

Private Function get_RGB(LColour As Long) As T_RGB
Dim szHEX As String
  szHEX = Hex(LColour)
  While Len(szHEX) < 6
    szHEX = "0" & szHEX
  Wend
  get_RGB.R = CInt("&H" & Mid(szHEX, 5, 2))
  get_RGB.G = CInt("&H" & Mid(szHEX, 3, 2))
  get_RGB.B = CInt("&H" & Mid(szHEX, 1, 2))
End Function

Private Function CharInstr(szChar As String, szString As String) As Boolean
  If InStr(1, szString, szChar) <> 0 Then CharInstr = True
End Function

Private Function SearchCache(szLookup As String) As String
On Error Resume Next
Dim szSearchcache As String

  szSearchcache = WordCache(szLookup).szRTFstring
  If szSearchcache = "" Then
    szSearchcache = szLookup
  Else
    szSearchcache = szSearchcache & " " & szLookup & "\cf0 "
  End If
  SearchCache = szSearchcache
End Function

Private Sub rtbString_Change()
Static lPrevLen As Long
Dim lCount As Long
  If m_AutoColour And Not bTextPropertySet Then
    lCount = Len(rtbString.Text)
    If (lCount > lPrevLen + 2) Or (lCount < lPrevLen - 2) Then QR
    lPrevLen = lCount
  End If
  RaiseEvent Change
End Sub

Private Sub ColourWord() 'Colour the previous word
On Error Resume Next
Dim lWordend As Long
Dim lWordstart As Long
Dim szChar As String
Dim szTemp As String

  lWordend = rtbString.SelStart
  If lWordend < 1 Then Exit Sub

  lWordstart = lWordend
  szChar = Mid(rtbString.Text, lWordstart - 1, 1)

  While (CharInstr(szChar, DELIMCHARS) = False) And szChar <> vbLf
    If lWordstart = 1 Then GoTo fred
    lWordstart = lWordstart - 1
    szChar = Mid(rtbString.Text, lWordstart, 1)
  Wend

fred:
  If lWordstart - 1 = 0 Then
    szTemp = Mid(rtbString.Text, lWordstart, (lWordend - lWordstart))
    rtbString.SelStart = lWordstart - 1
    rtbString.SelLength = (lWordend - lWordstart)
  Else
    rtbString.SelStart = lWordstart + 1
    rtbString.SelLength = (lWordend - lWordstart)
    szTemp = Mid(rtbString.Text, rtbString.SelStart, rtbString.SelLength)
    rtbString.SelStart = lWordstart
    rtbString.SelLength = (lWordend - lWordstart)
  End If

  If WordCache(szTemp).szString & "" <> "" Then
    rtbString.SelBold = WordCache(szTemp).bBold
    rtbString.SelItalic = WordCache(szTemp).bItalic
    rtbString.SelColor = Val(WordCache(szTemp).szColour) 'RGB(0, 0, 255)
  End If

  rtbString.SelStart = lWordend 'Reset cursor position
  rtbString.SelBold = False
  rtbString.SelItalic = False
  rtbString.SelColor = RGB(0, 0, 0)

End Sub
Private Sub QR()
On Error Resume Next
Dim lCurpos As Long
Dim lCount As Long
Dim Stringlist() As String
Dim szOutputstring As String
Dim szTemp As String
Dim szChar As String
Dim lArraypos As Long
Dim szData As String
Dim lX As Long
  ReDim Stringlist(Len(rtbString.Text))
  lCurpos = rtbString.SelStart

  szData = rtbString.Text
  lCount = Len(szData)
  For lX = 1 To lCount
    szChar = Mid(szData, lX, 1)
    If Not CharInstr(szChar, DELIMCHARS) And szChar <> vbLf Then
      szTemp = szTemp & szChar
    Else
      If szChar = vbLf Then szChar = "\par "
      If szChar = "{" Then szChar = "\{"
      If szChar = "}" Then szChar = "\}"
      If szTemp <> "" Then
        Stringlist(lArraypos) = Replace(szTemp, "\", "\\")
        lArraypos = lArraypos + 1
      End If
      Stringlist(lArraypos) = szChar
      lArraypos = lArraypos + 1
      szTemp = ""
    End If
  Next lX

  If szTemp <> "" Then Stringlist(lArraypos) = szTemp
  
  lCount = UBound(Stringlist)
  For lX = 0 To lCount
    If Stringlist(lX) <> "" Then Stringlist(lX) = SearchCache(Stringlist(lX))
  Next lX
  
  For lX = 0 To lCount
    If Stringlist(lX) <> "" Then szOutputstring = szOutputstring & Stringlist(lX)
  Next lX

  rtbString.TextRTF = szRTBFontinfo & szRTBColours & szRTBWordinfo & szOutputstring & "\par}"
  rtbString.SelStart = lCurpos
End Sub

Friend Sub BuildCache()
Dim szStrings() As String
Dim szValues() As String
Dim szColours() As String
Dim szRTBtmp As String
Dim iLoop As Integer
Dim WordDisplay As WordStyle
Dim colRGB As T_RGB
Dim iX As Integer
Dim bColour As Boolean

  rtbString.Text = "|"
  szRTBFontinfo = Mid(rtbString.TextRTF, 1, (InStr(1, rtbString.TextRTF, "|") - 1))
  szRTBWordinfo = Mid(szRTBFontinfo, InStr(1, szRTBFontinfo, "}}") + 2)
  szRTBFontinfo = Mid(szRTBFontinfo, 1, InStr(1, szRTBFontinfo, "}}") + 1)
  ReDim szColours(0)
  
  rtbString.TextRTF = ""
  Set WordCache = New Collection
  szRTBColours = "{\colortbl ;\red0\green0\blue0"
  szStrings = Split(m_Wordlist, ";")
  
  For iLoop = 0 To UBound(szStrings) - 1
    szValues = Split(szStrings(iLoop), "|")
    
    szRTBtmp = szValues(0)   ' Remove the ucase this line from if you upper case require not
    WordDisplay.szString = szValues(0)
    If szValues(1) = "1" Then
      szRTBtmp = "\b " & szRTBtmp & "\b "
      WordDisplay.bBold = True
    Else
      WordDisplay.bBold = False
    End If
    
    If szValues(2) = "1" Then
      szRTBtmp = "\i " & szRTBtmp & "\i "
      WordDisplay.bItalic = True
    Else
      WordDisplay.bItalic = False
    End If
        
    WordDisplay.szColour = szValues(3)
    
    
    'Search for colour
    
    bColour = False
    For iX = 0 To UBound(szColours)
      If szColours(iX) = szValues(3) Then
        bColour = True
        Exit For
      End If
    Next iX
         
    If bColour = False Then
      szColours(UBound(szColours)) = szValues(3)
      WordDisplay.szRTFstring = ("\cf" & UBound(szColours) + 2)
      ReDim Preserve szColours(UBound(szColours) + 1)
    Else
      WordDisplay.szRTFstring = ("\cf" & iX + 2)
    End If
    WordCache.Add WordDisplay, szValues(0)

    szRTBtmp = ""
  Next iLoop
  
  For iX = 0 To UBound(szColours) - 1
    colRGB = get_RGB("" & szColours(iX))
    szRTBColours = szRTBColours & ";" & "\red" & colRGB.R & "\green" & colRGB.G & "\blue" & colRGB.B
  Next iX
          
  szRTBColours = szRTBColours & ";}"
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

Private Sub rtbstring_KeyUp(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyTab Or _
   KeyCode = vbKeyReturn Or _
   KeyCode = vbKeySpace Or _
   KeyCode = vbKeyDelete) And m_AutoColour Then ColourWord

RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Resize()
  If UserControl.Width < 10 Then UserControl.Width = 10
  If UserControl.Height < 500 Then UserControl.Height = 500
  
  shpBar.Width = UserControl.Width
  If UserControl.BorderStyle = 0 Then
    imgDown.Left = UserControl.Width - 200
    imgUp.Left = UserControl.Width - 200
    rtbString.Width = UserControl.Width
    If m_ControlBarVisible = True Then
      rtbString.Top = 0 + shpBar.Height
      rtbString.Height = (UserControl.Height - shpBar.Height)
    Else
      rtbString.Top = 0
      rtbString.Height = UserControl.Height
    End If
  Else
    imgDown.Left = UserControl.Width - 250
    imgUp.Left = UserControl.Width - 250
    rtbString.Width = UserControl.Width - 50
    If m_ControlBarVisible = True Then
      rtbString.Top = 0 + shpBar.Height
      rtbString.Height = (UserControl.Height - shpBar.Height) - 50
    Else
      rtbString.Top = 0
      rtbString.Height = UserControl.Height - 50
    End If
  End If
End Sub

Public Sub Minimise()
Attribute Minimise.VB_Description = "Minimise the control."
  If bMaximised Then imgDown_Click
End Sub

Public Sub Maximise()
Attribute Maximise.VB_Description = "Maximise the control to the size of it's container."
 If Not bMaximised Then imgUp_Click
End Sub

Private Sub rtbstring_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub rtbstring_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub rtbstring_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub rtbstring_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rtbstring_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rtbstring_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub
Public Property Get MaximisedWidth() As Variant
Attribute MaximisedWidth.VB_Description = "Sets/Returns the Maximised Width (0 = width of container)."
  MaximisedWidth = m_MaximisedWidth
End Property

Public Property Let MaximisedWidth(ByVal New_MaximisedWidth As Variant)
  m_MaximisedWidth = New_MaximisedWidth
  PropertyChanged "MaximisedWidth"
End Property

Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font
  
  m_Wordlist = m_def_Wordlist
  m_ControlBarVisible = m_def_ControlBarVisible
  m_MaximisedWidth = m_def_MaximisedWidth
  m_MaximisedHeight = m_def_MaximisedHeight
  m_AutoColour = m_def_AutoColour
  UserControl.BorderStyle = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  LastTop = UserControl.Extender.Top
  LastLeft = UserControl.Extender.Left
  rtbString.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
  rtbString.Enabled = PropBag.ReadProperty("Enabled", True)
  Set rtbString.Font = PropBag.ReadProperty("Font", Ambient.Font)
  rtbString.Locked = PropBag.ReadProperty("Locked", False)
  m_MaximisedHeight = PropBag.ReadProperty("MaximisedHeight", m_def_MaximisedHeight)
  m_MaximisedWidth = PropBag.ReadProperty("MaximisedWidth", m_def_MaximisedWidth)
  lblCaption.Caption = PropBag.ReadProperty("Caption", "")
  rtbString.MaxLength = PropBag.ReadProperty("MaxLength", 0)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  m_ControlBarVisible = PropBag.ReadProperty("ControlBarVisible", m_def_ControlBarVisible)
  m_Wordlist = PropBag.ReadProperty("Wordlist", m_def_Wordlist)
  rtbString.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", True)
  m_AutoColour = PropBag.ReadProperty("AutoColour", m_def_AutoColour)
  rtbString.RightMargin = PropBag.ReadProperty("RightMargin", 0)
  BuildCache
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", rtbString.BackColor, &H80000005)
  Call PropBag.WriteProperty("Enabled", rtbString.Enabled, True)
  Call PropBag.WriteProperty("Font", rtbString.Font, Ambient.Font)
  Call PropBag.WriteProperty("Locked", rtbString.Locked, False)
  Call PropBag.WriteProperty("MaximisedHeight", m_MaximisedHeight, m_def_MaximisedHeight)
  Call PropBag.WriteProperty("MaximisedWidth", m_MaximisedWidth, m_def_MaximisedWidth)
  Call PropBag.WriteProperty("Caption", lblCaption.Caption, "")
  Call PropBag.WriteProperty("MaxLength", rtbString.MaxLength, 0)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
  Call PropBag.WriteProperty("ControlBarVisible", m_ControlBarVisible, m_def_ControlBarVisible)
  Call PropBag.WriteProperty("Wordlist", m_Wordlist, m_def_Wordlist)
  Call PropBag.WriteProperty("AutoVerbMenu", rtbString.AutoVerbMenu, True)
  Call PropBag.WriteProperty("AutoColour", m_AutoColour, m_def_AutoColour)
  Call PropBag.WriteProperty("RightMargin", rtbString.RightMargin, 0)
End Sub

Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
  UserControl.BorderStyle = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Public Property Get AutoColour() As Boolean
  AutoColour = m_AutoColour
End Property

Public Property Let AutoColour(ByVal New_AutoColour As Boolean)
  m_AutoColour = New_AutoColour
  PropertyChanged "AutoColour"
End Property

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns/Sets whether the text can be editted by the user."
  Locked = rtbString.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  rtbString.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Sets/Returns the Maximum text length allowable."
  MaxLength = rtbString.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  rtbString.MaxLength() = New_MaxLength
  PropertyChanged "MaxLength"
End Property

Public Property Get MaximisedHeight() As Variant
Attribute MaximisedHeight.VB_Description = "Sets/Returns the Maximised Height (0 = height of container)."
  MaximisedHeight = m_MaximisedHeight
End Property

Public Property Let MaximisedHeight(ByVal New_MaximisedHeight As Variant)
  m_MaximisedHeight = New_MaximisedHeight
  PropertyChanged "MaximisedHeight"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = rtbString.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  rtbString.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = rtbString.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  rtbString.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
  Set Font = rtbString.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set rtbString.Font = New_Font
  PropertyChanged "Font"
  BuildCache
End Property

Public Property Get Text() As String
Attribute Text.VB_Description = "Sets/Returns the Displayed Text."
  Text = rtbString.Text
End Property

Public Property Let Text(ByVal New_Text As String)
  bTextPropertySet = True
  rtbString.Text() = New_Text
  If m_AutoColour Then QR
  bTextPropertySet = False
  PropertyChanged "Text"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/Sets the Caption."
  Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  lblCaption.Caption() = New_Caption
  PropertyChanged "Caption"
End Property

Public Property Get ControlBarVisible() As Boolean
Attribute ControlBarVisible.VB_Description = "Shows / Hides the Control Bar"
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

Public Sub ColourText()
Attribute ColourText.VB_Description = "Causes a Re-Colouring of the text in the control."
  QR
End Sub

Public Property Get Wordlist() As String
Attribute Wordlist.VB_Description = "Sets/Returns the code string that determines the text highlight."
  Wordlist = m_Wordlist
End Property

Public Property Let Wordlist(ByVal New_Wordlist As String)
  m_Wordlist = New_Wordlist
  PropertyChanged "Wordlist"
  BuildCache
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Sets/Returns the Text Selection start point."
  SelStart = rtbString.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  rtbString.SelStart = New_SelStart
  PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Sets/Returns the Text Selection length."
  SelLength = rtbString.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  rtbString.SelLength = New_SelLength
  PropertyChanged "SelLength"
End Property

Public Property Get Maximised() As Boolean
Attribute Maximised.VB_Description = "Set/Returns a boolean indicating whether the control is Maximised."
  Maximised = bMaximised
End Property

Public Property Get AutoVerbMenu() As Boolean
Attribute AutoVerbMenu.VB_Description = "Returns/sets a value that indicating whether the selected object's verbs will be displayed in a popup menu when the right mouse button is clicked."
  AutoVerbMenu = rtbString.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
  rtbString.AutoVerbMenu() = New_AutoVerbMenu
  PropertyChanged "AutoVerbMenu"
End Property

Public Property Get RightMargin() As Single
Attribute RightMargin.VB_Description = "Sets the right margin used for textwrap, centering, etc."
  RightMargin = rtbString.RightMargin
End Property

Public Property Let RightMargin(ByVal New_RightMargin As Single)
  rtbString.RightMargin() = New_RightMargin
  PropertyChanged "RightMargin"
End Property

