VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTrasLang 
   Caption         =   "Traslation Language"
   ClientHeight    =   5220
   ClientLeft      =   2232
   ClientTop       =   2232
   ClientWidth     =   8700
   Icon            =   "frmTrasLang.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8700
   Begin VB.TextBox txtOriginal 
      BackColor       =   &H8000000F&
      Height          =   912
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   8592
   End
   Begin VB.ComboBox cboNewLang 
      Height          =   288
      Left            =   4140
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   360
      Width           =   2592
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   312
      Left            =   2700
      TabIndex        =   7
      Top             =   360
      Width           =   1392
   End
   Begin VB.CommandButton cmdAddLang 
      Caption         =   "&Add Language"
      Height          =   312
      Left            =   6780
      TabIndex        =   6
      Top             =   360
      Width           =   1392
   End
   Begin VB.ComboBox cboLang 
      Height          =   288
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2592
   End
   Begin MSComctlLib.ProgressBar PBarCompleted 
      Height          =   252
      Left            =   1440
      TabIndex        =   2
      Top             =   4920
      Width           =   2712
      _ExtentX        =   4784
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.TextBox txtTraslate 
      Height          =   912
      Left            =   60
      TabIndex        =   1
      Top             =   3960
      Width           =   8592
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2232
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   8592
      _ExtentX        =   15155
      _ExtentY        =   3937
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Original"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Traslate"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   7980
      Top             =   60
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrasLang.frx":06C2
            Key             =   "property"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "New Language"
      Height          =   252
      Left            =   4140
      TabIndex        =   10
      Top             =   120
      Width           =   2532
   End
   Begin VB.Label lblCompleted 
      Height          =   252
      Left            =   60
      TabIndex        =   5
      Top             =   4920
      Width           =   1272
   End
   Begin VB.Label Label1 
      Caption         =   "Language"
      Height          =   252
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   2532
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsExtractString 
         Caption         =   "Extract string from source"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsMerge 
         Caption         =   "Merge"
         Begin VB.Menu mnuToolsMergeFile 
            Caption         =   "Merge File..."
         End
         Begin VB.Menu mnuToolsMergeAllFile 
            Caption         =   "Merge All File"
         End
      End
   End
End
Attribute VB_Name = "frmTrasLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' frmTrasLang.frm - Traslation Lang

Option Explicit
Dim objDataLang() As StrLang

Public Sub Initialise()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.Initialise()", etFullDebug

Dim szTemp As String
Dim ii As Integer

  cmdSave.Enabled = False
  
  'add all language
  cboNewLang.Clear
  cboNewLang.AddItem "(Afan) Oromo"
  cboNewLang.AddItem "Abkhazian"
  cboNewLang.AddItem "Afar"
  cboNewLang.AddItem "Afrikaans"
  cboNewLang.AddItem "Albanian"
  cboNewLang.AddItem "Amharic"
  cboNewLang.AddItem "Arabic"
  cboNewLang.AddItem "Arabic (Algeria)"
  cboNewLang.AddItem "Arabic (Bahrain)"
  cboNewLang.AddItem "Arabic (Egypt)"
  cboNewLang.AddItem "Arabic (Iraq)"
  cboNewLang.AddItem "Arabic (Jordan)"
  cboNewLang.AddItem "Arabic (Kuwait)"
  cboNewLang.AddItem "Arabic (Lebanon)"
  cboNewLang.AddItem "Arabic (Libya)"
  cboNewLang.AddItem "Arabic (Morocco)"
  cboNewLang.AddItem "Arabic (Oman)"
  cboNewLang.AddItem "Arabic (Qatar)"
  cboNewLang.AddItem "Arabic (Saudi Arabia)"
  cboNewLang.AddItem "Arabic (Sudan)"
  cboNewLang.AddItem "Arabic (Syria)"
  cboNewLang.AddItem "Arabic (Tunisia)"
  cboNewLang.AddItem "Arabic (Uae)"
  cboNewLang.AddItem "Arabic (Yemen)"
  cboNewLang.AddItem "Armenian"
  cboNewLang.AddItem "Assamese"
  cboNewLang.AddItem "Aymara"
  cboNewLang.AddItem "Azeri"
  cboNewLang.AddItem "Azeri (Cyrillic)"
  cboNewLang.AddItem "Azeri (Latin)"
  cboNewLang.AddItem "Bashkir"
  cboNewLang.AddItem "Basque"
  cboNewLang.AddItem "Belarusian"
  cboNewLang.AddItem "Bengali"
  cboNewLang.AddItem "Bhutani"
  cboNewLang.AddItem "Bihari"
  cboNewLang.AddItem "Bislama"
  cboNewLang.AddItem "Breton"
  cboNewLang.AddItem "Bulgarian"
  cboNewLang.AddItem "Burmese"
  cboNewLang.AddItem "Cambodian"
  cboNewLang.AddItem "Catalan"
  cboNewLang.AddItem "Chinese"
  cboNewLang.AddItem "Chinese (Hongkong)"
  cboNewLang.AddItem "Chinese (Macau)"
  cboNewLang.AddItem "Chinese (Simplified)"
  cboNewLang.AddItem "Chinese (Singapore)"
  cboNewLang.AddItem "Chinese (Taiwan)"
  cboNewLang.AddItem "Chinese (Traditional)"
  cboNewLang.AddItem "Corsican"
  cboNewLang.AddItem "Croatian"
  cboNewLang.AddItem "Czech"
  cboNewLang.AddItem "Danish"
  cboNewLang.AddItem "Dutch"
  cboNewLang.AddItem "Dutch (Belgian)"
  cboNewLang.AddItem "English"
  cboNewLang.AddItem "English (Australia)"
  cboNewLang.AddItem "English (Belize)"
  cboNewLang.AddItem "English (Botswana)"
  cboNewLang.AddItem "English (Canada)"
  cboNewLang.AddItem "English (Caribbean)"
  cboNewLang.AddItem "English (Denmark)"
  cboNewLang.AddItem "English (Eire)"
  cboNewLang.AddItem "English (Jamaica)"
  cboNewLang.AddItem "English (New Zealand)"
  cboNewLang.AddItem "English (Philippines)"
  cboNewLang.AddItem "English (South Africa)"
  cboNewLang.AddItem "English (Trinidad)"
  cboNewLang.AddItem "English (U.K.)"
  cboNewLang.AddItem "English (U.S.)"
  cboNewLang.AddItem "English (Zimbabwe)"
  cboNewLang.AddItem "Esperanto"
  cboNewLang.AddItem "Estonian"
  cboNewLang.AddItem "Faeroese"
  cboNewLang.AddItem "Farsi"
  cboNewLang.AddItem "Fiji"
  cboNewLang.AddItem "Finnish"
  cboNewLang.AddItem "French"
  cboNewLang.AddItem "French (Belgian)"
  cboNewLang.AddItem "French (Canadian)"
  cboNewLang.AddItem "French (Luxembourg)"
  cboNewLang.AddItem "French (Monaco)"
  cboNewLang.AddItem "French (Swiss)"
  cboNewLang.AddItem "Frisian"
  cboNewLang.AddItem "Galician"
  cboNewLang.AddItem "Georgian"
  cboNewLang.AddItem "German"
  cboNewLang.AddItem "German (Austrian)"
  cboNewLang.AddItem "German (Belgium)"
  cboNewLang.AddItem "German (Liechtenstein)"
  cboNewLang.AddItem "German (Luxembourg)"
  cboNewLang.AddItem "German (Swiss)"
  cboNewLang.AddItem "Greek"
  cboNewLang.AddItem "Greenlandic"
  cboNewLang.AddItem "Guarani"
  cboNewLang.AddItem "Gujarati"
  cboNewLang.AddItem "Hausa"
  cboNewLang.AddItem "Hebrew"
  cboNewLang.AddItem "Hindi"
  cboNewLang.AddItem "Hungarian"
  cboNewLang.AddItem "Icelandic"
  cboNewLang.AddItem "Indonesian"
  cboNewLang.AddItem "Interlingua"
  cboNewLang.AddItem "Interlingue"
  cboNewLang.AddItem "Inuktitut"
  cboNewLang.AddItem "Inupiak"
  cboNewLang.AddItem "Irish"
  cboNewLang.AddItem "Italian"
  cboNewLang.AddItem "Italian (Swiss)"
  cboNewLang.AddItem "Japanese"
  cboNewLang.AddItem "Javanese"
  cboNewLang.AddItem "Kannada"
  cboNewLang.AddItem "Kashmiri"
  cboNewLang.AddItem "Kashmiri (India)"
  cboNewLang.AddItem "Kazakh"
  cboNewLang.AddItem "Kernewek"
  cboNewLang.AddItem "Kinyarwanda"
  cboNewLang.AddItem "Kirghiz"
  cboNewLang.AddItem "Kirundi"
  cboNewLang.AddItem "Konkani"
  cboNewLang.AddItem "Korean"
  cboNewLang.AddItem "Kurdish"
  cboNewLang.AddItem "Laothian"
  cboNewLang.AddItem "Latin"
  cboNewLang.AddItem "Latvian"
  cboNewLang.AddItem "Lingala"
  cboNewLang.AddItem "Lithuanian"
  cboNewLang.AddItem "Macedonian"
  cboNewLang.AddItem "Malagasy"
  cboNewLang.AddItem "Malay"
  cboNewLang.AddItem "Malay (Brunei Darussalam)"
  cboNewLang.AddItem "Malay (Malaysia)"
  cboNewLang.AddItem "Malayalam"
  cboNewLang.AddItem "Maltese"
  cboNewLang.AddItem "Manipuri"
  cboNewLang.AddItem "Maori"
  cboNewLang.AddItem "Marathi"
  cboNewLang.AddItem "Moldavian"
  cboNewLang.AddItem "Mongolian"
  cboNewLang.AddItem "Nauru"
  cboNewLang.AddItem "Nepali"
  cboNewLang.AddItem "Nepali (India)"
  cboNewLang.AddItem "Norwegian (Bokmal)"
  cboNewLang.AddItem "Norwegian (Nynorsk)"
  cboNewLang.AddItem "Occitan"
  cboNewLang.AddItem "Oriya"
  cboNewLang.AddItem "Pashto, Pushto"
  cboNewLang.AddItem "Polish"
  cboNewLang.AddItem "Portuguese"
  cboNewLang.AddItem "Portuguese (Brazilian)"
  cboNewLang.AddItem "Punjabi"
  cboNewLang.AddItem "Quechua"
  cboNewLang.AddItem "Rhaeto-Romance"
  cboNewLang.AddItem "Romanian"
  cboNewLang.AddItem "Russian"
  cboNewLang.AddItem "Russian (Ukraine)"
  cboNewLang.AddItem "Samoan"
  cboNewLang.AddItem "Sangho"
  cboNewLang.AddItem "Sanskrit"
  cboNewLang.AddItem "Scots Gaelic"
  cboNewLang.AddItem "Serbian"
  cboNewLang.AddItem "Serbian (Cyrillic)"
  cboNewLang.AddItem "Serbian (Latin)"
  cboNewLang.AddItem "Serbo-Croatian"
  cboNewLang.AddItem "Sesotho"
  cboNewLang.AddItem "Setswana"
  cboNewLang.AddItem "Shona"
  cboNewLang.AddItem "Sindhi"
  cboNewLang.AddItem "Sinhalese"
  cboNewLang.AddItem "Siswati"
  cboNewLang.AddItem "Slovak"
  cboNewLang.AddItem "Slovenian"
  cboNewLang.AddItem "Somali"
  cboNewLang.AddItem "Spanish"
  cboNewLang.AddItem "Spanish (Argentina)"
  cboNewLang.AddItem "Spanish (Bolivia)"
  cboNewLang.AddItem "Spanish (Chile)"
  cboNewLang.AddItem "Spanish (Colombia)"
  cboNewLang.AddItem "Spanish (Costa Rica)"
  cboNewLang.AddItem "Spanish (Dominican republic)"
  cboNewLang.AddItem "Spanish (Ecuador)"
  cboNewLang.AddItem "Spanish (El Salvador)"
  cboNewLang.AddItem "Spanish (Guatemala)"
  cboNewLang.AddItem "Spanish (Honduras)"
  cboNewLang.AddItem "Spanish (Mexican)"
  cboNewLang.AddItem "Spanish (Modern)"
  cboNewLang.AddItem "Spanish (Nicaragua)"
  cboNewLang.AddItem "Spanish (Panama)"
  cboNewLang.AddItem "Spanish (Paraguay)"
  cboNewLang.AddItem "Spanish (Peru)"
  cboNewLang.AddItem "Spanish (Puerto Rico)"
  cboNewLang.AddItem "Spanish (U.S.)"
  cboNewLang.AddItem "Spanish (Uruguay)"
  cboNewLang.AddItem "Spanish (Venezuela)"
  cboNewLang.AddItem "Sundanese"
  cboNewLang.AddItem "Swahili"
  cboNewLang.AddItem "Swedish"
  cboNewLang.AddItem "Swedish (Finland)"
  cboNewLang.AddItem "Tagalog"
  cboNewLang.AddItem "Tajik"
  cboNewLang.AddItem "Tamil"
  cboNewLang.AddItem "Tatar"
  cboNewLang.AddItem "Telugu"
  cboNewLang.AddItem "Thai"
  cboNewLang.AddItem "Tibetan"
  cboNewLang.AddItem "Tigrinya"
  cboNewLang.AddItem "Tonga"
  cboNewLang.AddItem "Tsonga"
  cboNewLang.AddItem "Turkish"
  cboNewLang.AddItem "Turkmen"
  cboNewLang.AddItem "Twi"
  cboNewLang.AddItem "Uighur"
  cboNewLang.AddItem "Ukrainian"
  cboNewLang.AddItem "Urdu"
  cboNewLang.AddItem "Urdu (India)"
  cboNewLang.AddItem "Urdu (Pakistan)"
  cboNewLang.AddItem "Uzbek"
  cboNewLang.AddItem "Uzbek (Cyrillic)"
  cboNewLang.AddItem "Uzbek (Latin)"
  cboNewLang.AddItem "Vietnamese"
  cboNewLang.AddItem "Volapuk"
  cboNewLang.AddItem "Welsh"
  cboNewLang.AddItem "Wolof"
  cboNewLang.AddItem "Xhosa"
  cboNewLang.AddItem "Yiddish"
  cboNewLang.AddItem "Yoruba"
  cboNewLang.AddItem "Zhuang"
  cboNewLang.AddItem "Zulu"

  'add language in directory
  cboLang.Clear
  szTemp = Dir(App.Path & "\*.lng")
  While szTemp <> ""
    szTemp = Left(szTemp, Len(szTemp) - 4)
    cboLang.AddItem szTemp
    
    'remove elementi in all language
    For ii = 0 To cboNewLang.ListCount
      If LCase(cboNewLang.List(ii)) = LCase(szTemp) Then
        cboNewLang.RemoveItem ii
        Exit For
      End If
    Next
    
    szTemp = Dir
  Wend
  If cboLang.ListCount > 0 And cboLang.ListIndex = -1 Then cboLang.ListIndex = 0
  If cboNewLang.ListCount > 0 Then
    cboNewLang.ListIndex = 0
    cboNewLang.Enabled = True
    cmdAddLang.Enabled = True
  Else
    cboNewLang.Enabled = False
    cmdAddLang.Enabled = False
  End If

  mnuToolsExtractString.Enabled = inIDE
  
  PatchForm Me
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.Initialise"
End Sub

Private Sub cboLang_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.cboLang_Click()", etFullDebug

Dim ii As Integer
Dim lvItem As ListItem

  LoadFileLang App.Path & "\" & cboLang.Text & ".lng", objDataLang
  lv.ListItems.Clear
  For ii = 1 To UBound(objDataLang)
    With objDataLang(ii)
      Set lvItem = lv.ListItems.Add(, , .MsgId)
      If .MsgStrValid Then lvItem.SubItems(1) = .MsgStr
    End With
  Next
  SetPBar
  cmdSave.Enabled = True
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.cboLang_Click"
End Sub

'add new lang
Private Sub cmdAddLang_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.cmdAddLang_Click()", etFullDebug

  'copy file template in new lang
  FileCopy App.Path & "\" & TEMPLATE_FILE_LANG, App.Path & "\" & cboNewLang.Text & ".lng"
  Initialise

  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.cmdAddLang_Click"
End Sub

Private Sub cmdSave_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.cmdSave_Click()", etFullDebug

  If MsgBox(§§TrasLang§§("Save traslation file. Are you sure you wish to continue?"), vbQuestion + vbYesNo) = vbNo Then Exit Sub
  If Not SaveFileLang(cboLang.Text, objDataLang) Then
    MsgBox §§TrasLang§§("Problem to save file!"), vbExclamation
    Exit Sub
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.cmdSave_Click"
End Sub

Private Sub Form_Resize()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.Form_Resize()", etFullDebug

  lv.Width = Me.Width - 190
  txtOriginal.Width = lv.Width
  txtTraslate.Width = lv.Width
  If Me.Height > 8796 Then Me.Height = 8796
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.Form_Resize"
End Sub

Private Sub lv_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.lv_DblClick()", etFullDebug

  If lv.SelectedItem Is Nothing Then Exit Sub
  txtOriginal.Text = lv.SelectedItem.Text
  txtTraslate.Text = lv.SelectedItem.SubItems(1)
  txtTraslate.SetFocus
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.lv_DblClick"
End Sub

Private Sub mnuToolsExtractString_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.mnuToolsExtractString_Click()", etFullDebug

  If MsgBox(§§TrasLang§§("Do you wish to create a Language template file?"), vbQuestion + vbYesNo, §§TrasLang§§("Reset Sequences")) = vbNo Then Exit Sub
  ExtractStringFromSource
  MsgBox §§TrasLang§§("Success!!")
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.mnuToolsExtractString_Click"
End Sub

Private Sub mnuToolsMergeAllFile_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.mnuToolsMergeAllFile_Click()", etFullDebug

Dim ii As Integer

  If MsgBox(§§TrasLang§§("Do you wish to merge all Language file?"), vbQuestion + vbYesNo, §§TrasLang§§("Merge Language file")) = vbNo Then Exit Sub
  For ii = 0 To cboLang.ListCount - 1
    MergeLangFileString cboLang.List(ii)
  Next
  Initialise
  MsgBox §§TrasLang§§("Success!!")
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.mnuToolsMergeAllFile_Click"
End Sub

Private Sub mnuToolsMergeFile_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.mnuToolsMergeFile_Click()", etFullDebug

  If MsgBox(§§TrasLang§§("Do you wish to merge a Language file '") & cboLang.Text & "' ?", vbQuestion + vbYesNo, §§TrasLang§§("Merge Language file")) = vbNo Then Exit Sub
  MergeLangFileString cboLang.Text
  Initialise
  MsgBox §§TrasLang§§("Success!!")
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.mnuToolsMergeFile_Click"
End Sub

Private Sub txtTraslate_KeyUp(KeyCode As Integer, Shift As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.txtTraslate_KeyUp(" & KeyCode & "," & Shift & ")", etFullDebug

  If lv.SelectedItem Is Nothing Then Exit Sub
  
  If KeyCode = vbKeyPageDown Then
    If lv.SelectedItem.Index < lv.ListItems.Count Then
      With lv.ListItems(lv.SelectedItem.Index + 1)
        .Selected = True
        .EnsureVisible
      End With
      lv_Click
    End If
  ElseIf KeyCode = vbKeyPageUp Then
    If lv.SelectedItem.Index > 1 Then
      With lv.ListItems(lv.SelectedItem.Index - 1)
        .Selected = True
        .EnsureVisible
      End With
      lv_Click
    End If
  ElseIf KeyCode = vbKeyF4 Then
    txtTraslate.Text = txtOriginal.Text
  Else
    lv.ListItems(lv.SelectedItem.Index).SubItems(1) = txtTraslate.Text
    With objDataLang(lv.SelectedItem.Index)
      .MsgStr = txtTraslate.Text
      .MsgStrValid = (Len(txtTraslate.Text) > 0)
    End With
    SetPBar
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.txtTraslate_KeyUp"
End Sub

Private Sub SetPBar()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTrasLang.SetPBar()", etFullDebug

Dim ii As Integer
Dim iVal As Integer
  
  PBarCompleted.Min = 0
  PBarCompleted.Max = lv.ListItems.Count
  iVal = 0
  For ii = 1 To lv.ListItems.Count
    If Len(lv.ListItems(ii).SubItems(1)) > 0 Then
      iVal = iVal + 1
      lv.ListItems(ii).Icon = "property"
      lv.ListItems(ii).SmallIcon = "property"
    Else
      lv.ListItems(ii).Icon = LoadPicture("")
      lv.ListItems(ii).SmallIcon = LoadPicture("")
    End If
  Next
  PBarCompleted.Value = iVal
  lblCompleted.Caption = PBarCompleted.Value & §§TrasLang§§(" of ") & PBarCompleted.Max & "   " & Int(iVal / PBarCompleted.Max * 100) & "%"
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTrasLang.SetPBar"
End Sub


