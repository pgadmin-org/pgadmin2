VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tip of the Day"
   ClientHeight    =   3285
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5190
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Do you want tips to be shown at startup?"
      Top             =   2925
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   3915
      TabIndex        =   2
      Top             =   585
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":0A02
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3915
      TabIndex        =   0
      Top             =   135
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' frmTip.frm - Tip of the day

Option Explicit

Dim colTips As New Collection


Private Sub DoNextTip()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.DoNextTip()", etFullDebug

Dim lTip As Long

  lTip = Int((colTips.Count * Rnd) + 1)
  Display lTip
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.DoNextTip"
End Sub

Function LoadTips() As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.LoadTips()", etFullDebug

Dim szNextTip As String
Dim fNum As Integer
    
  fNum = FreeFile
    
  'Check for the file
  If Dir(App.Path & "\Tips.txt") = "" Then
    LoadTips = False
    Exit Function
  End If
    
  'Load the tips
  Open App.Path & "\Tips.txt" For Input As fNum
  While Not EOF(fNum)
    Line Input #fNum, szNextTip
    colTips.Add szNextTip
  Wend
  Close fNum

  'Display a tip
  DoNextTip
  LoadTips = True
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.LoadTips"
End Function

Private Sub chkLoadTipsAtStartup_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.chkLoadTipsAtStartup_Click()", etFullDebug

  If chkLoadTipsAtStartup.Value = 1 Then
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tips", regString, "Y"
  Else
    RegWrite HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tips", regString, "N"
  End If
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.chkLoadTipsAtStartup_Click"
End Sub

Private Sub cmdNextTip_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.cmdNextTip_Click()", etFullDebug

  DoNextTip
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.cmdNextTip_Click"
End Sub

Private Sub cmdOK_Click()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.cmdOK_Click()", etFullDebug

  Unload Me

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.cmdOK_Click"
End Sub

Private Sub Form_Load()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.Form_Load()", etFullDebug

  PatchForm Me
  
  'See if we should be shown at startup
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tips", "Y")) = "Y" Then
    chkLoadTipsAtStartup.Value = 1
  Else
    chkLoadTipsAtStartup.Value = 0
  End If
    
  Randomize
    
  'Load tips
  If LoadTips() = False Then
    lblTipText.Caption = "That the Tips.txt file was not found? " & vbCrLf & vbCrLf & _
                         "Create a text file named Tips.txt using NotePad with 1 tip per line, " & _
                         "then place it in:" & vbCrLf & vbCrLf & App.Path & "\"
  End If


  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.Form_Load"
End Sub

Public Sub Display(lTip As Long)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":frmTips.Display(" & lTip & ")", etFullDebug

  If colTips.Count > 0 Then
    lblTipText.Caption = colTips.Item(lTip)
  End If

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":frmTips.Display"
End Sub
