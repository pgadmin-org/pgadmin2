Attribute VB_Name = "basMisc"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basMisc.bas - Contains miscellaneous functions and subroutines.

Option Explicit

Public Sub Main()

Dim sStart As Single
Dim szFrequency As String
Dim szFont() As String
Dim objFont As New StdFont
  
  'Show the splash screen
  Load frmSplash
  frmSplash.Show
  frmSplash.Refresh
  sStart = Timer
  
  'Load the main form
  Load frmMain
  frmMain.Visible = False
  
  'Create the Server Object
  Set frmMain.svr = New pgServer
 
  'Setup the logging and log the startup. Set DontLogErrors to prevent pgSchema errors
  'being logged internally in pgSchema, as they will go through the error traps here.
  frmMain.svr.DontLogErrors = True
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Mask Password", "Y")) = "Y" Then
    frmMain.svr.ShowPassword = False
  Else
    frmMain.svr.ShowPassword = True
  End If
  frmMain.svr.Logfile = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log File", "C:\" & App.Title & "_%ID.Log")
  
  'Store the log level locally, otherwise we get in a loop
  ctx.LogLevel = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Log Level", "2"))
  frmMain.svr.LogLevel = ctx.LogLevel
  
  'Display the log view first if required
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Log Window", "Visible", "N")) = "Y" Then
    Load frmLog
    frmLog.Show
    ctx.LogView = True
    frmMain.mnuViewShowLogWindow.Checked = True
  Else
    ctx.LogView = False
    frmMain.mnuViewShowLogWindow.Checked = False
  End If
  
  frmMain.svr.LogEvent "###################################################################", etMiniDebug
  frmMain.svr.LogEvent App.Title & " v" & App.Major & "." & App.Minor & " Build " & App.Revision & " Startup", etMiniDebug
  frmMain.svr.LogEvent "###################################################################", etMiniDebug
  
  'Show system objects. The Server object will always include them.
  frmMain.svr.IncludeSys = True
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Hide System Objects", "Y")) = "Y" Then
    ctx.IncludeSys = False
    frmMain.mnuViewSystemObjects.Checked = False
  Else
    ctx.IncludeSys = True
    frmMain.mnuViewSystemObjects.Checked = True
  End If
  
  'Encrypted Passwords?
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Encrypt Passwords", "Y")) = "Y" Then
    frmMain.svr.EncryptPasswords = True
  Else
    frmMain.svr.EncryptPasswords = False
  End If
  
  'Auto Row Count?
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Auto Row Count", "Y")) = "Y" Then
    ctx.AutoRowCount = True
  Else
    ctx.AutoRowCount = False
  End If
  
  'Defer Connections?
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Defer Connection", "Y")) = "Y" Then
    frmMain.svr.DeferConnection = True
  Else
    frmMain.svr.DeferConnection = False
  End If
  
  'Display/Hide the StausBar/ToolBar/Definition Pane
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Status Bar", "Y")) = "Y" Then
    frmMain.sb.Visible = True
    frmMain.mnuViewShowStatusBar.Checked = True
  Else
    frmMain.sb.Visible = False
    frmMain.mnuViewShowStatusBar.Checked = False
  End If
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tool Bar", "Y")) = "Y" Then
    frmMain.tb.Visible = True
    frmMain.mnuViewShowToolBar.Checked = True
  Else
    frmMain.tb.Visible = False
    frmMain.mnuViewShowToolBar.Checked = False
  End If
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Definition Pane", "Y")) = "Y" Then
    frmMain.txtDefinition.Visible = True
    frmMain.mnuViewShowDefinitionPane.Checked = True
  Else
    frmMain.txtDefinition.Visible = False
    frmMain.mnuViewShowDefinitionPane.Checked = False
  End If
  
  'Position & Size the form
  frmMain.Left = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Left", "0"))
  frmMain.Top = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Top", "0"))
  frmMain.Width = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Width", "9500"))
  frmMain.Height = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Height", "7000"))
  frmMain.Resize Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Vertical Splitter", "3500")), Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Horizontal Splitter", "5000"))
  frmMain.Caption = App.Title & " v" & App.Major & "." & App.Minor & " Build " & App.Revision


  'Build the connection menu
  BuildConnectionMenu
  
  'Build the Plugins Menu
  BuildPluginsMenu
  
  'Get the AutoHighlight colours
  ctx.AutoHighlight = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "AutoHighlight", DEFAULT_AUTOHIGHLIGHT)
  frmMain.txtDefinition.Wordlist = ctx.AutoHighlight
  
  'Get the Font
  szFont = Split(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Font", "MS Sans Serif|8|False|False"), "|")
  objFont.Name = szFont(0)
  objFont.Size = Val(szFont(1))
  objFont.Bold = CBool(szFont(2))
  objFont.Italic = CBool(szFont(3))
  Set ctx.Font = objFont
  Set frmMain.txtDefinition.Font = ctx.Font
  Set frmMain.tv.Font = ctx.Font
  Set frmMain.lv.Font = ctx.Font
  
  'Hide the splash screen
  Do Until Timer > sStart + 2
    DoEvents
  Loop

  Unload frmSplash
  
  'Show the main form.
  frmMain.Show
  
  'Show the Upgrade Wizard if required.
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Check", "Y")) = "Y" Then
    Select Case UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Frequency", "Week"))
      Case "DAY"
        szFrequency = "d"
      Case "WEEK"
        szFrequency = "ww"
      Case "MONTH"
        szFrequency = "m"
      Case "YEAR"
        szFrequency = "yyyy"
    End Select
    If DateAdd(szFrequency, 1, CDate(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Auto Upgrade", "Last Check", "2000-01-01"))) <= Date Then
      Load frmUpgradeWizard
      If InStr(0, Command, "-wine") <> 0 Then
        frmUpgradeWizard.Show
      Else
        frmUpgradeWizard.Show vbModal, frmMain
      End If
    End If
  End If
  
  'Show the Tips if required.
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Tips", "Y")) = "Y" Then
    Load frmTip
    If InStr(0, Command, "-wine") <> 0 Then
      frmTip.Show
    Else
      frmTip.Show vbModal, frmMain
    End If
  End If
   
End Sub

Public Function GetID() As String
On Error GoTo Err_Handler
'Don't log, accessed *all* the time.

Static lID As Long
  lID = lID + 1
  GetID = lID
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.GetID"
End Function

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.SetTopMostWindow(" & hWnd & ", " & Topmost & ")", etFullDebug

  If Topmost = True Then 'Make the window topmost
    SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  End If
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.SetTopMostWindow"
End Function
 
Public Sub BuildConnectionMenu()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.BuildConnectionMenu()", etFullDebug

Dim X As Integer
Dim szConnection As String
  frmMain.tb.Buttons(1).ButtonMenus.Clear
  For X = 1 To 10
    szConnection = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection " & X, "")
    szConnection = Replace(Replace(szConnection, "|", "@", 1, 1), "|", ":", 1, 1)
    If szConnection <> "" Then frmMain.tb.Buttons("connect").ButtonMenus.Add X, X & "|" & szConnection, szConnection
  Next X
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.BuildConnectionMenu"
End Sub

Public Sub BuildPluginsMenu()
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.BuildPluginsMenu()", etFullDebug

Dim objPlugin As pgPlugin
Dim X As Integer

  'Clear the menu
  frmMain.mnuPlugins.Visible = False
  frmMain.mnuPluginsPlg(0).Visible = True
  For X = 1 To 20
    frmMain.mnuPluginsPlg(X).Caption = "Plugin" & X
    frmMain.mnuPluginsPlg(X).Visible = False
  Next X
  
  'Load new plugins
  X = 1
  For Each objPlugin In plg
    If Not ((frmMain.svr.ConnectionString = "") And (objPlugin.PluginType = 1)) Then
      frmMain.mnuPluginsPlg(X).Caption = objPlugin.Description & "..."
      frmMain.mnuPluginsPlg(X).Visible = True
      X = X + 1
      frmMain.mnuPluginsPlg(0).Visible = False
    
      'Bomb out if there's more than 20 Plugins
      If X > 20 Then
        MsgBox App.Title & " currently only supports a maximum of 20 plugins loaded at the same time. Please email the Support mailing list listed in the Helpfile and let the developers know that you've exceeded this limit.", vbExclamation, "Error"
        Exit Sub
      End If
    End If
  Next objPlugin
  frmMain.mnuPluginsPlg(0).Visible = False
  If X > 1 Then frmMain.mnuPlugins.Visible = True

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.BuildPluginsMenu"
End Sub

Public Sub LogError(lError As Long, szError As String, szRoutine As String)
'No logging here, if anythings going wrong then we want the real error

  frmMain.svr.LogEvent "Error in " & szRoutine & ": " & lError & " - " & szError, etErrors
  MsgBox "An error has occured in " & szRoutine & ":" & vbCrLf & vbCrLf & "Number: " & lError & vbCrLf & "Description: " & szError, vbExclamation, App.Title & " Error"
  
  'If we are between StartMsg/EndMsg, call EndMsg with errors
  If Screen.MousePointer = vbHourglass Then
    EndMsg " with errors"
  Else
    frmMain.sb.Panels("info").Text = "An error has occured."
  End If
  
End Sub

Public Sub StartMsg(ByVal szMsg As String)
'Logging code, so no internal logging...

  frmMain.svr.LogEvent szMsg, etMiniDebug
  Screen.MousePointer = vbHourglass
  frmMain.sb.Panels("info").Text = szMsg
  frmMain.sb.Refresh
  sTimer = Timer
  
End Sub

Public Sub EndMsg(Optional ByVal szErr As String)
'Logging code, so no internal logging...

Dim szMsg As String
   
  szMsg = "Done" & szErr & " - " & Fix((Timer - sTimer) * 100) / 100 & " Secs."
  If InStr(1, frmMain.sb.Panels("info").Text, "Done") = 0 Then
    frmMain.svr.LogEvent szMsg, etMiniDebug
    frmMain.sb.Panels("timer").Text = Fix((Timer - sTimer) * 100) / 100 & " Secs."
    frmMain.sb.Panels("info").Text = frmMain.sb.Panels("info").Text & " Done" & szErr & "." 'szMsg '" Done."
    frmMain.sb.Refresh
  End If
  Screen.MousePointer = vbDefault
  
End Sub

Public Function dbSZ(szData As String) As String
'Don't log this - it needs to be fast and it's unlikely to go wrong...

  szData = Replace(szData, "\", "\\")
  szData = Replace(szData, "'", "\'")
  dbSZ = szData

End Function

Public Function Bool2Bin(bData As Boolean) As Integer
'Don't log this - it needs to be fast and it's unlikely to go wrong...

  If bData Then
    Bool2Bin = 1
  Else
    Bool2Bin = 0
  End If

End Function

Public Function Bin2Bool(iData As Integer) As Boolean
'Don't log this - it needs to be fast and it's unlikely to go wrong...

  If iData = 1 Then
    Bin2Bool = True
  Else
    Bin2Bool = False
  End If

End Function

'Parse an ACL and return | delimited User/Access lists
Public Sub ParseACL(ByVal szACL As String, ByRef szUserlist As String, ByRef szAccesslist As String)
On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.ParseACL(" & QUOTE & szACL & QUOTE & ", " & QUOTE & szUserlist & QUOTE & ", " & QUOTE & szAccesslist & QUOTE & ")", etFullDebug

Dim szEntries() As String
Dim szEntry As Variant
Dim szName As String
Dim szAccess As String
Dim szSQL As String
Dim szTemp As String
  
  szUserlist = ""
  szAccesslist = ""
  If szACL = "" Then Exit Sub
  szACL = Mid(szACL, 2, Len(szACL) - 2)
  szACL = Replace(szACL, QUOTE, "")
  szEntries = Split(szACL, ",")
  For Each szEntry In szEntries
  
    'Get the username
    szName = Left(szEntry, InStr(1, szEntry, "=") - 1)
    If szName = "" Then
      szName = "PUBLIC"
    ElseIf Len(szName) > 6 Then
      If Left(UCase(szName), 6) = "GROUP " Then
        szName = "GROUP " & Mid(szName, 7)
      Else
        szName = szName
      End If
    Else
      szName = szName
    End If
    
    'Get the Access
    szAccess = Mid(szEntry, InStr(1, szEntry, "=") + 1)
    szTemp = ""
    
    'ACLs are different in 7.2+
    If frmMain.svr.dbVersion.VersionNum < 7.2 Then
      
      Select Case szAccess
        Case "arwR"
          szAccess = "All"
        Case ""
          szAccess = "None"
        Case Else
          If InStr(1, szAccess, "a") <> 0 Then szTemp = szTemp & "Insert, "
          If InStr(1, szAccess, "r") <> 0 Then szTemp = szTemp & "Select, "
          If InStr(1, szAccess, "w") <> 0 Then szTemp = szTemp & "Update, Delete, "
          If InStr(1, szAccess, "R") <> 0 Then szTemp = szTemp & "Rule, "
          szAccess = Left(szTemp, Len(szTemp) - 2)
      End Select
    
    Else
      
      Select Case szAccess
        Case "arwdRxt"
          szAccess = "All"
        Case ""
          szAccess = "None"
        Case Else
          If InStr(1, szAccess, "a") <> 0 Then szTemp = szTemp & "Insert, "
          If InStr(1, szAccess, "r") <> 0 Then szTemp = szTemp & "Select, "
          If InStr(1, szAccess, "w") <> 0 Then szTemp = szTemp & "Update, "
          If InStr(1, szAccess, "d") <> 0 Then szTemp = szTemp & "Delete, "
          If InStr(1, szAccess, "R") <> 0 Then szTemp = szTemp & "Rule, "
          If InStr(1, szAccess, "x") <> 0 Then szTemp = szTemp & "References, "
          If InStr(1, szAccess, "t") <> 0 Then szTemp = szTemp & "Trigger, "
          szAccess = Left(szTemp, Len(szTemp) - 2)
      End Select
    
    End If

    szUserlist = szUserlist & szName & "|"
    szAccesslist = szAccesslist & szAccess & "|"
    
  Next szEntry
  
  If Len(szUserlist) > 1 Then szUserlist = Left(szUserlist, Len(szUserlist) - 1)
  If Len(szAccesslist) > 1 Then szAccesslist = Left(szAccesslist, Len(szAccesslist) - 1)
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.ParseACL"
End Sub

