Attribute VB_Name = "basMisc"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' basMisc.bas - Contains miscellaneous functions and subroutines.

Option Explicit

Public Sub Main()

Dim szFilename As String
Dim lCount As Long
Dim sStart As Single
Dim szFrequency As String
Dim szFont() As String
Dim objFont As New StdFont
  
  'Where are we running?
  szFilename = String(255, 0)
  lCount = GetModuleFileName(App.hInstance, szFilename, 255)
  szFilename = Left(szFilename, lCount)
  If UCase(Right(szFilename, 7)) = "VB6.EXE" Then
    inIDE = True
  Else
    inIDE = False
  End If
  
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
  Set frmMain.svr.pgApp = New clsPgApp
 
  'Startup language
  InitLang RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Current Lang", "")
 
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
  LoadAutoHighlight
  ctx.AutoHighlight = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "AutoHighlight", szDefaultAutoHighlight)
  frmMain.txtDefinition.Wordlist = ctx.AutoHighlight
  
  'add KeyWord Reserved to schema
  AddKeyWordReserved
  
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
  PatchForm frmMain
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
  
  'Show users for ForPrivileges
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Show Users For Privileges", "Y")) = "Y" Then
    ctx.ShowUsersForPrivileges = True
  Else
    ctx.ShowUsersForPrivileges = False
  End If
  
  'Ask delete object database
  If UCase(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Ask Delete Object Database", "Y")) = "Y" Then
    ctx.AskDeleteObjectDatabase = True
  Else
    ctx.AskDeleteObjectDatabase = False
  End If
  
  'max number of sql command in query
  ctx.MaxNumberSqlQuery = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Max Number Sql Query", "50"))
  
  'max number of Record in View Data
  ctx.MaxRecordViewData = Val(RegRead(HKEY_CURRENT_USER, "Software\" & App.Title, "Row Limit", "1000"))
  
  'Initialise stuff
  InitVarDb
  InitClone
   
End Sub

'Add KeyWord Reserved to schema
Private Sub AddKeyWordReserved()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.AddKeyWordReserved()", etFullDebug

Dim ii As Integer
Dim vData, vData1
  
  Set frmMain.svr.KeyWordReserved = New Collection
  vData = Split(ctx.AutoHighlight, ";")
  For ii = 0 To UBound(vData) - 1
    vData1 = Split(vData(ii), "|")
    frmMain.svr.KeyWordReserved.Add vData1(0), vData1(0)
  Next

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.AddKeyWordReserved"
End Sub

Public Function GetID() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler

'Don't log, accessed *all* the time.

Static lID As Long
  lID = lID + 1
  GetID = lID
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.GetID"
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.SetTopMostWindow(" & hwnd & ", " & Topmost & ")", etFullDebug

  If Topmost = True Then 'Make the window topmost
    SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  End If
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.SetTopMostWindow"
End Function
 
Public Sub BuildConnectionMenu()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.BuildConnectionMenu()", etFullDebug

Dim X As Integer
Dim szConnection As String
Dim vData

  frmMain.tb.Buttons(1).ButtonMenus.Clear
  For X = 1 To 10
    szConnection = RegRead(HKEY_CURRENT_USER, "Software\" & App.Title & "\Connections", "Connection " & X, "")
    If szConnection <> "" Then
      vData = Split(szConnection, "|")
      szConnection = vData(0) & "@" & vData(1) & ":" & vData(2)
      If UBound(vData) > 2 Then szConnection = szConnection & " - " & vData(3)
      frmMain.tb.Buttons("connect").ButtonMenus.Add X, X & "|" & szConnection, szConnection
    End If
  Next X
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.BuildConnectionMenu"
End Sub

Public Sub BuildPluginsMenu()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
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
        MsgBox App.Title & ??TrasLang??(" currently only supports a maximum of 20 plugins loaded at the same time. Please email the Support mailing list listed in the Helpfile and let the developers know that you've exceeded this limit."), vbExclamation, ??TrasLang??("Error")
        Exit Sub
      End If
    End If
  Next objPlugin
  frmMain.mnuPluginsPlg(0).Visible = False
  If X > 1 Then frmMain.mnuPlugins.Visible = True

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.BuildPluginsMenu"
End Sub

Public Sub LogError(lError As Long, szError As String, szRoutine As String, Optional bSendMail As Boolean = True)
'No logging here, if anythings going wrong then we want the real error
Dim objErrorForm As New frmError
Dim bShowFormErr As Boolean
Dim vData
Dim szErr As String
  
  szErr = Replace(szError, vbLf, vbCrLf)
  frmMain.svr.LogEvent "Error in " & szRoutine & ": " & lError & " - " & szErr, etErrors

  'find error in ignore error
  bShowFormErr = True
  For Each vData In ColIgnoreError
    If vData = szRoutine & "_" & lError & "_" & szErr Then
      bShowFormErr = False
      Exit For
    End If
  Next

  If bShowFormErr Then
    Load objErrorForm
    objErrorForm.Initialise lError, szErr, szRoutine, bSendMail
    objErrorForm.Show vbModal
  End If
  
  'If we are between StartMsg/EndMsg, call EndMsg with errors
  If Screen.MousePointer = vbHourglass Then
    EndMsg " with errors"
  Else
    frmMain.sb.Panels("info").Text = ??TrasLang??("An error has occured.")
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
   
  szMsg = ??TrasLang??("Done") & szErr & " - " & Fix((Timer - sTimer) * 100) / 100 & ??TrasLang??(" Secs.")
  If InStr(1, frmMain.sb.Panels("info").Text, ??TrasLang??("Done")) = 0 Then
    frmMain.svr.LogEvent szMsg, etMiniDebug
    frmMain.sb.Panels("timer").Text = Fix((Timer - sTimer) * 100) / 100 & ??TrasLang??(" Secs.")
    frmMain.sb.Panels("info").Text = frmMain.sb.Panels("info").Text & ??TrasLang??(" Done") & szErr & "." 'szMsg '" Done."
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

'Format an identifier as required
'This code is based on fmtID from the pg_dump code
Public Function fmtID(ByVal szData As String) As String
On Error Resume Next

Dim X As Integer
Dim iVal As Integer
Dim bFound As Boolean

  'Replace double quotes
  szData = Replace(szData, Quote, Quote & Quote)

  'verify KeyWord Reserved
  For X = 1 To frmMain.svr.KeyWordReserved.Count
    If LCase(frmMain.svr.KeyWordReserved(X)) = LCase(szData) Then
      szData = QUOTE & szData & QUOTE
      bFound = True
      Exit For
    End If
  Next

  If Not bFound Then
    If IsNumeric(szData) Then
      szData = Quote & szData & Quote
    Else
      For X = 1 To Len(szData)
        iVal = Asc(Mid(szData, X, 1))
        If Not ((iVal >= 48) And (iVal <= 57)) And _
           Not ((iVal >= 97) And (iVal <= 122)) And _
           Not (iVal = 95) Then
          szData = Quote & szData & Quote
          Exit For
        End If
      Next X
    End If
  End If

  fmtID = szData

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
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.ParseACL(" & Quote & szACL & Quote & ", " & Quote & szUserlist & Quote & ", " & Quote & szAccesslist & Quote & ")", etFullDebug

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
  szACL = Replace(szACL, Quote, "")
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
    If ctx.dbVer < 7.2 Then
      
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
          If InStr(1, szAccess, "C") <> 0 Then szTemp = szTemp & "Create, "
          If InStr(1, szAccess, "T") <> 0 Then szTemp = szTemp & "Temp, "
          If InStr(1, szAccess, "U") <> 0 Then szTemp = szTemp & "Usage, "
          If InStr(1, szAccess, "X") <> 0 Then szTemp = szTemp & "Execute, "
          If Len(szTemp) > 2 Then szAccess = Left(szTemp, Len(szTemp) - 2)
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

'Format a typename
Public Function fmtTypeName(objType As pgType) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.fmtTypeName(" & objType.FormattedID & ")", etFullDebug

Dim szTemp As String

  If ctx.dbVer >= 7.3 And objType.Namespace <> "pg_catalog" Then
    If objType.Element <> "" And objType.InternalLength = -1 Then 'Array Type
      szTemp = objType.Element & "[]"
    Else
      szTemp = fmtID(objType.Namespace) & "." & fmtID(objType.Name)
    End If
  Else
    If objType.Element <> "" And objType.InternalLength = -1 Then 'Array Type
      szTemp = fmtID(objType.Element) & "[]"
    Else
      szTemp = fmtID(objType.Name)
    End If
  End If
  
  fmtTypeName = szTemp
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.fmtTypeName"
End Function

Function MakeISODate(vDate As Variant) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.MakeISODate(" & vDate & ")", etFullDebug

  'If we can't figure it out, just return a string
  If Not IsDate(vDate) Then
    MakeISODate = CStr(vDate)
    Exit Function
  End If
  
  MakeISODate = Year(vDate) & "-" & Month(vDate) & "-" & Day(vDate)
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.MakeISODate"
End Function

Function MakeISOTimestamp(vTimestamp As Variant) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.MakeISOTimestamp(" & vTimestamp & ")", etFullDebug

  'If we can't figure it out, just return a string
  If Not IsDate(vTimestamp) Then
    MakeISOTimestamp = CStr(vTimestamp)
    Exit Function
  End If
  
  MakeISOTimestamp = Year(vTimestamp) & "-" & Month(vTimestamp) & "-" & Day(vTimestamp) & " " & Hour(vTimestamp) & ":" & Minute(vTimestamp) & ":" & Second(vTimestamp)
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.MakeISOTimestamp"
End Function

Public Sub AutoSizeColumnLv(lv As ListView)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.AutoSizeColumnLv(" & lv.Name & ")", etFullDebug
Dim ii As Integer
Dim szKey As String
Dim objItem As ListItem

  With lv
    If .ListItems.Count <= 0 Then Exit Sub
    szKey = CStr(Now)

    'frank_lupo add new element title in listview
    Set objItem = .ListItems.Add(1, szKey, .ColumnHeaders(1).Text & "  ")
    SendMessage .hwnd, LVM_SETCOLUMNWIDTH, 0, LVSCW_AUTOSIZE

    For ii = 1 To .ColumnHeaders.Count - 1
      objItem.SubItems(ii) = .ColumnHeaders(ii + 1).Text & "  "
      SendMessage .hwnd, LVM_SETCOLUMNWIDTH, ii, LVSCW_AUTOSIZE
    Next

    'frank_lupo drop element title in listview
    .ListItems.Remove szKey
  End With

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.AutoSizeColumnLv"
End Sub

Public Sub AutoSizeColumnFGrid(Grd As MSFlexGrid)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.AutoSizeColumnFGrid(" & Grd.Name & ")", etFullDebug
Dim lColWidth As Long
Dim lii As Long
Dim ljj As Long
Dim szCellText As String
Dim lLongestLen As Long
Dim szLongestString As String
    
  With Grd
    For lii = 0 To .Cols - 1
      szLongestString = ""
      lLongestLen = 0

      For ljj = 0 To .Rows - 1
        szCellText = .TextMatrix(ljj, lii)

        If Len(szCellText) > lLongestLen Then
          lLongestLen = Len(szCellText)
          szLongestString = szCellText
        End If
      Next
      If Len(szLongestString) > 0 Then
        .ColWidth(lii) = .Parent.TextWidth(szLongestString) + 200
      End If
    Next
  End With

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.AutoSizeColumnFGrid"
End Sub

Public Function NameImageByObjectType(ObjectType As String) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.NameImageByObjectType(" & ObjectType & ")", etFullDebug

  Select Case ObjectType
    Case "Aggregate"
      NameImageByObjectType = "aggregate"
          
    Case "Cast"
      NameImageByObjectType = "cast"
          
    Case "Column"
      NameImageByObjectType = "column"
          
    Case "Database"
      NameImageByObjectType = "database"
          
    Case "Domain"
      NameImageByObjectType = "domain"
          
    Case "Conversion"
      NameImageByObjectType = "conversion"
          
    Case "Foreign Key"
      NameImageByObjectType = "foreignkey"
          
    Case "Function"
      NameImageByObjectType = "function"

    Case "Group"
      NameImageByObjectType = "group"
    
    Case "Index"
      NameImageByObjectType = "index"
          
    Case "Language"
      NameImageByObjectType = "language"
          
    Case "Schema"
      NameImageByObjectType = "namespace"
          
    Case "Operator"
      NameImageByObjectType = "operator"
          
    Case "OperatorClass"
      NameImageByObjectType = "operatorclass"
          
    Case "Rule"
      NameImageByObjectType = "rule"
          
    Case "Server"
      NameImageByObjectType = "server"
          
    Case "Sequence"
      NameImageByObjectType = "sequence"

    Case "Table"
      NameImageByObjectType = "table"
          
    Case "Trigger"
      NameImageByObjectType = "trigger"
        
    Case "Type"
      NameImageByObjectType = "type"
          
    Case "User"
      NameImageByObjectType = "user"
          
    Case "View"
      NameImageByObjectType = "view"
          
    Case Else
      NameImageByObjectType = "property"
        
  End Select

  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.NameImageByObjectType"
End Function

Private Sub LoadAutoHighlight()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.LoadAutoHighlight()", etFullDebug

Dim iFile As Integer
Dim vData

  szDefaultAutoHighlight = ""
  iFile = FreeFile
  Open App.Path & "\KeyWords.txt" For Input As #iFile
  vData = Split(Input(LOF(iFile), #iFile), vbCrLf)
  Close #iFile
  szDefaultAutoHighlight = Join(vData, ";")
  
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.LoadAutoHighlight"
End Sub

'load user/group privileges
Public Sub LoadUGACL(cboUGACL As ImageCombo)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.LoadUGACL(" & cboUGACL.Name & ")", etFullDebug
    
Dim objUser As pgUser
Dim objGroup As pgGroup
    
    cboUGACL.ComboItems.Add , , "PUBLIC", "public"
    If ctx.ShowUsersForPrivileges Then
      For Each objUser In frmMain.svr.Users
        cboUGACL.ComboItems.Add , , objUser.Name, "user"
      Next objUser
    End If
    For Each objGroup In frmMain.svr.Groups
      cboUGACL.ComboItems.Add , , objGroup.Name, "group"
    Next objGroup
    cboUGACL.ComboItems(1).Selected = True

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.LoadUGACL"
End Sub

'Convert Boolean value in yes/no string
Public Function BoolToYesNo(ByVal bData As Boolean) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.BoolToYesNo(" & bData & ")", etFullDebug
  
  BoolToYesNo = IIf(bData, "Yes", "No")
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.BoolToYesNo"
End Function

'Convert yes/no string in Boolean value
Public Function YesNoToBool(ByVal szData As String) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.YesNoToBool(" & szData & ")", etFullDebug
  
  YesNoToBool = (Trim(szData) = "Yes")
  Exit Function
  
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.YesNoToBool"
End Function

Public Function BrowseFolder(hwnd As Long, DialogTitle As String) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.BrowseFolder(" & hwnd & ",'" & DialogTitle & "')", etFullDebug

Dim X As Long
Dim BI As BROWSEINFO
Dim dwIList As Long
Dim szPath As String
Dim wPos As Integer
    
  With BI
    .hOwner = hwnd
    .lpszTitle = DialogTitle
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  dwIList = SHBrowseForFolder(BI)
  szPath = Space$(512)
  X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
  If X Then
    wPos = InStr(szPath, Chr(0))
    BrowseFolder = Left$(szPath, wPos - 1)
  Else
    BrowseFolder = ""
  End If

  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.BrowseFolder"
End Function

Public Function ReadTextFile(ByVal PathFile As String) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.ReadFile(" & Quote & PathFile & Quote & ")", etFullDebug

Dim iFile As Integer

  iFile = FreeFile
  Open PathFile For Input As #iFile
  ReadTextFile = Input(LOF(iFile), #iFile)
  Close #iFile
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.ReadFile"
End Function

Public Function WriteTextFile(ByVal PathFile As String, ByVal szData As String) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basMisc.WriteFile(" & Quote & PathFile & Quote & ")", etFullDebug

Dim iFile As Integer

  iFile = FreeFile
  Open PathFile For Output As #iFile
  Print #iFile, szData
  Close #iFile
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.WriteFile"
End Function

