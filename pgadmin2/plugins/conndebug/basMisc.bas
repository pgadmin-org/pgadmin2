Attribute VB_Name = "basMisc"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

Public Sub LogError(lError As Long, szError As String, szRoutine As String)
'No logging here, if anythings going wrong then we want the real error

  svr.LogEvent "Error in " & App.Title & ":" & szRoutine & ": " & lError & " - " & szError, etErrors
  MsgBox "An error has occured in " & App.Title & ":" & szRoutine & ":" & vbCrLf & vbCrLf & "Number: " & lError & vbCrLf & "Description: " & szError, vbExclamation, App.Title & " Error"
  
End Sub

Public Sub StartMsg(ByVal szMsg As String)
'Logging code, so no internal logging...

  svr.LogEvent szMsg, etMiniDebug
  Screen.MousePointer = vbHourglass
  sb.Panels("info").Text = szMsg
  sb.Refresh
  sTimer = Timer
  
End Sub

Public Sub EndMsg()
'Logging code, so no internal logging...

Dim szMsg As String

  szMsg = "Done - " & Fix((Timer - sTimer) * 100) / 100 & " Secs."
  If Right(sb.Panels("info").Text, 5) <> "Done." Then
    svr.LogEvent szMsg, etMiniDebug
    sb.Panels("timer").Text = Fix((Timer - sTimer) * 100) / 100 & " Secs."
    sb.Panels("info").Text = sb.Panels("info").Text & " Done."
    sb.Refresh
  End If
  Screen.MousePointer = vbDefault
  
End Sub

Public Function NetInitialise() As String
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.NetInitialise()", etFullDebug

Dim udtWSAData As WSADATA

  If WSAStartup(257, udtWSAData) <> 0 Then
    NetInitialise = "Unable to initialize Winsock"
    Exit Function
  End If
  hICMP = IcmpCreateFile()
  
  If hICMP = 0 Then
    NetInitialise = "Unable to initialize ICMP"
    Exit Function
  End If
  NetInitialise = MyHostName
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.NetInitialise"
End Function

Public Function MyHostName() As String
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.MyHostName()", etFullDebug

Dim szTemp As String
Dim X As Long
  
  szTemp = Space$(256)
  X = gethostname(szTemp, Len(szTemp))
  X = InStr(szTemp, vbNullChar)
  If X > 0 Then szTemp = Left$(szTemp, X - 1)
  MyHostName = szTemp
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.MyHostName"
End Function

Public Function GetIPFromHostName(ByVal szHostName As String) As String
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.GetIPFromHostName(" & QUOTE & szHostName & QUOTE & ")", etFullDebug

Dim lHosent As Long
Dim lName As Long
Dim lAddress As Long
Dim lIPAddress As Long
   
  lHosent = gethostbyname(szHostName & vbNullChar)

  If lHosent <> 0 Then
    lName = lHosent
    lAddress = lHosent + 12
    MemCopy lAddress, ByVal lAddress, 4
    MemCopy lIPAddress, ByVal lAddress, 4
    MemCopy lAddress, ByVal lIPAddress, 4
    GetIPFromHostName = AddrToIP(lAddress)
  End If
    
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.GetIPFromHostName"
End Function

Public Function AddrToIP(ByVal lAddrOrIP As Long) As String
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.AddrToIP(" & lAddrOrIP & ")", etFullDebug

Dim lString As Long
   
  lString = inet_ntoa(lAddrOrIP)
  AddrToIP = GetStrFromPtrA(lString)
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.AddrToIP"
End Function

Public Function GetHostByAddress(ByVal lAddr As Long) As String
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.GetHostByAddress(" & lAddr & ")", etFullDebug

Dim lBytes As Long
Dim lHostEnt As Long
Dim szIP As String
   
  lHostEnt = gethostbyaddr(lAddr, 4, AF_INET)
         
  If lHostEnt <> 0 Then
    MemCopy lHostEnt, ByVal lHostEnt, 4
    lBytes = lstrlenA(ByVal lHostEnt)
    
    If lBytes > 0 Then
      szIP = Space$(lBytes)
      MemCopy ByVal szIP, ByVal lHostEnt, lBytes
      GetHostByAddress = szIP
    End If

   Else
      GetHostByAddress = ""
   End If
  
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.GetHostByAddress"
End Function

Public Sub NetShutDown()
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.NetShutDown()", etFullDebug

  If hICMP Then IcmpCloseHandle (hICMP)
  WSACleanup
    
  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.NetShutDown"
End Sub

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String
On Error GoTo Err_Handler
svr.LogEvent "Entering " & App.Title & ":basMisc.GetStrFromPtrA(" & lpszA & ")", etFullDebug

  GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
  lstrcpyA ByVal GetStrFromPtrA, ByVal lpszA
     
  Exit Function
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basMisc.GetStrFromPtrA"
End Function
