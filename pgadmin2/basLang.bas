Attribute VB_Name = "basLang"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence
'
' basLang.bas - Contains language traslation functions and subroutines.

Option Explicit

Private Type TypeMsgId
  MsgId As String
  Reference As String
End Type

Private DataMsgId() As TypeMsgId
Private DataLang() As StrLang

Private Const TAG_START_MSGID As String = "<msgid>"
Private Const TAG_START_MSGSTR As String = "<msgstr>"
Private Const FncTrasLang = "§§TrasLang§§("""
Private Const FormCaption = "Caption         ="
Private Const FormTabCaption = "TabCaption("
Private Const FormToolTipText = "ToolTipText     ="

'return the traslation MessageId in specific Lang
' Note: Because these function may be *very* frequently accessed, no logging is performed.
Public Function §§TrasLang§§(ByVal MsgId) As String
Dim szMsgStr As String
Dim ii As Integer

  szMsgStr = MsgId
  For ii = 1 To UBound(DataLang)
    With DataLang(ii)
      If .MsgId = MsgId Then
        If .MsgStrValid Then szMsgStr = .MsgStr
        Exit For
      End If
    End With
  Next
  
  §§TrasLang§§ = szMsgStr
End Function

'Startup Language
Public Sub InitLang(ByVal Lang As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.InitLang(" & Quote & Lang & Quote & ")", etFullDebug

  LoadFileLang App.Path & "\" & Lang & ".lng", DataLang

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.GetImageFromValCast"
End Sub

'Load file lang
Public Sub LoadFileLang(ByVal FileLang As String, ByRef objDataLang() As StrLang)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.LoadFileLang(" & Quote & FileLang & Quote & ")", etFullDebug

Dim vData
Dim ii As Integer
Dim ijj As Integer
  
  ReDim objDataLang(0) As StrLang
  If FileLang = "" Then Exit Sub
  If Dir(FileLang) = "" Then Exit Sub
  
  'Load data
  vData = Split(ReadTextFile(FileLang), vbCrLf)
  For ii = 0 To UBound(vData)
    If Left(vData(ii), 1) <> "#" And vData(ii) <> "" Then
      If Left(vData(ii), Len(TAG_START_MSGID)) = TAG_START_MSGID Then
        ijj = UBound(objDataLang) + 1
        ReDim Preserve objDataLang(ijj) As StrLang
        With objDataLang(ijj)
          .MsgId = Mid(vData(ii), Len(TAG_START_MSGID) + 1)
          .MsgStr = Mid(vData(ii + 1), Len(TAG_START_MSGSTR) + 1)
          .MsgStrValid = (Len(.MsgId) > 0 And Len(.MsgStr) > 0)
          .LineNumber = ii + 1
        End With
        ii = ii + 1
      End If
    End If
  Next

  Exit Sub
Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.LoadFileLang"
End Sub

'Save Lang File
Public Function SaveFileLang(ByVal Lang As String, ByRef objDataLang() As StrLang) As Boolean
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.SaveFileLang(" & Quote & Lang & Quote & ")", etFullDebug

Dim vData
Dim ii As Integer
Dim szFile As String
  
  SaveFileLang = False
  szFile = App.Path & "\" & Lang & ".lng"
  If szFile = "" Then Exit Function
  If Dir(szFile) = "" Then Exit Function
  
  'read file language
  vData = Split(ReadTextFile(szFile), vbCrLf)
  
''''Dim ijj As Integer
''''  'loop verify
''''  For ii = 0 To UBound(vData)
''''    For ijj = 1 To UBound(objDataLang)
''''      If vData(ii) = TAG_START_MSGID & objDataLang(ijj).MsgId Then
''''        ii = ii + 1
''''        vData(ii) = TAG_START_MSGSTR
''''        If objDataLang(ijj).MsgStrValid Then vData(ii) = vData(ii) & objDataLang(ijj).MsgStr
''''        Exit For
''''      End If
''''    Next
''''  Next
  
  'loop verify
  For ii = 1 To UBound(objDataLang)
    With objDataLang(ii)
      vData(.LineNumber) = TAG_START_MSGSTR
      If .MsgStrValid Then vData(.LineNumber) = vData(.LineNumber) & .MsgStr
    End With
  Next
  
  'save file
  WriteTextFile szFile, Join(vData, vbCrLf)
  SaveFileLang = True
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.SaveFileLang"
End Function

'extract string from source
Public Sub ExtractStringFromSource()
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.ExtractStringFromSource()", etFullDebug

Dim szFile As String
Dim szExt As String
Dim vData
Dim ii As Integer
Dim ijj As Integer
Dim iPosStart As Integer
Dim iPosEnd As Integer
Dim szTemp As String
Dim szTemp1 As String
Dim bFound As Boolean

  StartMsg §§TrasLang§§("Extract string from source.....")

  ReDim DataMsgId(0) As TypeMsgId

  'loop directory work
  szFile = Dir(App.Path & "\")
  While szFile <> ""
    szExt = LCase(Right(szFile, 3))
    Select Case szExt
      Case "bas", "frm", "cls", "ctl"
        
        'verify file
        vData = Split(ReadTextFile(App.Path & "\" & szFile), vbCrLf)
        For ii = 0 To UBound(vData)
          DoEvents
          szTemp = vData(ii)
          If InStr(vData(ii), "Private Const FncTrasLang") = 1 Then
            iPosStart = 0
          Else
            iPosStart = InStr(szTemp, FncTrasLang)
          End If

          While iPosStart > 0
            szTemp1 = Mid(szTemp, iPosStart + Len(FncTrasLang))
            iPosEnd = InStr(szTemp1, Quote & ")") - 1
            szTemp1 = Mid(szTemp1, 1, iPosEnd)
    
            szTemp = Mid(szTemp, iPosStart + Len(FncTrasLang) + iPosEnd)
    
            'verify if string exists
            StringExists szTemp1, szFile, ii
    
            iPosStart = InStr(szTemp, FncTrasLang)
          Wend
          
          'verify definition form
          Select Case szExt
            Case "frm", "ctl"
              szTemp = LTrim(vData(ii))
  
              szTemp1 = ""
              bFound = False
              If Left(szTemp, Len(FormCaption)) = FormCaption Then
                'caption
                szTemp = Trim(Mid(szTemp, Len(FormCaption) + 1))
                bFound = True
                szTemp1 = "Caption"
              ElseIf Left(szTemp, Len(FormTabCaption)) = FormTabCaption Then
                'tabcaption
                szTemp = Trim(Mid(szTemp, Len(FormTabCaption) + 1))
                szTemp = Mid(szTemp, InStr(szTemp, Quote))
                bFound = True
                szTemp1 = "TabCaption"
              ElseIf Left(szTemp, Len(FormToolTipText)) = FormToolTipText Then
                'ToolTipText
                szTemp = Trim(Mid(szTemp, Len(FormToolTipText) + 1))
                bFound = True
                szTemp1 = "ToolTipText"
              End If
              
              If bFound Then
                If Left(szTemp, 1) = "$" Then
                  'the text are in frx file
                  szTemp1 = szTemp1 & "->frx"
                  szTemp = GetStringFromFrx(szTemp)
                Else
                  szTemp = Mid(szTemp, 2)
                  szTemp = Left(szTemp, Len(szTemp) - 1)
                End If
                If Len(Trim(szTemp)) > 0 Then
                  StringExists szTemp, szTemp1 & "->" & szFile, ii
                End If
              End If
          
          End Select
        Next
    End Select
    szFile = Dir
  Wend

  'output
  szTemp = "# pgAdmin II - PostgreSQL Tools" & vbCrLf
  szTemp = szTemp & "# Copyright (C) 2001 - 2003, The pgAdmin Development Team" & vbCrLf
  szTemp = szTemp & "# This software is released under the pgAdmin Public Licence" & vbCrLf
  szTemp = szTemp & "# pgAdmin Developers <pgadmin-hackers@postgresql.org>" & vbCrLf

  For ii = 1 To UBound(DataMsgId)
    vData = Split(DataMsgId(ii).Reference)
    szTemp1 = ""
    For ijj = 0 To UBound(vData)
      If Fix(ijj / 3) = ijj / 3 Then
        szTemp1 = szTemp1 & vbCrLf
        szTemp1 = szTemp1 & "#:"
      End If
      szTemp1 = szTemp1 & " " & vData(ijj)
    Next
    szTemp = szTemp & szTemp1 & vbCrLf
    szTemp = szTemp & "<msgid>" & DataMsgId(ii).MsgId & vbCrLf
    szTemp = szTemp & "<msgstr>" & vbCrLf
  Next
  
  WriteTextFile App.Path & "\" & TEMPLATE_FILE_LANG, szTemp
  EndMsg

  Exit Sub
Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.ExtractStringFromSource"
End Sub

'verify if string exists
Private Sub StringExists(ByVal StrMsg As String, ByVal NameFile As String, ByVal LineNumber As Integer)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.GetStringFrx(" & Quote & StrMsg & Quote & "," & Quote & NameFile & Quote & "," & LineNumber & ")", etFullDebug

Dim bFound As Boolean
Dim ijj As Integer
Dim iBoundData As Integer

  iBoundData = UBound(DataMsgId)

  'verify if string exists
  bFound = False
  For ijj = 1 To iBoundData
    With DataMsgId(ijj)
      If .MsgId = StrMsg Then
        .Reference = .Reference & " " & NameFile & ":" & LineNumber
        bFound = True
        Exit For
      End If
    End With
  Next

  'add string
  If Not bFound Then
    iBoundData = iBoundData + 1
    ReDim Preserve DataMsgId(iBoundData) As TypeMsgId
    With DataMsgId(iBoundData)
      .MsgId = StrMsg
      .Reference = NameFile & ":" & LineNumber
    End With
  End If
  Exit Sub

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.StrExists"
End Sub

Private Function GetStringFromFrx(ReferenceFrx As String) As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.GetStringFrx(" & Quote & ReferenceFrx & Quote & ")", etFullDebug

Dim szTemp As String
Dim szFileFrx As String
Dim lPos As Long
Dim iFile As Integer

  'read frx
  szTemp = Mid(ReferenceFrx, 3)
  szFileFrx = Mid(szTemp, 1, InStr(szTemp, Quote) - 1)
  lPos = "&H" & Mid(szTemp, InStr(szTemp, ":") + 1)
  
  iFile = FreeFile
  Open App.Path & "\" & szFileFrx For Binary As #iFile
  Get #iFile, lPos + 5, szTemp
  szTemp = szTemp & Input(500, iFile)
  Close #iFile
  
  lPos = InStr(szTemp, Chr$(0))
  If lPos > 0 Then szTemp = Left$(szTemp, lPos - 2)
  GetStringFromFrx = szTemp
  Exit Function

Err_Handler: If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.GetStringFrx"
End Function

'merge template file and lang file
Public Sub MergeLangFileString(ByVal Lang As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
frmMain.svr.LogEvent "Entering " & App.Title & ":basLang.MergeFileString(" & Quote & Lang & Quote & ")", etFullDebug

Dim objDataLang() As StrLang
Dim objDataLangTempl() As StrLang
Dim ii As Integer
Dim ijj As Integer

  StartMsg §§TrasLang§§("Merge language file...")
    
  'read language file
  LoadFileLang App.Path & "\" & Lang & ".lng", objDataLang
  
  'read template
  LoadFileLang App.Path & "\" & TEMPLATE_FILE_LANG, objDataLangTempl

  For ii = 1 To UBound(objDataLangTempl)
    For ijj = 1 To UBound(objDataLang)
      If objDataLang(ijj).MsgId = objDataLangTempl(ii).MsgId Then
        objDataLangTempl(ii).MsgStr = objDataLang(ijj).MsgStr
        objDataLangTempl(ii).MsgStrValid = objDataLang(ijj).MsgStrValid
        Exit For
      End If
    Next
  Next

  'copy file
  WriteTextFile App.Path & "\" & Lang & ".lng", ReadTextFile(App.Path & "\" & TEMPLATE_FILE_LANG)
 
  'rewrite file
  SaveFileLang Lang, objDataLangTempl
  EndMsg
  Exit Sub

Err_Handler:
  EndMsg
  If Err.Number <> 0 Then LogError Err.Number, Err.Description, App.Title & ":basLang.GetStringFrx"
End Sub

