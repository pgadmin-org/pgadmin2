VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgAdmin2_Exporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

Const QUOTE = """"
Implements pgExporter

Private Property Get pgExporter_Description() As String
  pgExporter_Description = "OLE Link to MS Access"
End Property

Private Property Get pgExporter_Version() As String
  pgExporter_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Frank_Lupo" & vbCrLf & _
                      "http://www.pgadmin.org/"
End Property

Private Sub pgExporter_Export(rsData As Variant)
On Error GoTo Err_Handler
Dim X As Integer
Dim AccessApp As Object
Dim szSqlInsert As String
Dim szSqlCreate As String
Dim szSql As String
Dim szCurrentSql As String
Dim iLng As Integer
Dim iDummy As Integer
Dim szName As String
Dim szTemp As String
Dim szTableName As String
Dim bTableExists As Boolean
Dim lCount As Long
Dim lNumRecord As Long
Dim rsDataMdb
  
  frmExport.Show vbModal
  If frmExport.txtFileName = "" Then
    MsgBox "No filename specified, export aborted!", vbExclamation, "Error"
    Exit Sub
  End If
  
  If frmExport.txtTableName = "" Then
    MsgBox "Invalid table name!", vbExclamation, "Error"
    Exit Sub
  End If
  
  szTableName = frmExport.txtTableName.Text
  
  Set AccessApp = CreateObject("Access.Application")
  AccessApp.Visible = True

  'create new project
  If Dir(frmExport.txtFileName) <> "" Then
    AccessApp.OpenCurrentDatabase frmExport.txtFileName
  Else
    AccessApp.NewCurrentDatabase frmExport.txtFileName
  End If

  szCurrentSql = ""
  iDummy = 0
  szSqlInsert = "INSERT INTO " & szTableName & " ("
  szSqlCreate = "CREATE TABLE " & szTableName & " ("
  For X = 0 To rsData.Fields.Count - 1
    szName = rsData.Fields(X).Name
    
    'column name calculate
    If szName = "?column?" Then
      szName = "Dummy" & iDummy
      iDummy = iDummy + 1
    End If
    
    'escape column
    szName = "[" & szName & "]"
    
    szSqlInsert = szSqlInsert & " " & szName & " "
    szSqlCreate = szSqlCreate & " " & szName & " "
    Select Case rsData.Fields(X).Type
      Case adInteger, adBinary, adUnsignedInt
        szSqlCreate = szSqlCreate & " Integer "
      Case adBoolean
        szSqlCreate = szSqlCreate & " Boolean "
      Case adUnsignedTinyInt
        szSqlCreate = szSqlCreate & " Byte "
      Case adCurrency
        szSqlCreate = szSqlCreate & " Money "
      Case adDate, adDBDate
        szSqlCreate = szSqlCreate & " Date "
      Case adDBTime
        szSqlCreate = szSqlCreate & " Time "
      Case adDBTimeStamp
        szSqlCreate = szSqlCreate & " TimeStamp "
      Case adDouble, adNumeric
        szSqlCreate = szSqlCreate & " Double "
      Case adLongVarBinary
        szSqlCreate = szSqlCreate & " LongBinary "
      Case adLongVarWChar
        szSqlCreate = szSqlCreate & " LongText "
      Case adSingle
        szSqlCreate = szSqlCreate & " Single "
      Case adSmallInt
        szSqlCreate = szSqlCreate & " Short "
      Case adVarWChar, adLongVarChar
        szSqlCreate = szSqlCreate & " Text "
      Case adChar
        iLng = rsData.Fields(X).DefinedSize
        If iLng > 255 Then iLng = 255
        szSqlCreate = szSqlCreate & " Char(" & iLng & ") "
      Case adVarChar
        iLng = rsData.Fields(X).DefinedSize
        If iLng > 255 Then iLng = 255
        szSqlCreate = szSqlCreate & " Varchar(" & iLng & ") "
      Case Else
        Err.Raise -1, , "Type field not found Postgresql->Access !! " & rsData.Fields(X).Type
        Exit Sub
    End Select
    If X < rsData.Fields.Count - 1 Then
      szSqlCreate = szSqlCreate & " ,"
      szSqlInsert = szSqlInsert & " ,"
    End If
  Next
  szSqlCreate = szSqlCreate & " )"
  szSqlInsert = szSqlInsert & " ) VALUES ("

  'verify if table exists
  On Error Resume Next
  bTableExists = AccessApp.CurrentDb.TableDefs(szTableName).Name <> ""
  On Error GoTo Err_Handler
  
  If bTableExists Then
    Select Case frmExport.cboCond.ItemData(frmExport.cboCond.ListIndex)
      Case 1
        'drop table
        AccessApp.CurrentDb.TableDefs.Delete (szTableName)
        bTableExists = False
      Case 2
        'empty table
        Set rsDataMdb = AccessApp.CurrentDb.TableDefs(szTableName).OpenRecordset
        While Not rsDataMdb.EOF
          rsDataMdb.Delete
          rsDataMdb.MoveNext
        Wend
    End Select
  End If
  
  'Create table
  If Not bTableExists Then
    szCurrentSql = szSqlCreate
    AccessApp.CurrentDb.Execute szSqlCreate
  End If

  'status insert
  frmStatus.pb.Value = 1
  frmStatus.pb.Min = 1
  
  'calcuate number row
  lNumRecord = 0
  While Not rsData.EOF
    lNumRecord = lNumRecord + 1
    rsData.MoveNext
  Wend
  If lNumRecord = 1 Then lNumRecord = lNumRecord + 1
  frmStatus.pb.Max = lNumRecord
  frmStatus.Show
  
  rsData.MoveFirst
  lCount = 1
  'Enter Data
  While Not rsData.EOF
    frmStatus.pb.Value = lCount
    frmStatus.lblCount = lCount & " of " & lNumRecord
    DoEvents
    szCurrentSql = ""
    szSql = szSqlInsert
    For X = 0 To rsData.Fields.Count - 1
      If VarType(rsData.Fields(X).Value) = vbNull Then
          szSql = szSql & "NULL"
      Else
        Select Case rsData.Fields(X).Type
          Case adInteger, adBinary, adUnsignedInt, adBoolean, adUnsignedTinyInt, _
               adDouble, adNumeric, adLongVarBinary, adSingle, adSmallInt
            szSql = szSql & QUOTE & rsData.Fields(X).Value & QUOTE
          Case adCurrency
            szSql = szSql & QUOTE & rsData.Fields(X).Value & QUOTE
          Case adDate, adDBDate, adDBTime, adDBTimeStamp
            szSql = szSql & QUOTE & rsData.Fields(X).Value & QUOTE
          Case adVarWChar, adLongVarChar, adChar, adVarChar, adLongVarWChar
            szSql = szSql & QUOTE & Replace(rsData.Fields(X).Value, QUOTE, QUOTE & QUOTE) & QUOTE
          Case Else
            szSql = szSql & QUOTE & Replace(rsData.Fields(X).Value, QUOTE, QUOTE & QUOTE) & QUOTE
        End Select
      End If
      If X < rsData.Fields.Count - 1 Then szSql = szSql & " ,"
    Next
    szSql = szSql & " ) "
    szCurrentSql = szSql
    AccessApp.CurrentDb.Execute szSql
    rsData.MoveNext
    lCount = lCount + 1
  Wend
  
  szCurrentSql = ""
  Screen.MousePointer = vbNormal
  Unload frmStatus
  
  Exit Sub
Err_Handler:
  Screen.MousePointer = vbNormal
  Unload frmStatus
  
  If Err.Number <> 0 Then
    szTemp = "An error has occured in pgHba:pgExporter_Export:" & vbCrLf & vbCrLf & _
             "Number: " & Err.Number & vbCrLf & "Description: " & Err.Description
    If Len(szCurrentSql) > 0 Then szTemp = szTemp & vbCrLf & vbCrLf & "SQL:" & vbCrLf & vbCrLf & szCurrentSql
    MsgBox szTemp, vbCritical + vbApplicationModal, App.Title & ":pgHba"
  End If
End Sub
