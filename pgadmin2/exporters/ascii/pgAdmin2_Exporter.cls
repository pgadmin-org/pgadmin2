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
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit
Implements pgExporter

Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Dave Page" & vbCrLf & _
                      "http://pgadmin.postgresql.org/"
End Property

Private Property Get pgExporter_Description() As String
  pgExporter_Description = "ASCII Text File"
End Property

Private Property Get pgExporter_Version() As String
  pgExporter_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Sub pgExporter_Export(rsData As Variant)
Dim fNum As Integer
Dim X As Integer
Dim Y As Integer
Dim lCount As Long
Dim szFilename As String
Dim szQuote As String
Dim szDelimiter As String
Dim bTrailing As Boolean
Dim szData As String
Dim szField As String

  frmExport.Show vbModal
  szFilename = frmExport.szFilename
  szQuote = frmExport.szQuote
  szDelimiter = frmExport.szDelimiter
  bTrailing = frmExport.bTrailing
  frmExport.Hide
  Unload frmExport
  DoEvents
  
  If szFilename = "" Then
    MsgBox "No filename specified - Export aborted!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Open the file
  Screen.MousePointer = vbArrowHourglass
  fNum = FreeFile
  Open szFilename For Output As #fNum
      
  'Loop through the records, and save line by line.
  
  lCount = 0
  Load frmStatus
  frmStatus.Show
  While Not rsData.EOF
    For X = 0 To rsData.Fields.Count - 1
      If frmExport.lvSubMap.ListItems.Count > 0 Then
        For Y = 1 To frmExport.lvSubMap.ListItems.Count
          If rsData.Fields(X).Type = adDBTime Then
            szField = Replace(Format(rsData.Fields(X).Value & "", "ttttt"), frmExport.lvSubMap.ListItems(Y).Text, frmExport.lvSubMap.ListItems(Y).SubItems(1))
          Else
            szField = Replace(rsData.Fields(X).Value & "", frmExport.lvSubMap.ListItems(Y).Text, frmExport.lvSubMap.ListItems(Y).SubItems(1))
          End If
        Next
      Else
        If rsData.Fields(X).Type = adDBTime Then
          szField = Format(rsData.Fields(X).Value & "", "ttttt")
        Else
          szField = rsData.Fields(X).Value & ""
        End If
      End If
      szData = szData & szQuote & Replace(szField, szQuote, szQuote & szQuote) & szQuote & szDelimiter
    Next
    If bTrailing = 0 Then
      szData = Mid(szData, 1, Len(szData) - Len(szDelimiter))
    End If
    Print #fNum, szData
    lCount = lCount + 1
    frmStatus.lblCount = lCount
    frmStatus.Refresh
    DoEvents
    szData = ""
    rsData.MoveNext
  Wend
  rsData.Close
   
  'Close the file
  
  Set rsData = Nothing
  Close #fNum
  Screen.MousePointer = vbNormal
  Unload frmStatus
  
  'If szConfig is blank (ie. not unattended) then display results.
  MsgBox "Exported " & lCount & " records to " & szFilename, vbInformation, "Data Export"
End Sub


