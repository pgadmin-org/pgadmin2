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

Private Property Get pgExporter_Description() As String
  pgExporter_Description = "OLE Link to MS Excel"
End Property

Private Property Get pgExporter_Version() As String
  pgExporter_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Dave Page" & vbCrLf & _
                      "http://pgadmin.postgresql.org/"
End Property

Private Sub pgExporter_Export(rsData As Variant)
On Error Resume Next
Dim X As Integer
Dim Y As Integer
  Dim ExcelApp As Object
  Set ExcelApp = CreateObject("excel.application")
  ExcelApp.Visible = True
  ExcelApp.Workbooks.Add
  
  'Create Header
  For X = 0 To rsData.Fields.Count - 1
    ExcelApp.Cells(1, X + 1).NumberFormat = "@"
    ExcelApp.Cells(1, X + 1).Font.Bold = True
    ExcelApp.Cells(1, X + 1).Value = rsData.Fields(X).Name
  Next

  'Enter Data

  Y = 2
  While Not rsData.EOF
    For X = 0 To rsData.Fields.Count - 1
      ExcelApp.Cells(Y, X + 1).NumberFormat = "@"
      ExcelApp.Cells(Y, X + 1).Value = rsData.Fields(X).Value
    Next
    rsData.MoveNext
    Y = Y + 1
  Wend
  Screen.MousePointer = vbNormal

  'Autofit
  
  For X = 0 To rsData.Fields.Count - 1
    ExcelApp.Columns(GetCol(X + 1)).AutoFit
  Next
  
End Sub

Private Function GetCol(ColNum As Integer) As String
Dim Char1 As String
Dim Char2 As String
Dim X As Double
  If ColNum < 1 Or ColNum > 256 Then Exit Function
  X = ColNum / 26
  If X < 1 Then
    Char1 = ""
  Else
    Char1 = Chr(96 + Int(X))
  End If
  Char2 = Chr(96 + (ColNum - (Int(X) * 26)))
  GetCol = Char1 & Char2
End Function


