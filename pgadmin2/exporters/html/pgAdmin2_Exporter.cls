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

'Note:
' pgAdmin Exporters must
' a) Have 'pgAdmin2_Exporter' (without the quotes) somewhere in the class name
' b) Reference 'pgAdmin II Abstract Exporter' & 'Microsoft ActiveX Data Objects 2.0'

'Implement the Exporter Interface
Implements pgExporter

'The pgExporter_Author function returns info about the Exporter's Author
Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Dave Page" & vbCrLf & _
                      "http://www.pgadmin.org/"
End Property

'The pgExporter_Description function simply returns the description that is used
'in menus etc.
Private Property Get pgExporter_Description() As String
  pgExporter_Description = "Basic HTML"
End Property

'The pgExporter_Version function simply returns the Exporter Version
'This code should never need to be changed.
Private Property Get pgExporter_Version() As String
  pgExporter_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

'The pgExporter_Export subroutine is called to do the actual export. The sub is
'passed the recordset to export.
Private Sub pgExporter_Export(rsData As Variant)
Dim fNum As Integer
Dim X As Integer
Dim szTemp As String

  frmExport.Show vbModal
  If frmExport.txtFileName = "" Then
    MsgBox "No filename specified, export aborted!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Open the file
  Screen.MousePointer = vbArrowHourglass
  fNum = FreeFile
  Open frmExport.txtFileName For Output As #fNum
  'Write the HTML Headers
  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "  <TITLE>Query Results</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "  <CENTER><H3>Query Results</H3></CENTER>"
  Print #1, "  <P><B>Executed: </B>" & Now & "<BR>"
  Print #1, "  <B>Query: </B>" & rsData.Source & "</P>"
  Print #1, "  <P><CENTER><TABLE BORDER=1>"
  Print #1, "    <TR>"
  
  'Print the header row of the table
  szTemp = ""
  For X = 0 To rsData.Fields.Count - 1
    szTemp = szTemp & "<TH><B>" & rsData.Fields(X).Name & "&nbsp;</B></TH>"
  Next
  Print #1, "      " & szTemp
  Print #1, "    </TR>"
  Print #1, "    <TR>"
  
  'Print the data
  While Not rsData.EOF
    szTemp = ""
    For X = 0 To rsData.Fields.Count - 1
      If rsData.Fields(X).Type = adDBTime Then
        szTemp = szTemp & "<TD>" & Format(rsData.Fields(X).Value & "", "ttttt") & "&nbsp;</TD>"
      Else
        szTemp = szTemp & "<TD>" & rsData.Fields(X).Value & "&nbsp;</TD>"
      End If
    Next
    Print #1, "      " & szTemp
    Print #1, "    </TR>"
    rsData.MoveNext
    If Not rsData.EOF Then Print #1, "    <TR>"
  Wend
    
  'Print the footer
  Print #1, "  </TABLE></CENTER></P>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  
  'Close the file
  Close #fNum
  Screen.MousePointer = vbNormal
End Sub


