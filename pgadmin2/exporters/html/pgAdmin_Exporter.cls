VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgAdmin_Exporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Note:
' pgAdmin Exporters must
' a) Have 'pgAdmin_Exporter' (without the quotes) somewhere in the class name
' b) Reference 'pgAdmin Abstract Exporter' & 'Microsoft ActiveX Data Objects 2.5'

'Implement the Exporter Interface
Implements pgExporter

'The pgExporter_Author function returns info about the Exporter's Author
Private Property Get pgExporter_Author() As String
  pgExporter_Author = "Written by Dave Page" & vbCrLf & _
                      "http://pgadmin.postgresql.org/"
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
'passed the recordset and optionally the name of a configuration set to use.
'The configuration set should be created/saved/loaded by the plugin to allow
'unattended operation. In this demo, the only configuration data we require is
'a filename, so if a string is present we will save to that file automatically,
'if not, we will prompt the user for a filename.
Private Sub pgExporter_Export(rsData As Variant, Optional szConfig As String)
Dim fNum As Integer
Dim X As Integer
Dim szTemp As String

  'Get a filename if we don't already have one.
  If szConfig = "" Then
    frmExport.Show vbModal
    szConfig = frmExport.txtFileName
  End If
  If szConfig = "" Then
    MsgBox "No filename specified, export aborted!", vbExclamation, "Error"
    Exit Sub
  End If
  
  'Open the file
  Screen.MousePointer = vbArrowHourglass
  fNum = FreeFile
  Open szConfig For Output As #fNum
  
  'Write the HTML Headers
  Print #1, "<HTML>"
  Print #1, "<HEAD>"
  Print #1, "  <TITLE>Query Results</TITLE>"
  Print #1, "</HEAD>"
  Print #1, "<BODY>"
  Print #1, "  <CENTER><H3>Query Results</H3></CENTER>"
  Print #1, "  <CENTER><TABLE BORDER=1>"
  Print #1, "    <TR>"
  
  'Print the header row of the table
  szTemp = ""
  For X = 0 To rsData.Fields.Count - 1
    szTemp = szTemp & "<TH><B>" & rsData.Fields(X).Name & "</B></TH>"
  Next
  Print #1, "      " & szTemp
  Print #1, "    </TR>"
  Print #1, "    <TR>"
  
  'Print the data
  While Not rsData.EOF
    szTemp = ""
    For X = 0 To rsData.Fields.Count - 1
      szTemp = szTemp & "<TD>" & rsData.Fields(X).Value & "</TD>"
    Next
    Print #1, "      " & szTemp
    Print #1, "    </TR>"
    rsData.MoveNext
    If Not rsData.EOF Then Print #1, "    <TR>"
  Wend
    
  'Print the footer
  Print #1, "  </TABLE></CENTER>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
  
  'Close the file
  Close #fNum
  Screen.MousePointer = vbNormal
End Sub


