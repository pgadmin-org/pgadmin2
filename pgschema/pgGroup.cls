VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Collections
Private colMembers As Members

'Properties
Private szName As String
Private szIdentifier As String
Private lID As Long
Private szMemberList As String
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get FormattedID()", etFullDebug

  FormattedID = fmtID(szName)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Member Users Collection
Public Property Get Members() As Members
Attribute Members.VB_Description = "Returns a collection of Users in the Group."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get Members()", etFullDebug

Dim szConnect As String

  If colMembers Is Nothing Then
    Set colMembers = New Members
    colMembers.Name = szName
    colMembers.Initialize szMemberList
  End If
  Set Members = colMembers
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Group name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the group identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get ObjectType()", etFullDebug

  ObjectType = "Group"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return ID
Friend Property Let ID(ByVal vData As Long)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let ID(" & vData & ")", etFullDebug

  lID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ID() As Long
Attribute ID.VB_Description = "Returns the Group ID."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get ID()", etFullDebug

  ID = lID
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Members (Internal only)
Friend Property Let MemberList(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Let MemberList(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szMemberList = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Friend Property Get MemberList() As String
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgGroup.Property Get MemberList()", etFullDebug

  MemberList = szMemberList
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the objects reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":Property Get pgGroup.SQL()", etFullDebug
Dim szSQL As String
Dim szMember As Variant
Dim szTemp As String
  
  szSQL = "-- Group: " & FormattedID & vbCrLf & _
          "CREATE GROUP " & FormattedID & " WITH SYSID " & lID
  
  If Members.Count > 0 Then szSQL = szSQL & " USER "
  For Each szMember In Members
    'For some reason dbSZing (and probably fmtID as well) szMember gives
    'a ByRef mismatch, so copy to szTemp first
    szTemp = szMember
    szSQL = szSQL & fmtID(szTemp) & ", "
  Next szMember
  szSQL = Left(szSQL, Len(szSQL) - 2) & ";" & vbCrLf

  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
