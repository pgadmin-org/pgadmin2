VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the Artistic Licence

Option Explicit

'Properties
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private bTrusted As Boolean
Private szHandler As String
Private szValidator As String
Private cnDatabase As Connection
Private vTag As Variant
Private szACL As String
Private colDependent As Collection
Private colReferenced As Collection

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get FormattedID()", etFullDebug

  FormattedID = fmtID(szName)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let Oid(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Oid() As Double
Attribute Oid.VB_Description = "Returns the Language OID."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get OID()", etFullDebug

  Oid = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Language name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the object identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get ObjectType()", etFullDebug

  ObjectType = "Language"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return Trusted?
Friend Property Let Trusted(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Trusted(" & vData & ")", etFullDebug

  bTrusted = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Trusted() As Boolean
Attribute Trusted.VB_Description = "Returns a value indicating whether the Language is trusted."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Trusted()", etFullDebug

  Trusted = bTrusted
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Handler
Friend Property Let Handler(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Handler(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szHandler = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Handler() As String
Attribute Handler.VB_Description = "Returns the Language handler."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Handler()", etFullDebug

  Handler = szHandler
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Validator
Friend Property Let Validator(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let Validator(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szValidator = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Validator() As String
Attribute Validator.VB_Description = "Returns the name of a function used for function validation."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get Validator()", etFullDebug

  Validator = szValidator
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the ACL
Friend Property Let ACL(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Let ACL(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szACL = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get ACL() As String
Attribute ACL.VB_Description = "Returns the Access Control List for the object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get ACL()", etFullDebug

  ACL = szACL
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Grant Permissions
Public Sub Grant(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Grant.VB_Description = "Grant permissions on the object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Grant(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset
  
  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 14, App.Title, "Language security is only available in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "GRANT "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclUsage) = aclUsage Then szSQL = szSQL & "USAGE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON LANGUAGE " & FormattedID & " TO " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT lanacl FROM pg_language WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!lanacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

'Revoke Permissions
Public Sub Revoke(ByVal UserOrGroup As String, ByVal Privilege As aclPrivileges)
Attribute Revoke.VB_Description = "Revoke permissions from the object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Revoke(" & QUOTE & UserOrGroup & QUOTE & ", " & Privilege & ")", etFullDebug

Dim szSQL As String
Dim rs As New Recordset

  If objVersion.VersionNum < 7.3 Then
    Err.Raise vbObjectError + 14, App.Title, "Language security is only available in PostgreSQL 7.3+."
    Exit Sub
  End If
  
  'Generate and execute the Grant Query
  szSQL = "REVOKE "
  If (Privilege And aclAll) = aclAll Then
    szSQL = szSQL & "ALL "
  Else
    If (Privilege And aclUsage) = aclUsage Then szSQL = szSQL & "USAGE, "
    If Right(szSQL, 2) = ", " Then
      szSQL = Left(szSQL, Len(szSQL) - 2) & " "
    Else
      Exit Sub
    End If
  End If
  
  szSQL = szSQL & "ON LANGUAGE " & FormattedID & " FROM " & UserOrGroup
  objServer.ExecSQL szSQL, cnDatabase
  
  'Get the current ACL
  szSQL = "SELECT lanacl FROM pg_language WHERE oid = " & dOID & "::oid"
  Set rs = objServer.ExecSQL(szSQL, cnDatabase)
  If Not rs.EOF Then
    szACL = rs!lanacl & ""
  End If
  
  'Cleanup
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  
  Exit Sub
Err_Handler:
  If rs.State <> adStateClosed Then rs.Close
  Set rs = Nothing
  objServer.iLogError Err.Number, Err.Description
End Sub

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Property Get SQL()", etFullDebug
Dim szSQL As String
Dim szMember As Variant
  
  szSQL = "-- Language: " & FormattedID & vbCrLf & _
          "CREATE"
          
  If bTrusted Then szSQL = szSQL & " TRUSTED"
  szSQL = szSQL & " PROCEDURAL LANGUAGE '" & szName & "'" & vbCrLf
  If Handler = "" Then
    szSQL = szSQL & "  HANDLER " & QUOTE & QUOTE
  Else
    szSQL = szSQL & "  HANDLER " & szHandler
  End If
  If (objVersion.VersionNum >= 7.3) And (szValidator <> "") Then szSQL = szSQL & vbCrLf & "  VALIDATOR " & szValidator
  szSQL = szSQL & ";" & vbCrLf

  If szACL <> "" Then szSQL = szSQL & ParseACL(FormattedID, szACL, aclLanguage)
  
  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the dependent object is in
Public Function Dependent() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Function Dependent()", etFullDebug
  
  If colDependent Is Nothing Then Set colDependent = DepRef(dOID, cnDatabase, Me.Database, EDR_Depend)
  Set Dependent = colDependent
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function

'Return the referenced object is in
Public Function Referenced() As Collection
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgLanguage.Function Dependent()", etFullDebug
  
  If colReferenced Is Nothing Then Set colReferenced = DepRef(dOID, cnDatabase, Me.Database, EDR_Reference)
  Set Referenced = colReferenced
  
  Exit Function
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Function



