VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' pgSchema - PostgreSQL Schema Objects
' Copyright (C) 2001 - 2003, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence

Option Explicit

'Properties
Private dOID As Double
Private szName As String
Private szIdentifier As String
Private bSystemObject As Boolean
Private szOwner As String
Private szInputFunction As String
Private szOutputFunction As String
Private lInternalLength As Long
Private szDefault As String
Private szElement As String
Private szDelimiter As String
Private bPassedByValue As Boolean
Private szAlignment As String
Private szStorage As String
Private cnDatabase As Connection
Private dNamespaceOID As Double
Private szNamespace As String
Private vTag As Variant

Public Property Get FormattedID() As String
Attribute FormattedID.VB_Description = "Returns a formatted identifier for use in SQL statements."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get FormattedID()", etFullDebug
 
  If objVersion.VersionNum >= 7.3 Then
    FormattedID = fmtID(szNamespace) & "." & fmtID(szName)
  Else
    FormattedID = fmtID(szName)
  End If
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Tag
Public Property Let Tag(ByVal vData As Variant)
Attribute Tag.VB_Description = "Set/Returns a variant for any use."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Tag(" & vData & ")", etFullDebug

  vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Set Tag(ByVal vData As Variant)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Set Tag(" & vData & ")", etFullDebug

  Set vTag = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Tag() As Variant
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Tag()", etFullDebug

  If IsObject(vTag) Then
    Set Tag = vTag
  Else
    Tag = vTag
  End If
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace OID
Friend Property Let NamespaceOID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let NamespaceOID(" & vData & ")", etFullDebug

  dNamespaceOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get NamespaceOID() As Double
Attribute NamespaceOID.VB_Description = "Returns the OID of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get NamespaceOID()", etFullDebug

  NamespaceOID = dNamespaceOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Namespace Name
Friend Property Let Namespace(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Namespace(" & vData & ")", etFullDebug

  szNamespace = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Namespace() As String
Attribute Namespace.VB_Description = "Returns the name of the parent Namespace."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Namespace()", etFullDebug

  Namespace = szNamespace
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set the Connection to use
Friend Property Set Connection(vData As Connection)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Set Connection(Connection: " & vData & ")", etFullDebug

  Set cnDatabase = vData
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Get the Database Name
Public Property Get Database() As String
Attribute Database.VB_Description = "Returns the parent database name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Database()", etFullDebug

Static szDB As String

  If szDB = "" Then szDB = GetDatabase(cnDatabase.ConnectionString)
  Database = szDB
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OID
Friend Property Let OID(ByVal vData As Double)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let OID(" & vData & ")", etFullDebug

  dOID = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OID() As Double
Attribute OID.VB_Description = "Returns the Type OID."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get OID()", etFullDebug

  OID = dOID
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Name
Friend Property Let Name(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Name(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szName = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Name() As String
Attribute Name.VB_Description = "Returns the Type name."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Name()", etFullDebug

  Name = szName
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Identifier
Friend Property Let Identifier(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Identifier(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szIdentifier = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Identifier() As String
Attribute Identifier.VB_Description = "Returns the onject identifier."
Attribute Identifier.VB_MemberFlags = "200"
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Identifier()", etFullDebug

  Identifier = szIdentifier
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Return the Object Type
Public Property Get ObjectType() As String
Attribute ObjectType.VB_Description = "Returns the object type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get ObjectType()", etFullDebug

  ObjectType = "Type"
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return a boolean indicating whether this is a System Object
Friend Property Let SystemObject(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let SystemObject(" & vData & ")", etFullDebug

  bSystemObject = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get SystemObject() As Boolean
Attribute SystemObject.VB_Description = "Indicates whether this is a system object."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get SystemObject()", etFullDebug

  SystemObject = bSystemObject
  
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Owner
Friend Property Let Owner(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Let Owner(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOwner = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Owner() As String
Attribute Owner.VB_Description = "Returns the Type's owner."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgDatabase.Property Get Owner()", etFullDebug

  Owner = szOwner
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the InputFunction
Friend Property Let InputFunction(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let InputFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szInputFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get InputFunction() As String
Attribute InputFunction.VB_Description = "Returns the Input Function name used by the Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get InputFunction()", etFullDebug

  InputFunction = szInputFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the OutputFunction
Friend Property Let OutputFunction(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let OutputFunction(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szOutputFunction = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get OutputFunction() As String
Attribute OutputFunction.VB_Description = "Returns the Output Function name used by the Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get OutputFunction()", etFullDebug

  OutputFunction = szOutputFunction
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the InternalLength
Friend Property Let InternalLength(ByVal vData As Long)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let InternalLength(" & vData & ")", etFullDebug

  lInternalLength = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get InternalLength() As Long
Attribute InternalLength.VB_Description = "Returns the Internal Length of the data Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get InternalLength()", etFullDebug

  InternalLength = lInternalLength
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Default
Friend Property Let Default(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Default(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDefault = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Default() As String
Attribute Default.VB_Description = "Returns the Default value for the Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Default()", etFullDebug

  Default = szDefault
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Element
Friend Property Let Element(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Element(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szElement = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Element() As String
Attribute Element.VB_Description = "Returns the Element Type if this is an array Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Element()", etFullDebug

  Element = szElement
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Delimiter
Friend Property Let Delimiter(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Delimiter(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szDelimiter = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Delimiter() As String
Attribute Delimiter.VB_Description = "Returns the delimiter used for arrays of this Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Delimiter()", etFullDebug

  Delimiter = szDelimiter
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return PassedByValue?
Friend Property Let PassedByValue(ByVal vData As Boolean)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let PassedByValue(" & vData & ")", etFullDebug

  bPassedByValue = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get PassedByValue() As Boolean
Attribute PassedByValue.VB_Description = "Returns a value indicating whether the Type is Passed By Value (or By Reference if False)."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get PassedByValue()", etFullDebug

  PassedByValue = bPassedByValue
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Alignment
Friend Property Let Alignment(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Alignment(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szAlignment = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Alignment() As String
Attribute Alignment.VB_Description = "Returns trhe Alignment of the Type."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Alignment()", etFullDebug

  Alignment = szAlignment
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Storage
Friend Property Let Storage(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Storage(" & QUOTE & vData & QUOTE & ")", etFullDebug

  szStorage = vData
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Storage() As String
Attribute Storage.VB_Description = "Returns the Type storage method (PostgreSQL 7.1+)."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Storage()", etFullDebug

  Storage = szStorage
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

'Set/Return the Comment
Public Property Let Comment(ByVal vData As String)
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Let Comment(" & vData & ")", etFullDebug

Dim szSQL As String

  szSQL = "COMMENT ON TYPE " & FormattedID & " IS '" & dbSZ(vData) & "'"
  objServer.ExecSQL szSQL, cnDatabase
  objServer.Databases(Database).KillCommentCache
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
Public Property Get Comment() As String
Attribute Comment.VB_Description = "Sets/Returns the Type comment."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get Comment()", etFullDebug

  Comment = objServer.Databases(Database).LookupComment(dOID)
    
  Exit Property
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property

Public Property Get SQL() As String
Attribute SQL.VB_Description = "Returns the object's reverse engineered SQL."
If inIDE Then: On Error GoTo 0: Else: On Error GoTo Err_Handler
objServer.iLogEvent "Entering " & App.Title & ":pgType.Property Get SQL()", etFullDebug
Dim szSQL As String
  
  If objVersion.VersionNum >= 7.3 Then
    szSQL = "-- Type: " & fmtID(szNamespace) & "." & fmtID(szName) & vbCrLf & _
          "CREATE TYPE " & FormattedID & "(" & _
          "INPUT = " & szInputFunction & ", " & _
          "OUTPUT = " & szOutputFunction
  Else
    szSQL = "-- Type: " & fmtID(szName) & vbCrLf & _
          "CREATE TYPE " & FormattedID & "(" & _
          "INPUT = " & szInputFunction & ", " & _
          "OUTPUT = " & szOutputFunction
  End If
  
  If lInternalLength = 0 Then
    szSQL = szSQL & ", INTERNALLENGTH = VARIABLE"
  Else
    szSQL = szSQL & ", INTERNALLENGTH = " & lInternalLength
  End If
  
  If szDefault <> "" Then szSQL = szSQL & ", DEFAULT = " & szDefault
  If szElement <> "" Then szSQL = szSQL & ", ELEMENT = " & szElement
  If szDelimiter <> "" Then szSQL = szSQL & ", DELIMITER = " & fmtID(szDelimiter)
  If bPassedByValue Then szSQL = szSQL & ", PASSEDBYVALUE"
  If szAlignment <> "" Then szSQL = szSQL & ", ALIGNMENT = " & szAlignment
  If szStorage <> "" Then szSQL = szSQL & ", STORAGE = " & szStorage
  
  szSQL = szSQL & ");" & vbCrLf
          
  If Comment <> "" Then szSQL = szSQL & "COMMENT ON TYPE " & FormattedID & " IS '" & dbSZ(Comment) & "';" & vbCrLf

  SQL = szSQL
  
  Exit Sub
Err_Handler:  objServer.iLogError Err.Number, Err.Description
End Property
