Attribute VB_Name = "basGlobal"
' pgAdmin II - PostgreSQL Tools
' Copyright (C) 2001, 2002, The pgAdmin Development Team
' This software is released under the pgAdmin Public Licence
'
' basGlobal.bas - Contains global declarations and constants.

Option Explicit

'Makes life easier...
Global Const QUOTE = """"

'Global Context object. This contains Globals.
Global ctx As New clsContext

'Global Exporters Class
Global exp As New clsExporters

'Global Plugins Class
Global plg As New clsPlugins

'Msg Timer start value.
Global sTimer As Single

'Default HighLight Colours
Global Const DEFAULT_AUTOHIGHLIGHT = "ALTER|0|0|16711680;COMMENT|0|0|16711680;CREATE|0|0|16711680;DELETE|0|0|16711680;DROP|0|0|16711680;EXPLAIN|0|0|16711680;GRANT|0|0|16711680;INSERT|0|0|16711680;REVOKE|0|0|16711680;" & _
                                     "SELECT|0|0|16711680;UPDATE|0|0|16711680;VACUUM|0|0|16711680;AGGREGATE|0|0|255;CONSTRAINT|0|0|255;DATABASE|0|0|255;FUNCTION|0|0|255;GROUP|0|0|255;INDEX|0|0|255;" & _
                                     "LANGUAGE|0|0|255;OPERATOR|0|0|255;RULE|0|0|255;SEQUENCE|0|0|255;TABLE|0|0|255;TRIGGER|0|0|255;ABORT|0|0|11998061;BEGIN|0|0|11998061;" & _
                                     "CHECKPOINT|0|0|11998061;CLOSE|0|0|11998061;CLUSTER|0|0|11998061;COMMIT|0|0|11998061;COPY|0|0|11998061;DECLARE|0|0|11998061;FETCH|0|0|11998061;LISTEN|0|0|11998061;" & _
                                     "LOAD|0|0|11998061;LOCK|0|0|11998061;MOVE|0|0|11998061;NOTIFY|0|0|11998061;REINDEX|0|0|11998061;RESET|0|0|11998061;ROLLBACK|0|0|11998061;SET|0|0|11998061;SHOW|0|0|11998061;TRUNCATE|0|0|11998061;" & _
                                     "UNLISTEN|0|0|11998061;AND|0|0|32768;AS|0|0|32768;ASC|0|0|32768;ASCENDING|0|0|32768;BY|0|0|32768;CASE|0|0|32768;DESC|0|0|32768;DESCENDING|0|0|32768;ELSE|0|0|32768;FROM|0|0|32768;END|0|0|32768;HAVING|0|0|32768;INTO|0|0|32768;" & _
                                     "ON|0|0|32768;OR|0|0|32768;ORDER|0|0|32768;THEN|0|0|32768;USING|0|0|32768;WHEN|0|0|32768;WHERE|0|0|32768;"
                                     


