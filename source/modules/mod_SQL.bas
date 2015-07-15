Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_SQL
' Level:        Framework module
' VERSION:      1.03
' Description:  Database/SQL properties, functions & subroutines
'
' Source/date:  Bonnie Campbell, 7/24/2014
' Revisions:    BLC, 7/24/2014 - 1.00 - initial version
'               BLC, 8/19/2014 - 1.01 - added versioning
'               BLC, 5/26/2015 - 1.02 - added mod_db_Templates subs/functions - GetQuerySQL, GetSQLDbTemplate
'               BLC, 6/30/2015 - 1.03 - combined GetDbQuerySQL with GetQuerySQL, renamed get... to Get... functions
' =================================

' ---------------------------------
' PROPERTY:     dbCurrent
' Description:  Gets a single instance of the current db to avoid multiple calls
'               to CurrentDb which can yield to Error 3048 "Cannot open any more databases" errors
'               due to multiple open db
' Parameters:   -
' Returns:      current database object
' Throws:       -
' References:   -
' Source/date:  Dirk Goldgar, MS Access MVP - May 22, 2013
'   http://social.msdn.microsoft.com/Forums/office/en-US/9993d229-8a00-4a59-a796-dfa2dad505bc/cannot-open-any-more-databases?forum=accessdev
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Revisions:    BLC, 7/23/2014 - initial version
' ---------------------------------
Private m_db As DAO.Database
Public Property Get dbCurrent() As DAO.Database

    If (m_db Is Nothing) Then
        Set m_db = CurrentDb
    End If

    Set dbCurrent = m_db

End Property

' ---------------------------------
'   Retrieve SQL
' ---------------------------------

' ---------------------------------
' FUNCTION:     GetSQL
' Description:  Retrieve query SQL string using query name
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:
'   Albert D. Kallal  (Access MVP) Edmonton, Alberta Canada kallal@msn.com - Sept 8, 2010
'   http://social.msdn.microsoft.com/Forums/office/en-US/3a26a941-b75b-49e4-bfe8-10c152f2b6c0/sql-or-querydef-in-vba-code?forum=accessdev
'   Daniel Pineault, CARDA Consultants Inc. - June 10, 2010
'   http://www.devhut.net/2010/06/10/ms-access-vba-edit-a-querys-sql-statement/
' Adapted:      Bonnie Campbell, July, 2014 for NCPN tools
' Revisions:    BLC, 7/23/2014 - initial version
'               BLC, 6/30/2015 - rename get... to Get...
' ---------------------------------
Public Function GetSQL(strQuery As String) As String
On Error GoTo Err_Handler:

   GetSQL = dbCurrent.QueryDefs(strQuery).sql
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSQL[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetWhereSQL
' Description:  Prepare a SQL WHERE clause based on the parameters, parameter types, fields, and
'               current WHERE clause (strWhere) passed into the function
' Assumptions:  Assumes parameters passed through params will each have the parameter name, type, and field name
'                   params(x,0) = parameter value
'                   params(x,1) = parameter type
'                   params(x,2) = database field name
'               NOTE: The function does not currently handle dependent parameters which require
'                     the presence of other parameters to be included in the WHERE clause
'                     These have to be accommodated separately.
' Parameters:   Completed SQL WHERE clause (string)
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, August, 2014 for NCPN tools
' Adapted:      Bonnie Campbell, July, 2014 for NCPN tools
' Revisions:    BLC, 8/11/2014 - initial version
'               BLC, 6/30/2015 - rename from get... to Get...
' ---------------------------------
Public Function GetWhereSQL(strWhere As String, params As Variant) As String
On Error GoTo Err_Handler:
Dim blnCheck As Boolean
Dim strParam As String
Dim i As Integer

    'default
    blnCheck = False

    For i = 0 To UBound(params) - 1
    
        'handle empty field values
        If Len(params(i, 2)) > 0 Then
    
            'handle when param isn't the only parameter (need ' AND ' in SQL WHERE clause)
            If Len(strWhere) > 0 Then strWhere = strWhere & " AND"
    
            'check if parameter is is non-empty (string) or non-zero (integer)
            Select Case params(i, 1)
                Case "string"
                    If Len(Trim(params(i, 0))) > 0 Then blnCheck = True
                    strParam = "'" & params(i, 0) & "'"
                Case "integer"
                    If params(i, 0) > 0 Then blnCheck = True
                    strParam = params(i, 0)
            End Select
        
            'prepare SQL
            If Not IsNull(params(i, 0)) And blnCheck Then
             strWhere = strWhere & " " & params(i, 2) & " = " & strParam
            End If
        
        Else
            Exit For 'done
        End If
    Next
    
   GetWhereSQL = strWhere
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetWhereSql[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetQuerySQL
' Description:  Get SQL for a query
' Assumptions:  -
' Parameters:   strQueryName - Name of query to fetch SQL for (string)
' Returns:      full SQL for the query (string)
' Throws:       none
' References:   -
' Source/date:
' S. Phinney, July 13, 2009
' http://bytes.com/topic/access/answers/871500-getting-sql-string-query
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/23/2015 - initial version
'   BLC, 5/1/2015 - moved from mod_App_Data to mod_SQL
'   ----------------- GetDbQuerySQL revisions -----------
'   BLC, 6/16/2014 - initial version
'   BLC, 5/26/2015 - moved from mod_db_Templates to mod_SQL, added error handling
'   ------------------------------------------------------
'   BLC, 6/30/2015 - combined with GetDbQuerySQL (similar functions)
' ---------------------------------
Private Function GetQuerySQL(strQueryName As String) As String
Dim qdf As DAO.QueryDef
 
    'fetch query
    Set qdf = CurrentDb.QueryDefs(strQueryName)
    
    'return SQL
    GetQuerySQL = qdf.sql
 
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetQuerySQL[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:     GetSQLTemplate
' Description:  loads SQL templates (queries as SQL string) into memory as a dictionary object
'               with query SQL strings available without querying the db tsys_SQL_templates table
' Parameters:
' Returns:      dictionary object stored in tempVars.Item("SQL")
' Assumptions:  placing
' Throws:       none
' References:   tsys_SQL_templates, Microsoft Scripting Runtime (dictionary object)
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - initial version
'               BLC, 5/26/2015 - moved from mod_db_Templates to mod_SQL, added error handling
' ---------------------------------
Public Sub GetSQLTemplates(Optional strVersion As String = "")
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, key As String, Value As String
    
    'handle default
    strSQLWhere = " WHERE Is_Supported > 0"
    
    If Len(strVersion) > 0 Then
        strSQLWhere = " AND LCase(versionID) = LCase(" & strVersion & " )"
    End If
    
    'sql
    strSQL = "SELECT * FROM tsys_Db_Templates" & strSQLWhere
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSQL)
    
    'handle no records
    If rst.EOF Then
        MsgBox "Sorry, no templates were found for this database version.", vbExclamation, _
            "Linked Database Templates Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If
    
    'prepare dictionary
    Dim dict As New Scripting.Dictionary
    Dim ary(1 To 4) As String
    Dim i As Integer
    
    'prepare the dictionary key array
    ary(1) = "context"
    ary(2) = "template_Name"
    ary(3) = "SQLstring" 'template
    ary(4) = "var_list"
    
    rst.MoveFirst
    Do Until rst.EOF
        'populate the dictionary
        For i = 1 To UBound(ary)
            key = ary(i)
            If (ary(i) = "SQLstring") Then
                Value = rst!template
            Else
                Value = rst.Fields(ary(i))
            End If
            If Not dict.Exists(key) Then
                dict.Add key, Value
            End If
        Next
        rst.MoveNext
    Loop
    
    TempVars.Add "SQL", dict

    'cleanup
    Set dict = Nothing
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSQLTemplates[mod_SQL])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
'   SQL Parameters
' ---------------------------------

' ---------------------------------
' FUNCTION:     SetParam
' Description:  Set a parameter value (useful for parameter queries)
' Assumptions:  Companion GetParam() function exists & param is publicly defined
' Parameters:   paramValue - parameter name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 24, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/24/2015  - initial version
'   BLC, 5/1/2015 - moved from mod_App_Data to mod_SQL
' ---------------------------------
Public Function SetParam(paramValue As Variant)

On Error GoTo Err_Handler
Dim param As Variant
    
    param = paramValue
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetParam[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetParam
' Description:  Get a parameter value (useful for parameter queries)
' Assumptions:  Companion GetParam() function exists & param is publicly defined
' Parameters:   paramValue - parameter name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 24, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/24/2015  - initial version
'   BLC, 5/1/2015 - moved from mod_App_Data to mod_SQL
' ---------------------------------
Public Function GetParam()

On Error GoTo Err_Handler
Dim param As Variant

    GetParam = param
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParam[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
'   SQL Functions
' ---------------------------------

' ---------------------------------
' SUB:          ConcatRelated
' Description:  Used in SQL queries to generate concatenated string of related records
' Assumptions:  used in Access SQL or control
' Parameters:   strField - field to retrieve results from & concatenate (string)
'               strTable - table or query name (string)
'               strWHERE - limiting WHERE clause (string)
'               strOrderBy - sorting ORDER BY clause (string)
'               strSeparator - character to use between concatenated values (string)
' Returns:      SQL (string, variant, or NULL if no matches)
' Notes:        1. Use square brackets around field/table names with spaces or odd characters.
'               2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
'               3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
'               4. Returning more than 255 characters to a recordset triggers this Access bug:
'                  http://allenbrowne.com/bug-16.html
' Usage:        SQL string:
'                SELECT CompanyName,  ConcatRelated("OrderDate", "tblOrders", "CompanyID = "
'                   & [CompanyID]) FROM tblCompany;
'               Access textbox control:
'                =ConcatRelated("OrderDate", "tblOrders", "CompanyID = " & [CompanyID])
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June, 2008
' http://allenbrowne.com/func-concat.html
' Adapted:      Bonnie Campbell, April 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/7/2015 - initial version
'   BLC - 5/1/2015 - integrated into Invasives Reporting tool
' ---------------------------------
Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWhere As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = ", ") As Variant
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Function:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConcatRelated[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          Coalesce
' Description:  Used in SQL queries to generate concatenated string of records
' Assumptions:  used in Access SQL or control
' Parameters:   strSQL - field to retrieve results from & concatenate (string)
'               NameList() - list of items to concatenate (string)
'               strDelim - character to use between concatenated values (string)
' Returns:      SQL (string, variant, or NULL if no matches)
' Usage:        SQL string:
'               SELECT documents.MembersOnly, Coalsce("SELECT FName From Persons WHERE Member=True",":") AS Who,
'               Coalsce("", ":", "Mary", "Joe", "Pat?") As Others FROM documents;
' Throws:       none
' References:   none
' Source/date:
' Fionuala, September 18, 2008
' http://stackoverflow.com/questions/92698/combine-rows-concatenate-rows?lq=1
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015  - initial version
'   BLC - 5/1/2015 - integrated into Invasives Reporting tool
' ---------------------------------
Function Coalsce(strSQL As String, strDelim, ParamArray NameList() As Variant)
Dim db As Database
Dim rs As DAO.Recordset
Dim strList As String

    Set db = CurrentDb

    If strSQL <> "" Then
        Set rs = db.OpenRecordset(strSQL)

        Do While Not rs.EOF
            strList = strList & strDelim & rs.Fields(0)
            rs.MoveNext
        Loop

        strList = Mid(strList, Len(strDelim))
    Else

        strList = Join(NameList, strDelim)
    End If

    Coalsce = strList

Exit_Function:
    'Clean up
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Coalesce[mod_SQL])"
    End Select
    Resume Exit_Function
End Function