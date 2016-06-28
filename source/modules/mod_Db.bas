Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Db
' Level:        Framework module
' Version:      1.04
' Description:  Database related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added mod_db_Templates subs/functions - qryExists
'               BLC, 5/26/2016 - 1.02 - added VirtualDAORecordset()
'               BLC, 6/6/2016  - 1.03 - added error handling for duplicate templates, renamed global to g_AppTemplates
'                                       also added SQL sanitization (escape/replace special chars)
'               BLC, 6/9/2016  - 1.04 - added CreateTempRecords()
' =================================

' ---------------------------------
' Declarations
' ---------------------------------
'   AppTemplates global dictionary --> defined in std template [mod_Db]
Public g_AppTemplates As Scripting.Dictionary
Public Const PARAM_SEPARATOR As String = ">>"

' ---------------------------------
' Types & Type Descriptions
' ---------------------------------
' -32768  Form                    1   Table - Local Access Tables
' -32766  Macro                   2   Access Object - Database
' -32764  Reports                 3   Access Object - Containers
' -32761  Module                  4   Table - Linked ODBC Tables
' -32758  Users                   5   Queries
' -32757  Database Document       6   Table - Linked Access Tables
' -32756  Data Access Pages       8   SubDataSheets
' ---------------------------------

' ---------------------------------
'  Database & Recordset Actions
' ---------------------------------

' =================================
' FUNCTION:     BEUpdates
' Description:  Runs SQL statement updates from the systems table tsys_BE_Updates. Such
'               updates are sometimes necessary when there is a remote copy of the back-end
'               file that the developer cannot access, but which needs to be updated to
'               include the current release information. tsys_BE_Updates has the following
'               structure:  Update_ID (txt serial number autoincrementing), Is_done (yes/no),
'               Run_date (datetime), SQL_statement (memo), Update_desc (txt 100)
' Parameters:   bRunAll - True (default), or False if only running lines where [Is_done]=False
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/10/2008
' Revisions:    JRB, 11/21/2008 - added optional parameter to either run all update lines
'                   (default), or just one where [Is_done]=False
'               BLC, 4/30/2015  - moved to mod_Db framework module from mod_Custom_Functions
'                                 added check for BOF & EOF to avoid Error #3021 no current record on rs.MoveLast when no records exist
'               BLC, 5/18/2015 - renamed & removed fxn prefix
'               BLC, 6/5/2016  - adapted for Big Rivers App naming revisions (removed field underscores)
' =================================
Public Function BEUpdates(Optional ByVal bRunAll As Boolean = True)
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim intNumUpdates As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strSQL As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT tsys_BE_Updates.* FROM tsys_BE_Updates " & _
        "ORDER BY tsys_BE_Updates.ID;", dbOpenDynaset)

    ' Check for BOF & EOF to avoid Error # 3021 No current record
    If Not rs.BOF And rs.EOF Then

        ' Counts the number of db update records in the system table
        rs.MoveLast    ' Need to do this to make the record count accurate
        intNumUpdates = rs.RecordCount
        If intNumUpdates = 0 Then    ' No records in the recordset
            GoTo Exit_Procedure
        End If
    
        ' First pass to verify the tables in the specified database
        '   Initialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Performing database updates", intNumUpdates)
        intI = 0
        rs.MoveFirst
        On Error Resume Next
        Do Until rs.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            If bRunAll = True Or rs![IsDone] = False Then
                DoCmd.SetWarnings False
                strSQL = rs![SQLStatement]
                DoCmd.RunSQL strSQL
                With rs
                    .Edit
                    ![RunDate] = Now()
                    ![IsDone] = True
                    .Update
                End With
            End If
            rs.MoveNext
        Loop
        
    End If

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    varReturn = SysCmd(acSysCmdRemoveMeter)
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3061   ' Bad parameters for the SQL string
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - BEUpdates[mod_Db])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - BEUpdates[mod_Db])"
    End Select
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     MergeRecordsets
' Description:  Merge two recordsets into one (useful when the recordsets already exist vs. direct SQL union)
' Assumptions:  Recordsets have the same fields in the same order
' Parameters:   rsA - DAO recordset A
'               rsB - DAO recordset B to merge with A
' Returns:      DAO.Recordset
' Throws:       none
' References:   none
' Source/date:
' Chris Oswald, January 26, 2011
' http://www.mrexcel.com/forum/excel-questions/524214-visual-basic-applications-joining-multiple-recordets-multiple-databases.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/13/2015 - moved from mod_App_Data to mod_Db
' ---------------------------------
Public Function MergeRecordsets(rsA As DAO.Recordset, rsB As DAO.Recordset) As DAO.Recordset

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rsOut As DAO.Recordset
    Dim iCount As Integer
    
    'handle empty recordsets
    If rsA Is Nothing Then
        'check rsB
        If rsB Is Nothing Then
            GoTo Exit_Handler
        Else
            Set MergeRecordsets = rsB
            GoTo Exit_Handler
        End If
    End If
    

'With rsA
    'check if rsA and rsB are both populated --> if not, exit
    If (rsA.EOF And rsA.BOF) Then
        'rsA not populated
        If (rsB.EOF And rsB.BOF) Then
            'neither is populated --> EXIT!
            GoTo Exit_Handler
        Else
            'rsB populated --> return rsB
            Set MergeRecordsets = rsB
        End If
    Else
        'rsA populated --> if rsB not populated, return rsA
        If (rsB.EOF And rsB.BOF) Then
            Set MergeRecordsets = rsA
            GoTo Exit_Handler
        End If
    'End If
    
    'create output recordset vs. just adding to rsB
    Set rsOut = rsA
    Do Until rsB.EOF
        'add rsB values as new rsOut records
        rsOut.AddNew
        For iCount = 0 To rsB.Fields.Count - 1
            rsOut.Fields(iCount).Value = rsB.Fields(iCount).Value
        Next
        rsOut.Update
        rsB.MoveNext
    Loop
    
    'rsOut.Edit
    
    'iterate through recordset
    'rsA.MoveFirst
    'Do Until rsA.EOF
        'add rsA values as new rsOut records
     '   rsOut.AddNew
     '   For iCount = 0 To rsA.Fields.count - 1
     '       rsOut.Fields(iCount).Value = rsA.Fields(iCount).Value
     '   Next
     '   rsOut.Update
     '   rsA.MoveNext
    'Loop
'End With
End If
    Set MergeRecordsets = rsOut

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MergeRecordsets[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ClearTable
' Description:  Deletes records from table
' Assumptions:  Table is in the current database (not linked)
' Parameters:   strTable - table name (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015  - initial version
' ---------------------------------
Public Sub ClearTable(strTable As String)

On Error GoTo Err_Handler
    
    Dim strSQL As String
    
    'clear table
    strSQL = "DELETE * FROM " & strTable & ";"
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearTable[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Validate Database Objects
' ---------------------------------

' =================================
' FUNCTION:     TableExists
' Description:  Returns whether the specified table exists in the current database collection
' Parameters:   strTableName - string for the name of the table to check
' Returns:      True if the specified table exists in the master systems table, or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/29/2009
' Revisions:    JRB, 6/29/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function TableExists(ByVal strTableName As String) As Boolean
    On Error GoTo Err_Handler

    TableExists = DCount("*", "MSysObjects", "(([Type] In (1,4,6)) AND ([Name]=""" & _
        strTableName & """))")

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TableExists[mod_Db])"
    End Select
    Resume Exit_Handler

End Function

' ---------------------------------
' FUNCTION:     QueryExists
' Description:  Determine if a query exists in a database
' Parameters:   strQueryName - query name(string)
' Returns:      true - if found (boolean); false - if not found
' Throws:       -
' References:   -
' Source/date:  SOS, 3/20/2010
'               http://www.access-programmers.co.uk/forums/showthread.php?t=190747
' Adapted:      Bonnie Campbell, May 1, 2015
' Revisions:    BLC, 5/1/2015 - initial version
' ---------------------------------
Function QueryExists(strQueryName As String) As Boolean
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.QueryDef
    
    On Error GoTo Err_Handler
    Set db = CurrentDb
    Set tdf = db.QueryDefs(strQueryName)
    
    QueryExists = True

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3265
        QueryExists = False
        Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - QueryExists[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          qryExists
' Description:  Checks if query exists in database as a permanent query(QueryDefs)
' Parameters:   strQueryName - query name as a string
' Returns:      true - if found (boolean); false - if not found
' Throws:       -
' References:   -
' Source/date:  Nick Vans, January 31, 2008
'               http://bytes.com/topic/access/answers/765384-determine-if-query-x-exists
' Adapted:      Bonnie Campbell, June 17, 2014
' Revisions:    6/17/2014 - BLC - initial version
' ---------------------------------
Public Function qryExists(strQueryName As String) As Boolean

    Dim qdf As DAO.QueryDef
    
    'default
    qryExists = False
  
    For Each qdf In CurrentDb.QueryDefs
'        Debug.Print qdf.Name
        If qdf.Name = strQueryName Then
            qryExists = True
            Exit For
        End If
    Next

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - qryExists[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Get Database Application Info
' ---------------------------------

' ---------------------------------
' FUNCTION:     getAccessObjectType
' Description:  looks up object type in Access sys tables
' Parameters:   strName  - name of object w/in Access
' Returns:      long (type) or NULL if object doesn't exist
'                   ----------------
'                   1 = Access Table
'                   4 = OBDB-Linked Table / View
'                   5 = Access Query
'                   6 = Attached (Linked) File  (such as Excel, another Access Table or query, text file, etc.)
'                   -32768 = Access Form
'                   -32764 = Access Report
'                   -32761 = Access Module
'                   ----------------
' Throws:       none
' References:   Tom Davidson, April 8, 2011
'   http://stackoverflow.com/questions/2090578/ms-access-determine-object-type

' Source/date:  Bonnie Campbell August 20, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 8/20/2014 - initial vesrion
'               BLC, 4/30/2015 - moved from mod_Common_UI
' ---------------------------------
Public Function getAccessObjectType(strObject As String)
On Error GoTo Err_Handler:

    getAccessObjectType = DLookup("Type", "MSysObjects", "NAME = '" & strObject & "'")
   
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getAccessObjectType[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetTempVarIndex
' Description:  Retrieves the index of a TempVar item
' Parameters:   strItem - item name(string)
' Returns:      index of item, if found (integer); not found returns -1
' Throws:       -
' References:   -
' Source/date:  Dal Jeanis, 7/11/2013
'               http://www.accessforums.net/modules/demo-module-vba-code-syntax-using-tempvars-36353.html
' Adapted:      Bonnie Campbell, Sep 1, 2014
' Revisions:    BLC, 9/1/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_Db
' ---------------------------------
Public Function GetTempVarIndex(strItem) As String
On Error GoTo Err_Handler

Dim i As Integer

    For i = 0 To [TempVars].Count - 1
        If [TempVars].item(i).Name = strItem Then
            'fetch the index and exit
            GetTempVarIndex = i
            Exit Function
        End If
    Next i
    
    'none found -> return -1
    GetTempVarIndex = -1
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTempVarIndex[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     HasRecords
' Description:  Returns whether the specified table has records or not
' Parameters:   strName - string for the name of the table or query to check
' Returns:      True if the specified table/query has records, or False if not
' Throws:       none
' References:   Fionnuala, Oct 22, 2010
'               http://stackoverflow.com/questions/3994956/meaning-of-msysobjects-values-32758-32757-and-3-microsoft-access
' Source/date:  Bonnie Campbell, May 26, 2015
' Revisions:    BLC, 5/26/2015 - initial version
' =================================
Public Function HasRecords(ByVal strName As String) As Boolean
    On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Dim blnHasRecords As Boolean
    
    blnHasRecords = False
    
    ' check for table/query - 1(table), 4(Linked ODBC), 6(Linked Access), 5(query)
    If DCount("*", "MSysObjects", "(([Type] In (1,4,6,5)) AND ([Name]=""" & _
        strName & """))") > 0 Then
            Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strName & ";")
            
            'check if empty (BOF & EOF = true)
            If Not (rs.BOF And rs.EOF) Then
                blnHasRecords = True
            End If
    End If

    HasRecords = blnHasRecords

Exit_Handler:
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HasRecords[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:     GetTemplates
' Description:  loads templates into memory as a global dictionary object (dictTemplates)
'               makes current templates available without querying the db tsys_SQL_templates table
' Parameters:   strSyntax - specifies syntax of the template to retrieve (T-SQL, JET, etc.)
'               strParams - specifies the parameters & their datatypes for the template
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   tsys_Db_templates, Microsoft Scripting Runtime (dictionary object)
'   HansUp, June 27, 2013
'   http://stackoverflow.com/questions/17328092/how-to-display-access-query-results-without-having-to-create-temporary-query
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - initial version
'               BLC, 5/13/2016 - shifted from mod_Db_Templates to mod_Db & adjusted to match tsys_Db_Templates
'               BLC, 5/19/2016 - revised documentation & renamed GetTemplates() vs. GetSQLTemplates() since tsys_Db_Templates
'                                can accommodate more than SQL
'               BLC, 6/5/2016  - revised to set strSyntax to "T-SQL" to avoid error due to multiple items of same name in dict
'               BLC, 6/6/2016  - added error handling for duplicate templates, renamed global to g_AppTemplates
' ---------------------------------
Public Sub GetTemplates(Optional strSyntax As String = "", Optional params As String = "")

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, key As String
    Dim Value As Variant
    
    'handle default
    strSQLWhere = " WHERE IsSupported > 0"
    
    If Len(strSyntax) = 0 Then
        strSyntax = "T-SQL"
    End If
    
    strSQLWhere = strSQLWhere & " AND LCase(Syntax) = LCase('" & strSyntax & "')"
    
    'sql -> ID, Version, IsSupported, Context, Syntax, TemplateName, Params, Template, Remarks,
    '       EffectiveDate, RetireDate, CreateDate, CreatedBy_ID, LastModified, LastModifiedBy_ID
    strSQL = "SELECT * FROM tsys_Db_Templates" & strSQLWhere & ";"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'handle no records
    If rs.EOF Then
        MsgBox "Sorry, no templates were found for this database version.", vbExclamation, _
            "Linked Database Templates Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If
    
    'prepare dictionary
    Dim dict As New Scripting.Dictionary, dictParam As New Scripting.Dictionary
    Dim ary(1 To 5) As String, ary2() As String, param() As String
    Dim i As Integer, j As Integer
    
    'prepare the dictionary key array
    ary(1) = "Context"
    ary(2) = "TemplateName"
    ary(3) = "Template" 'template
    ary(4) = "Params"
    ary(5) = "Syntax"
    
    'prepare array of dictionaries
    Dim dictTemplates As Dictionary
    Set dictTemplates = New Scripting.Dictionary
    
    rs.MoveFirst
    Do Until rs.EOF
        'create new dictionary object
        Set dict = New Scripting.Dictionary
        
        'populate the dictionary
        For i = 1 To UBound(ary)
            
            key = ary(i)
            
            If key = "Params" Then
                'create new dictionary for param name & data type
                Set dictParam = New Scripting.Dictionary

'Debug.Print rs.Fields(ary(i))

                'separate parameters
                ary2 = Split(Nz(rs.Fields(ary(i)), ":"), "|")
                
                'prepare sets of param name & data type --> split(ary2(i), ":") yields name & data type
                For j = 0 To UBound(ary2)
                
                    'split the param into name & data type
                    param = Split(ary2(j), ":")
                    
                    If Not dictParam.Exists(param(0)) And Len(param(0)) <> 0 Then
                        dictParam.Add param(0), param(1)
                    End If
                
                Next
                
                Set Value = dictParam

            Else
                Value = Nz(rs.Fields(ary(i)), "")
            End If
            
            'add key if it isn't already there
            If Not dict.Exists(key) Then
                If IsNull(Value) Then MsgBox key, vbOKCancel, "is NULL"
                'Debug.Print Nz(Value, key & "-NULL")
                dict.Add key, Value
            End If
            
        Next
        
'        Debug.Print dict("TemplateName")
        
'        If dictTemplates.Exists("TemplateName") Then
'            Debug.Print "dict: " & dict("TemplateName")
'        End If
        
        'add template dictionary to dictionary of templates
        dictTemplates.Add dict("TemplateName"), dict
        
        rs.MoveNext
    Loop
    
    'load global AppTemplates As Scripting.Dictionary of templates
    Set g_AppTemplates = dictTemplates
    
Exit_Handler:
    'cleanup
    Set dict = Nothing
    Set dictTemplates = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 457  'Duplicate template -- tsys_Db_Templates finds more than one w/ same name
        MsgBox "A duplicate template was found." & vbCrLf & vbCrLf & _
            "When you click 'OK' a query will run to identify the problem template." & vbCrLf & vbCrLf & _
            "You can close the query after it runs (save it if you like)." & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - GetTemplates[mod_Db]:" & vbCrLf & _
            Err.Description, vbExclamation, "Duplicate Db Template Found! [tsys_Db_Templates]"

            Dim strErrorSQL As String
            strErrorSQL = "SELECT TemplateName, Count(TemplateName) AS NumberOfDupes " & _
                    "FROM tsys_Db_Templates " & _
                    "GROUP By TemplateName " & _
                    "HAVING Count(TemplateName) > 1;"

            Dim qdf As DAO.QueryDef
            
            If Not QueryExists("UsysTempQuery") Then
                Set qdf = CurrentDb.CreateQueryDef("UsysTempQuery")
            Else
                Set qdf = CurrentDb.QueryDefs("UsysTempQuery")
            End If
            
            qdf.SQL = strErrorSQL
            
            DoCmd.OpenQuery "USysTempQuery", acViewNormal

            '********** FATAL ERROR ****************
            'terminate *ALL* VBA code to prevent other popups
            'End
            
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTemplates[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetTemplate
' Description:  retrieves template from templates global template dictionary (AppTemplates)
' Parameters:   strTemplate - name of template to fetch (string)
'               params - pipe (|) separated parameter listing w/ parameter name:value pairs (: separated) (string)
' Returns:      template - value of the template (string)
'               most templates are SQL strings, so the SQL string (template) field of the given
'               template name is retrieved
' Assumptions:  tsys_Db_templates correctly list parameter:parameter type values & AppTemplates contain them
'               params do not include PARAM_SEPARATOR w/in them as this is considered a separator
' Throws:       none
' References:   tsys_Db_templates, Microsoft Scripting Runtime (dictionary object)
'   HansUp, June 27, 2013
'   http://stackoverflow.com/questions/17328092/how-to-display-access-query-results-without-having-to-create-temporary-query
' Source/date:  Bonnie Campbell, May 2016
' Revisions:    BLC, 5/19/2016 - initial version
'               BLC, 6/6/2016  - added error handling for duplicate templates, renamed global to g_AppTemplates
' ---------------------------------
Public Function GetTemplate(strTemplate As String, Optional params As String = "") As String
On Error GoTo Err_Handler

    Dim aryParams() As Variant
    Dim ary() As String, ary2() As String
    Dim i As Integer
    Dim template As String, swap As String, param As String

Debug.Print strTemplate

    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates

    template = g_AppTemplates(strTemplate).item("Template")
    
    If Len(params) > 0 Then
    
        'prepare passed in param array --> array contains param:value pairs
        'ary = Split(params, "|")
        If InStr(params, "|") Then
            ary = Split(params, "|")
        Else
            ReDim Preserve ary(0) 'avoid Error #9 subscript out of range
            ary(0) = params
            'ary = Split(params, PARAM_SEPARATOR)
        End If
        
        'prepare array of template parameters w/ their data type
        'aryParams = Split(AppTemplates(strTemplate).item("Params"), "|")
        'AppTemplates("s_tagline").Item("Params").Item("SourceID") --> integer
    
        'iterate through params
        For i = 0 To UBound(ary)
            
            'split name:value pair --> ary2(0) = name, ary2(1) = value
            'If InStr(ary(1), PARAM_SEPARATOR) Then
            If InStr(ary(i), PARAM_SEPARATOR) Then
                ary2 = Split(ary(i), PARAM_SEPARATOR)
            Else
                ary2 = Split(ary(i), ":")
            End If
            'compare datatype to aryParams value
            If IsTypeMatch(ary2(1), g_AppTemplates(strTemplate).item("Params").item(ary2(0))) Then
                
                'prepare replaced value
                swap = "[" & ary2(0) & "]"
                
                'SQL-ize parameter values to avoid SQL syntax errors
                param = SQLencode(ary2(1))
'Debug.Print param
                'swap out the placeholder in the template
                template = Replace(template, swap, ary2(1))
                
            End If
            
        Next
    
    End If
    
    Debug.Print template
    
    GetTemplate = template
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 457  'Duplicate template -- tsys_Db_Templates finds more than one w/ same name
        MsgBox "A duplicate template was found." & vbCrLf & vbCrLf & _
            "When you click 'OK' a query will run to identify the problem template." & vbCrLf & vbCrLf & _
            "You can close the query after it runs (save it if you like)." & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - GetTemplate[mod_Db]:" & vbCrLf & _
            Err.Description, vbExclamation, "Duplicate Db Template Found! [tsys_Db_Templates]"

            Dim strErrorSQL As String
            strErrorSQL = "SELECT TemplateName, Count(TemplateName) AS NumberOfDupes " & _
                    "FROM tsys_Db_Templates " & _
                    "GROUP By TemplateName " & _
                    "HAVING Count(TemplateName) > 1;"

            Dim qdf As DAO.QueryDef
            
            If Not QueryExists("UsysTempQuery") Then
                Set qdf = CurrentDb.CreateQueryDef("UsysTempQuery")
            Else
                Set qdf = CurrentDb.QueryDefs("UsysTempQuery")
            End If
            
            qdf.SQL = strErrorSQL
            
            DoCmd.OpenQuery "USysTempQuery", acViewNormal

            '********** FATAL ERROR ****************
            'terminate *ALL* VBA code to prevent other popups
            'End
        
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTemplate[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     VirtualDAORecordset
' Description:  prepares a virtual -in memory only- DAO recordset
' Parameters:   strTemplate - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  the virtual recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
'   Tom van Stiphout, July 17, 2006
'   https://bytes.com/topic/access/answers/512790-dao-connectionless-recordset
' Source/date:  Bonnie Campbell, May 2016
' Revisions:    BLC, 5/26/2016 - initial version
' ---------------------------------
Public Function VirtualDAORecordset(iCount As Integer, Optional strTable As String = "temp") As Recordset
On Error GoTo Err_Handler

    Dim Counter As Long
    Dim rs As DAO.Recordset
    Dim i As Integer

    With DBEngine
        .BeginTrans
        With .Workspaces(0)(0)

            .Execute "CREATE TABLE " & strTable _
                    & "(RecCount INT CONSTRAINT RecCount UNIQUE);"

            Set rs = .OpenRecordset(strTable, dbOpenTable)
            With rs
                For i = 1 To iCount
                    .AddNew
                    .Fields("RecCount") = i
                    .Update
                Next
                
                .index = "RecCount"
                '.Close
            End With
        End With
    End With

    Set VirtualDAORecordset = rs

Exit_Handler:
    'Set rs = Nothing
    'DBEngine.Rollback
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3010
        Counter = Counter + 1
        strTable = "temp" & CStr(Counter)
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VirtualDAORecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SQLencode
' Description:  sanitizes SQL to remove special characters
' Parameters:   strSQL - SQL to sanitize (string)
' Returns:      strSanitized - sanitized SQL (string)
' Assumptions:
' Throws:       none
' References:
'   Susan Harkins, March 2, 2011
'   http://www.techrepublic.com/blog/microsoft-office/5-rules-for-embedding-strings-in-vba-code/
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/6/2016 - initial version
' ---------------------------------
Public Function SQLencode(strSQL)
On Error GoTo Err_Handler
    
    Dim aryReplace(1, 2) As String
    Dim i As Integer
    Dim strNewSQL As String
    
    'default
    strNewSQL = ""
    
    'exit if no description
    If Len(strSQL) = 0 Then GoTo Exit_Handler
    
    '--------------------------
    ' replacement characters
    '--------------------------
    '   "   Chr(34)
    '   '   Chr(39)
    '--------------------------
    aryReplace(0, 0) = """"
    aryReplace(0, 1) = 34
    aryReplace(1, 0) = "'"
    aryReplace(1, 1) = 39
    
    For i = 0 To UBound(aryReplace, 1)
        strNewSQL = Replace(strSQL, aryReplace(i, 0), "Chr(" & aryReplace(i, 1) & ")")
    Next

    SQLencode = strNewSQL
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SQLencode[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     RemoveVirtualDAORecordset
' Description:  removes a virtual -in memory only- DAO recordset
' Parameters:   strTemplate - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  the virtual recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
'   Tom van Stiphout, July 17, 2006
'   https://bytes.com/topic/access/answers/512790-dao-connectionless-recordset
' Source/date:  Bonnie Campbell, May 2016
' Revisions:    BLC, 5/26/2016 - initial version
' ---------------------------------
Public Sub RemoveVirtualDAORecordset()
On Error GoTo Err_Handler

    DBEngine.Rollback

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveVirtualDAORecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     CreateVirtualADORecordset
' Description:  creates a virtual -in memory only- ADO recordset
' Parameters:   strTemplate - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (ADO.recordset)
' Assumptions:  the virtual recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
'   Danny Lesandrini, November 2, 2009
'   http://www.databasejournal.com/features/msaccess/article.php/3846361/Create-In-Memory-ADO-Recordsets.htm
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
' ---------------------------------
Public Sub CreateVirtualADORecordset(iCount As Integer)
On Error GoTo Err_Handler

'    Dim rsADO As ADODB.Recordset
'    Dim fld As ADODB.Field
'
'    'create rs
'    Set rsADO = New ADODB.Recordset
'    With rsADO
'        .Fields.Append "Number", adInteger, , adFldMayBeNull
'
'        .CursorType = adOpenKeyset
'        .CursorLocation = adUseClient
'        .LockType = adLockPessimistic
'        .Open
'    End With
'
'    'populate rs
'    For i = 0 To iCount - 1
'        rsADO.AddNew
'    Next


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateVirtualADORecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     CreateTempRecordset
' Description:  creates a temporary DAO recordset
' Parameters:   strTemplate - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  the temporary recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
' ---------------------------------
Public Function CreateTempRecordset(iCount As Integer) As DAO.Recordset
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim i As Integer
    
'    strSQL = "SELECT * FROM usys_Temp_Table;"
    
    Set rs = CurrentDb.OpenRecordset("usys_Temp_Table") 'strSQL, dbOpenSnapshot)

    'add records to recordset
    For i = 1 To iCount

        rs.AddNew
        rs.Fields(0) = i 'number integer field
        rs.Update
    Next
    
       
    Set CreateTempRecordset = rs

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateTempRecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          CreateTempRecords
' Description:  fills a temporary table of numbers
'               first clears usys_temp_table of values, then populates w/ desired set of #s
' Parameters:   iCount - number of records (integer)
'               iStart - starting point (integer)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  used for reports when a recordset doesn't exist for the report
'               but it is necessary to repeat the report detail
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
' ---------------------------------
Public Sub CreateTempRecords(iStart As Integer, iCount As Integer)
On Error GoTo Err_Handler

    Dim strSQL As String, strSQLDelete As String, strSQLInsert As String
    Dim i As Integer
    
    'clear table
    strSQLDelete = GetTemplate("d_usys_temp_table")
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQLDelete
    DoCmd.SetWarnings True
    
    'prep for inserts
    strSQL = GetTemplate("i_usys_temp_table")
     
    'add records to table
    For i = iStart To iCount

        strSQLInsert = Replace(strSQL, "[i]", i)

        DoCmd.SetWarnings False
        DoCmd.RunSQL strSQLInsert
        DoCmd.SetWarnings True
    
    Next

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateTempRecords[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub