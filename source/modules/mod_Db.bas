Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Db
' Level:        Framework module
' Version:      1.01
' Description:  Database related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added mod_db_Templates subs/functions - qryExists
' =================================

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
        "ORDER BY tsys_BE_Updates.Update_ID;", dbOpenDynaset)

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
            If bRunAll = True Or rs![Is_done] = False Then
                DoCmd.SetWarnings False
                strSQL = rs![SQL_statement]
                DoCmd.RunSQL strSQL
                With rs
                    .Edit
                    ![Run_date] = Now()
                    ![Is_done] = True
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
        For iCount = 0 To rsB.Fields.count - 1
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

    For i = 0 To [TempVars].count - 1
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
' SUB:     GetSQLTemplates
' Description:  loads SQL templates (queries as SQL string) into memory as a dictionary object
'               with query SQL strings available without querying the db tsys_SQL_templates table
' Parameters:
' Returns:      dictionary object stored in tempVars.Item("SQL")
' Assumptions:  placing
' Throws:       none
' References:   tsys_Db_templates, Microsoft Scripting Runtime (dictionary object)
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - initial version
'               BLC, 5/13/2016 - shifted from mod_Db_Templates to mod_Db & adjusted to match tsys_Db_Templates
' ---------------------------------
Public Sub GetSQLTemplates(Optional strVersion As String = "")

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, key As String, Value As String
    
    'handle default
    strSQLWhere = " WHERE Is_Supported > 0"
    
    If Len(strVersion) > 0 Then
        strSQLWhere = " AND LCase(versionID) = LCase(" & strVersion & " )"
    End If
    
    'sql -> ID, Version, IsSupported, Context, Syntax, TemplateName, Params, Template, Remarks,
    '       EffectiveDate, RetireDate, CreateDate, CreatedBy_ID, LastModified, LastModifiedBy_ID
    strSQL = "SELECT * FROM tsys_Db_Templates" & strSQLWhere
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSQL)
    
    'handle no records
    If rst.EOF Then
        MsgBox "Sorry, no templates were found for this database version.", vbExclamation, _
            "Linked Database Templates Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Handler
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
    
Exit_Handler:
    'cleanup
    Set dict = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSQLTemplates[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub