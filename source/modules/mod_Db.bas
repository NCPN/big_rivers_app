Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Db
' Level:        Framework module
' Version:      1.08
' Description:  Database related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added mod_db_Templates subs/functions - qryExists
'               BLC, 5/26/2016 - 1.02 - added VirtualDAORecordset()
'               BLC, 6/6/2016  - 1.03 - added error handling for duplicate templates, renamed global to g_AppTemplates
'                                       also added SQL sanitization (escape/replace special chars)
'               BLC, 6/9/2016  - 1.04 - added CreateTempRecords()
'               BLC, 10/4/2016 - 1.05 - added GetParamsFromSQL()
'               BLC, 10/11/2016 - 1.06 - added IsRecordset(), FieldCount(), MaxDbFieldCount()
'               BLC, 10/20/2016 - 1.07 - added IsLinked()
'               BLC, 1/9/2017 - 1.08   - added SetTempVar()
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
            rsOut.Fields(iCount).value = rsB.Fields(iCount).value
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
' FUNCTION:     DbTableExists
' Description:  determine if a table exists w/in a database
' Assumptions:  used for retrieving table field data for mapping fields, etc.
' Parameters:   tbl - name of database table (string)
'               tdfRefresh - whether to refresh table defs or not (boolean, optional)
'               db - database to reference (DAO.database object)
' Returns:      whether or not table exists (boolean)
' Throws:       none
' References:
'   David W. Fenton, June 7, 2010
'   http://stackoverflow.com/questions/2985513/check-if-access-table-exists
'   Based on testing, when passed an existing db variable, this function is fastest
'   Tony Toews, unknown
'   http://www.granite.ab.ca/access/temptables.htm
'   David Fenton's functino originally based on Tony Toew's function in TempTables.MDB
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
' ---------------------------------
Public Function DbTableExists(tbl As String, Optional tdfRefresh As Boolean, _
                                Optional db As DAO.Database) As Boolean
On Error GoTo Err_Handler
  
  Dim tdf As DAO.TableDef

  'set db if passed
  If db Is Nothing Then Set db = CurrentDb()
  
  'refresh tables
  If tdfRefresh Then db.TableDefs.Refresh
  
  Set tdf = db(tbl)
  
  DbTableExists = True

Exit_Handler:
    'cleanup
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3265
            DbTableExists = False
        Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DbTableExists[mod_Db])"
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
' FUNCTION:     GetDescription
' Description:  retrieves object description property
' Assumptions:  object has a description property
' Parameters:   obj - item to check (object)
' Returns:      description text for object (string)
' Throws:       none
' References:
'   Allen Browne, April, 2010
'   http://allenbrowne.com/func-06.html
' Source/date:  Bonnie Campbell, September 2016 for NCPN tools
' Revisions:    BLC, 9/16/2016 - initial version
' ---------------------------------
Public Function GetDescription(obj As Object) As String
On Error GoTo Err_Handler

    GetDescription = obj.Properties("Description")

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescription[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' https://support.office.com/en-us/article/VarType-Function-1e08636c-1892-40c2-aff3-2b894389e82d?ui=en-US&rs=en-US&ad=US&fromAR=1
'   vbEmpty         0   Empty (uninitialized)
'   vbNull          1   Null (no valid data)
'   vbInteger       2   Integer
'   vbLong          3   Long integer
'   vbSingle        4   Single-precision floating-point number
'   vbDouble        5   Double-precision floating-point number
'   vbCurrency      6   Currency value
'   vbDate          7   Date Value
'   vbString        8   String
'   vbObject        9   Object
'   vbError         10  Error Value
'   vbBoolean       11  Boolean value
'   vbVariant       12  Variant (used only with arrays of variants)
'   vbDataObject    13  A data access object
'   vbDecimal       14  Decimal value
'   vbByte          17  Byte value
'   vbUserDefinedType 36    Variants that contain user-defined types
'   vbArray         8192    Array

' ---------------------------------
' FUNCTION:     FieldTypeName
' Description:  retrieves field type property name from the numeric field type
' Assumptions:  -
' Parameters:   fld - field to retrieve type for (DAO.field)
' Returns:      name for the field type (string)
' Throws:       none
' References:
'   Allen Browne, April, 2010
'   http://allenbrowne.com/func-06.html
'   TofuBug     May 28, 2015
'   http://stackoverflow.com/questions/30511987/why-does-vartype-always-return-8204-for-arrays
' Source/date:  Bonnie Campbell, September 2016 for NCPN tools
' Revisions:    BLC, 9/16/2016 - initial version
' ---------------------------------
Public Function FieldTypeName(fld As DAO.field) As String
On Error GoTo Err_Handler
    
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) ' fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        
'        'Arrays
'        Case vbArray:
'            strReturn = "Array"                     '8192
'
'        Case Is > 8192
'            Select Case (fld.Type - 8192)
'                Case vbString                       '8 --> Overall 8200 = 8192+8
'                    strReturn = "String Array"
'                Case vbVariant                      '12 --> Overall 8204 = 8192+12
'                    strReturn = "Variant Array"
'                Case Else
'                    strReturn = "Field type " & fld.Type & " unknown"
'            End Select
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FieldTypeName[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     VarTypeName
' Description:  retrieves var type name from the numeric variable type
' Assumptions:  -
' Parameters:   vType - numeric type (integer)
' Returns:      name for the type (string)
' Throws:       none
' References:
'   Allen Browne, April, 2010
'   http://allenbrowne.com/func-06.html
'   TofuBug     May 28, 2015
'   http://stackoverflow.com/questions/30511987/why-does-vartype-always-return-8204-for-arrays
' Source/date:  Bonnie Campbell, September 2016 for NCPN tools
' Revisions:    BLC, 9/16/2016 - initial version
' ---------------------------------
Public Function VarTypeName(vType As Integer) As String
On Error GoTo Err_Handler
    
    Dim strReturn As String    'Name to return

    Select Case CLng(vType) ' vType is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
'            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
'            Else
'                strReturn = "AutoNumber"
'            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
'            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
'            Else
'                strReturn = "Text (fixed width)"        '(no interface)
'            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
'            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
'            Else
'                strReturn = "Hyperlink"
'            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        
        'Arrays
        Case vbArray:
            strReturn = "Array"                     '8192

        Case Is > 8192
            Select Case (vType - 8192)
                Case vbString                       '8 --> Overall 8200 = 8192+8
                    strReturn = "String Array"
                Case vbVariant                      '12 --> Overall 8204 = 8192+12
                    strReturn = "Variant Array"
                Case Else
                    strReturn = "Field type " & vType & " unknown"
            End Select
        Case Else: strReturn = "Field type " & vType & " unknown"
    End Select

    VarTypeName = strReturn

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VarTypeName[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FetchDbTableFieldInfo
' Description:  retrieves field information from a database table
' Parameters:   tbl - name of database table (string)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  used for retrieving table field data for mapping fields, etc.
' Throws:       none
' References:
'   David W. Fenton, July 27, 2010
'   http://stackoverflow.com/questions/3343922/get-column-names
'   HansUp, October 18, 2013
'   http://stackoverflow.com/questions/19452952/how-to-count-number-of-fields-in-a-table
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
' ---------------------------------
Public Function FetchDbTableFieldInfo(tbl As String) As Variant 'DAO.Recordset
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset ', rsFields As ADODB.Recordset
    Dim fld As DAO.field
    Dim aryFieldInfo() As Variant
    Dim icols As Integer, iCol As Integer
    Dim strTypeName As String
    
    Set db = CurrentDb()
    
    'determine if table is in database
    If Not DbTableExists(tbl) Then GoTo Err_Handler
    
    Set rs = db.OpenRecordset(tbl)
    
'    Set rsFields = New ADODB.Recordset
    
    'get count
    icols = rs.Fields.Count
    iCol = 0
    
    ReDim Preserve aryFieldInfo(0 To icols - 1)
    
    'iterate through fields
    For Each fld In rs.Fields
'        Debug.Print fld.Name
'        Debug.Print fld.Attributes
'        Debug.Print fld.Size
'        Debug.Print fld.Properties
'        Debug.Print fld.Required
'        Debug.Print fld.Type
'        Debug.Print fld.ValidationRule
'        With rsFields
'            .Append
'        End With
        
'        Debug.Print (fld)
        
        'fetch name for type
'        GetFieldTypeName fld
        strTypeName = VarTypeName(fld.Type)

        With fld
                
            'prepare array of info
            aryFieldInfo(iCol) = .Name & "|" & _
                            .Type & "|" & _
                            .Size & "|" & _
                            .Required & "|" & _
                            .AllowZeroLength & "|" & _
                            strTypeName

        End With
        
        iCol = iCol + 1
    Next

    FetchDbTableFieldInfo = aryFieldInfo

    'cleanup
    'Set fld = Nothing
'    Set rs = Nothing
'    Set db = Nothing
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FetchDbTableFieldInfo[mod_Db])"
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
Public Sub GetTemplates(Optional strSyntax As String = "", Optional Params As String = "")

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, key As String
    Dim value As Variant
    
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
                
                Set value = dictParam

            Else
                value = Nz(rs.Fields(ary(i)), "")
            End If
            
            'add key if it isn't already there
            If Not dict.Exists(key) Then
                If IsNull(value) Then MsgBox key, vbOKCancel, "is NULL"
                'Debug.Print Nz(Value, key & "-NULL")
                dict.Add key, value
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
            
            qdf.sql = strErrorSQL
            
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
Public Function GetTemplate(strTemplate As String, Optional Params As String = "") As String
On Error GoTo Err_Handler

    Dim aryParams() As Variant
    Dim ary() As String, ary2() As String
    Dim i As Integer
    Dim Template As String, swap As String, param As String

Debug.Print strTemplate

    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates

    Template = g_AppTemplates(strTemplate).item("Template")
    
    If Len(Params) > 0 Then
    
        'prepare passed in param array --> array contains param:value pairs
        'ary = Split(params, "|")
        If InStr(Params, "|") Then
            ary = Split(Params, "|")
        Else
            ReDim Preserve ary(0) 'avoid Error #9 subscript out of range
            ary(0) = Params
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
                Template = Replace(Template, swap, ary2(1))
                
            End If
            
        Next
    
    End If
    
Debug.Print Template
    
    GetTemplate = Template
    
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
            
            qdf.sql = strErrorSQL
            
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

' ---------------------------------
' SUB:          CreateTempTable
' Description:  creates a temp table from an array containing table field definitions
' Assumptions:  field array is 1-dimensional
'               fields are represented with name|type|length/size|required|allowZLS
'               only name|type are required (except for dbText where length/size is also reqd)
'               ex: "col1|CStr(dbText)|2|True|False"
'               data array includes same # of columns as fields array
' Parameters:   tblName - table name (string)
'               aryFields - array containing field definitions (variant)
' Returns:
' References:
' Source/date:  Bonnie Campbell, September 20 2016
' Revisions:    BLC, 9/20/2016 - initial version
' ---------------------------------
Public Sub CreateTempTable(tblName As String, aryFields() As Variant)
On Error GoTo Err_Handler

    'check for blank table name or no fields
    If Not IsArray(aryFields) Or Len(tblName) = 0 Then GoTo Exit_Handler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.field
    Dim item As Variant, fldDef As Variant
    Dim i As Integer

    Set db = CurrentDb()
    
    'delete it if it already exists
    If TableExists(tblName) Then RemoveTempTable (tblName)
    
    Set tdf = db.CreateTableDef(tblName)
    
    'prepare array
    For Each item In aryFields
    
        'fldDef(0) = name, fldDef(1) = type, fldDef(2) = length (as applicable)
        fldDef = Split(item, "|")
        
        'establish field w/ name & type
        Set fld = tdf.CreateField(fldDef(0), CLng(fldDef(1)))
        
        'add attributes - size (if applicable), required & allow ZLS
        For i = LBound(fldDef) To UBound(fldDef)
            Select Case i
                Case 0  'column name
                Case 1  'column type
                Case 2  'column size
                    fld.Size = fldDef(2)
                Case 3  'column required
                    fld.Required = fldDef(3)
                Case 4  'column allow ZLS
                    fld.AllowZeroLength = fldDef(4)
                Case 5
                Case Else
            End Select
        Next
        tdf.Fields.Append fld
        tdf.Fields.Refresh
    Next
    
    'add table
    db.TableDefs.Append tdf
    
    'update window
    db.TableDefs.Refresh
    RefreshDatabaseWindow
    
    'cleanup
'    db.Close

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateTempTable[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RemoveTempTable
' Description:  removes a temp table from database
' Assumptions:  -
' Parameters:   tblName - table name (string)
' Returns:      -
' References:   -
' Source/date:  Bonnie Campbell, September 20 2016
' Revisions:    BLC, 9/20/2016 - initial version
' ---------------------------------
Public Sub RemoveTempTable(tblName As String)
On Error GoTo Err_Handler

    'check for blank table name
    If Len(tblName) = 0 Then GoTo Exit_Handler

    'check if table exists
    If TableExists(tblName) Then
    
        'delete table
        DoCmd.DeleteObject acTable, tblName
    
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveTempTable[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddRecords
' Description:  adds records to a recordset & table
' Assumptions:  -
' Parameters:   rs - (DAO.Recordset)
'               aryCols - field/column names (string array)
'               aryData - data for each record (variant array)
' Returns:      -
' References:
'   simoco, February 9, 2014
'   http://stackoverflow.com/questions/21885101/can-you-use-a-variable-for-the-field-name-when-using-addnew-to-a-record-set
' Source/date:  Bonnie Campbell, September 20 2016
' Revisions:    BLC, 9/20/2016 - initial version
' ---------------------------------
Public Sub AddRecords(rs As DAO.Recordset, aryCols() As String, aryData() As Variant, _
                            delimiter As String)
On Error GoTo Err_Handler

    Dim aryRecord As String
    Dim i As Integer, j As Integer
    Dim strColName As String

    With rs
        
        'add new record
        .AddNew
        
        
        'iterate through data records
        For i = 0 To UBound(aryData)
        
            'get record array
            aryRecord = Split(aryData(i), delimiter)
            
            'iterate through columns
            For j = 0 To UBound(aryCols)
            
                strColName = aryCols(j)
                
                .Fields(strColName) = aryRecord
            
            Next
        
        Next
    
    
    End With
    
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddRecords[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetParamsFromSQL
' Description:  extracts parameters from SQL string
' Assumptions:  -
' Parameters:   sql - SQL to retrieve parameters from(string)
' Returns:      params - delimited string of parameters and parameter types (string)
' References:   -
' Source/date:  Bonnie Campbell, September 20 2016
' Revisions:    BLC, 9/20/2016 - initial version
' ---------------------------------
Public Function GetParamsFromSQL(sql As String) As String
On Error GoTo Err_Handler

    Dim Params As String
    
    'default
    Params = ""
    
    If Len(sql) > 0 Then
        If InStr(sql, "PARAMETERS ") Then
            Dim delimPos As Integer
            
            Params = Replace(sql, "PARAMETERS ", "")
            delimPos = InStr(Params, ";")
            Params = Left(Params, delimPos - 1)
            Params = Replace(Params, ", ", "|")
            Params = Replace(Params, " ", ":")
            
            'convert TEXT(#) values to STRING
            If InStr(Params, "TEXT(") Then
                'remove TEXT( )
                Params = Replace(Params, "TEXT(", "STRING")
                Params = Replace(Params, ")", "")
                'remove numerics
                Params = RemoveChars(Params, False)
            End If
            
        End If
    End If
    
Exit_Handler:
    GetParamsFromSQL = Params
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParamsFromSQL[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ListTables
' Description:  List database tables
' Assumptions:  -
' Parameters:   ShowMSysTables - whether or not to show msys_ tables (boolean)
'               ShowTsysTables - whether or not to show tsys_ tables (boolean)
'               ShowUsysTables - whether or not to show usys_ tables (boolean)
'               ShowLinkedTables - whether or not to show linked tables (boolean)
' Returns:      tables - delimited string of tables (string)
' References:   -
'   Daniel Pineault, June 10, 2010
'   http://www.devhut.net/2010/06/10/ms-access-vba-list-the-tables-in-a-database/
'   HansUp, December 17, 2013
'   http://stackoverflow.com/questions/20643263/how-can-one-search-tabledefs-for-linked-tables
' Source/date:  Bonnie Campbell, October 6 2016
' Revisions:    BLC, 10/6/2016 - initial version
'               BLC, 10/20/2016 - revised to include linked tables, added Tsys, Usys parameters
' ---------------------------------
Public Function ListTables(ShowMSysTables As Boolean, _
                            ShowTSysTables As Boolean, _
                            ShowUSysTables As Boolean, _
                            ShowLinkedTables As Boolean) As String
On Error GoTo Err_Handler

    Dim tdf As DAO.TableDef
    Dim tbls As String
    
    'default
    tbls = ""
    
    'fetch tables
    For Each tdf In CurrentDb.TableDefs
'Debug.Print tdf.Name
        'handle MSys tables
        If Len(tdf.Name) > Len(Replace(tdf.Name, "MSys", "")) And ShowMSysTables = False Then GoTo Continue
        
        'handle tsys tables
        If Len(tdf.Name) > Len(Replace(tdf.Name, "tsys", "")) And ShowMSysTables = False Then GoTo Continue
                
        'handle usys tables
        If Len(tdf.Name) > Len(Replace(tdf.Name, "usys", "")) And ShowMSysTables = False Then GoTo Continue
        
        'handle linked tables
        If Len(tdf.Connect) > 0 And ShowLinkedTables = False Then GoTo Continue
        
        tbls = tbls & "|" & tdf.Name
        
Continue:
    Next
    
    'trim starting delimiter
    tbls = Right(tbls, Len(tbls) - 1)
'    Debug.Print tbls
    
Exit_Handler:
    ListTables = tbls
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ListTables[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsRecordset
' Description:  Determines if the object is a recordset or not
' Assumptions:
'               Error handling is ignored since rs.Recordcount would produce
'               an error if rs is not a recordset. In that case isRS remains
'               false and that is returned through the Exit_Handler
' Parameters:   rs - recordset object (object)
' Returns:      isRS - if object was determined to be a recordset (boolean)
'                      true = is a recordset object, false = is not a recordset object
' References:   -
' Source/date:  Bonnie Campbell, October 11 2016
' Revisions:    BLC, 10/11/2016 - initial version
' ---------------------------------
Public Function IsRecordset(rs As Object)
On Error GoTo Err_Handler

    Dim isRS As Boolean
    
    isRS = False
    
    If Not rs Is Nothing Then
            
'        If Not IsError(IsNumeric(rs.RecordCount)) Then isRS = True
        If IsNumeric(rs.RecordCount) Then isRS = True
    
    End If

Exit_Handler:
    IsRecordset = isRS
    Exit Function
Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - IsRecordset[mod_Db])"
'    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FieldCount
' Description:  Determines the number of fields in a table/query
' Assumptions:  -
' Parameters:   TableName - name of table/query (variant)
' Returns:      FieldCount - number of fields (variant)
' References:
'   Sinndho, May 8, 2012
'   http://www.dbforums.com/showthread.php?1678970-Count-the-number-of-columns-(fields)-in-a-table
' Source/date:  Bonnie Campbell, October 11 2016
' Revisions:    BLC, 10/11/2016 - initial version
' ---------------------------------
Public Function FieldCount(ByVal TableName As String) As Long
'Public Function FIeldCount(ByVal TableName As Variant) As Variant <<-- if including in query
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset(TableName, dbOpenSnapshot)
    
    FieldCount = rs.Fields.Count

Exit_Handler:
    'cleanup
    rs.Close
    Set rs = Nothing
    
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FieldCount[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     MaxDbFieldCount
' Description:  Determines the maximum number of fields in Db tables/queries
' Assumptions:  -
' Parameters:   -
' Returns:      MaxDbFieldCount - maximum number of fields (variant)
' References:
'   Sinndho, May 8, 2012
'   http://www.dbforums.com/showthread.php?1678970-Count-the-number-of-columns-(fields)-in-a-table
' Source/date:  Bonnie Campbell, October 11 2016
' Revisions:    BLC, 10/11/2016 - initial version
' ---------------------------------
Public Function MaxDbFieldCount() As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim max As Long
    Dim qtName As String
    
    Set db = CurrentDb
    
    'default
    max = 0
    
    For Each tdf In db.TableDefs
        
        If tdf.Fields.Count > max Then
            max = tdf.Fields.Count
            qtName = tdf.Name
        End If

    Next
    
    For Each qdf In db.QueryDefs
    
        If qdf.Fields.Count > max Then
            max = qdf.Fields.Count
            qtName = qdf.Name
        End If

    Next
    
    Debug.Print qtName
    
    MaxDbFieldCount = max

Exit_Handler:
    'cleanup
    Set tdf = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MaxDbFieldCount[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsLinked
' Description:  Determines if a table is linked
' Assumptions:  -
' Parameters:   tblName - name of table to evaluate (string)
' Returns:      IsLinked - whether table is linked (boolean)
'                          returns true for types 4 (ODBC linked), 6 (other linked)
'                                  false for type 1 (non-linked tables)
' References:
'   Douglas J. Steele, February 20, 2009
'   http://www.pcreview.co.uk/threads/check-if-a-table-is-linked.3748757/
' Source/date:  Bonnie Campbell, October 20, 2016
' Revisions:    BLC, 10/20/2016 - initial version
' ---------------------------------
Public Function IsLinked(tblName As String) As Boolean
On Error GoTo Err_Handler
    
    IsLinked = Nz(DLookup("Type", "MSysObjects", "Name='" & tblName & "'"), 0) <> 1

Exit_Handler:
    'cleanup
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsLinked[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          SetTempVar
' Description:  Checks if TempVar exists, creates it if not, & sets value
' Assumptions:  -
' Parameters:   strVar - TempVar name (string)
'               Val - value to set (variant)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  -
' Adapted:      Bonnie Campbell, January 9, 2017 - for NCPN tools
' Revisions:
'   BLC - 1/9/2017 - initial version
' ---------------------------------
Public Sub SetTempVar(strVar As String, Val As Variant)
On Error GoTo Err_Handler

    If Not TempVars(strVar) Is Nothing Then
        TempVars(strVar) = Val
    Else
        TempVars.Add strVar, Val
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetTempVar[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub