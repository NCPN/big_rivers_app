Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_QA
' VERSION:      1.02
' Description:  QA related properties, functions & subroutines
'
' Source/date:  Bonnie Campbell, 8/22/2014
' Revisions:    BLC, 8/22/2014 - 1.00 - initial version
'               BLC, 6/12/2015 - 1.01 - replaced TempVars.item(... with TempVars("...
'               BLC, 4/4/2016  - 1.02 - changed Exit_Procedure/Exit_Function > Exit_Handler
' =================================

' ---------------------------------
' FUNCTION:     UpdateQAResults
' Description:  Updates the data validation results table
'
'   This function requires that the database contain tbl_QA_Results with the
'   following fields:  Query_name (txt 100), Time_frame (txt 30), Data_scope (tinyint),
'   Query_type (txt 20), Query_result (txt 50), Query_run_time (date/time),
'   Query_description (memo), Query_expression (memo);
'   optional fields:  Remedy_desc (memo), Remedy_date (date/time), QA_user (txt 50),
'   Is_done (yes/no)
'
'   Also required is the query "qsys_QA_query_expressions":  SELECT MSysObjects.Name,
'   MSysQueries.Attribute, MSysQueries.Expression FROM MSysObjects LEFT JOIN MSysQueries
'   ON MSysObjects.Id = MSysQueries.ObjectId WHERE (((MSysObjects.Name) Like "qa*") And
'   ((MSysQueries.Attribute) = 8 Or (MSysQueries.Attribute) = 10) And ((MSysQueries.Expression)
'   Is Not Null)) ORDER BY MSysObjects.Name;
'
'   Also required is the query "qsys_QA_query_errors":
'   SELECT tbl_QA_Results.Query_name, "No longer exists, but in result set" AS Issue,
'   tbl_QA_Results.Time_frame FROM MSysObjects RIGHT JOIN tbl_QA_Results ON
'   MSysObjects.Name = tbl_QA_Results.Query_name WHERE (((tbl_QA_Results.Time_frame)
'   = [Forms]![frm_Switchboard]![cTimeframe]) And ((MSysObjects.Name) Is Null))
'   UNION SELECT MSysObjects.Name AS Query_name, "Not in result set" AS Issue,
'   tbl_QA_Results.Time_frame FROM MSysObjects LEFT JOIN tbl_QA_Results ON
'   MSysObjects.Name = tbl_QA_Results.Query_name WHERE (((MSysObjects.Name) Like "qa_*")
'   And ((tbl_QA_Results.Time_frame) = [Forms]![frm_Switchboard]![cTimeframe])
'   And ((tbl_QA_Results.Query_name) Is Null))
'   UNION SELECT tbl_QA_Results.Query_name, "Not running properly" AS Issue,
'   tbl_QA_Results.Time_frame FROM tbl_QA_Results WHERE (((tbl_QA_Results.Time_frame)
'   =[Forms]![frm_Switchboard]![cTimeframe]) AND ((tbl_QA_Results.Query_run_time) Is Null))
'   OR (((tbl_QA_Results.Time_frame)=[Forms]![frm_Switchboard]![cTimeframe]) AND
'   ((tbl_QA_Results.Query_result) Is Null));
'
'   The following code assumes the following naming convention for all validation queries:
'   1) prefix of "qa_" for all queries that are intended to return results to the user
'       (subqueries may have a prefix of "qasub_")
'   2) 4th character may be used for sorting queries hierarchically (e.g., a-z)
'   3) 5th and 6th characters are the sort order within each level of the hierarchy
'   4) 7th character indicates the severity of the error if it returns records:  1=critical,
'       2=warning, 3=information
'   5) 8th character is an underbar "_" and from the 9th character on is a descriptive name
'       with words separated by an underbar "_" character (no spaces or special characters!)
'
' Parameters:   blnUpdateAll - boolean, false if only one of the QA queries is to be updated
'               strSingleQName - string, the name of the single query to be updated
'               blnCreateNew - boolean, true if a new query needs to be created (given new
'                   filter criteria)
' Returns:      none
' Throws:       none
' References:   ChangeDelimiter
' Source/date:  John R. Boetsch, 2006 February
' Revisions:    JRB, 3/9/2006 - added a line to handle nulls for query descriptions
'               JRB, 5/9/2006 - added function call to clean the query expression string
'                   by replacing double quotes with single quotes (thanks to Simon Kingston)
'               JRB, 10/3/2006 - added code to include timeframe in the insert into statement
'               JRB, 11/14/2007 - revised the naming convention description above and
'                   updated Mid() statement to reflect revised naming conventions
'               JRB, 12/17/2007 - updated code to make sure records specify time frame as
'                   well as query name; code also now allows a single query to be updated
'                   instead of the full set, through use of blnUpdateAll and strSingleQName;
'                   code also allows user to use the Is_done flag to sort records
'               JRB, 5/14/2008 - updated qsys_QA_query_errors statement to filter results
'                   records by the current selected timeframe
'               JRB, 9/17/2008 - updated by adding reference to frm_Progress_Meter to show query
'                   names while running queries (helps for optimizing slow QA queries)
'               JRB, 9/19/2008 - updated tbl_QA_Results by adding Data_scope field
'               JRB, 2/20/2009 - added blnCreateNew
'               JRB, 6/10/2009 - added qdfs.Refresh to capture query description changes
' ---------------------------------

' ---------------------------------
' SUB:          UpdateQAResults
' Description:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  John R. Boetsch, May 2008
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/29/2014 - updated to use TempVars.Item("Timeframe") vs. cTimeframe
'               BLC, 8/22/2014 - shifted to mod_QA & dropped fxn prefix
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 4/4/2016  - changed Exit_Function > Exit_Handler, dbCurrent to CurrentDb
' ---------------------------------
Public Function UpdateQAResults(Optional blnUpdateAll As Boolean = True, _
    Optional strSingleQName As String, Optional blnCreateNew As Boolean = False)

    On Error GoTo Err_Handler

    Dim qdf As dao.QueryDef     ' Individual query objects
    Dim qdfs As dao.QueryDefs   ' The database query set
    Dim strTimeframe As String  ' Data timeframe, from the switchboard
    Dim intScope As Integer     ' Indicates whether or not certified records are included
                                '   in query runs: 0=no, 1=yes, 2=both certified and uncertified
    Dim strSQL As String        ' The SQL statement
    Dim strQName As String      ' Name of the query
    Dim strQType As String      ' Type of query (embedded in name; 1=critical, 2=warning, 3=info)
    Dim strQDesc As String      ' Description of the query
    Dim strQResult As String    ' N records currently returned by the query
    Dim strTResult As String    ' N records previously returned by the query, from the QA table
    Dim strQExp As String       ' WHERE clause expression of the query
    Dim dtRunTime As Date       ' Query run time
    Dim intNErrors As Integer   ' Number of queries that have update problems
    Dim intNQueries As Integer  ' Number of queries total
    Dim varReturn As Variant    ' For manipulating the system meter
    Dim intI As Integer         ' Counter for updating the system meter
    Dim frm As Form             ' Reference to the progress popup form
    Dim strProgForm As String   ' Name of the progress popup form
    Dim strProgress As String   ' Progress bar string

    Set qdfs = DBEngine(0)(0).QueryDefs
    qdfs.Refresh

    DoCmd.Hourglass True

    dtRunTime = Now()   ' Set the run time variable to now

    ' Initialize the progress popup form
    strProgForm = "frm_Progress_Meter"
    DoCmd.OpenForm strProgForm
    Set frm = Forms!frm_Progress_Meter
    frm.Caption = " Running validation queries"
    frm!txtPercent = 0
    intNQueries = 0

    For Each qdf In qdfs
        If Left(qdf.Name, 3) = "qa_" Then intNQueries = intNQueries + 1
    Next qdf

    On Error Resume Next
    ' Initialize the system meter to indicate progress
    varReturn = SysCmd(acSysCmdInitMeter, "Running validation queries", intNQueries)
    intI = 0

    strTimeframe = "unknown"
    If IsNull(Forms!frm_QA_Tool.cmbTimeframe) = False Then
        strTimeframe = Forms!frm_QA_Tool.cmbTimeframe
    Else
        If IsNull(TempVars("Timeframe")) = False Then _
            strTimeframe = TempVars("Timeframe")
    End If
    intScope = Forms!frm_QA_Tool.optgScope ' Me.optgScope

    For Each qdf In qdfs
        If Left(qdf.Name, 3) = "qa_" Then
            intI = intI + 1
            ' Update the percent complete in the progress popup
            frm!txtPercent = Round(100 * intI / intNQueries)
            ' Update the progress bar in the progress popup with sequential "Û" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNQueries), "Û")
            frm!txtProgress = strProgress
            ' Update the progress meter in the status bar
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            strQName = qdf.Name
            ' Update the query name in the progress popup
            frm!txtMsg = strQName
            frm.Repaint
            ' Create the record if all queries are being updated
            If blnUpdateAll Or (blnCreateNew And strQName = strSingleQName) Then
                strQType = Mid(strQName, 7, 1)
                If strQType = "" Then strQType = "0"
                ' Create the statement to insert new records
                strSQL = "INSERT INTO tbl_QA_Results " & _
                    "(Query_name, Time_frame, Data_scope, Query_type, Is_done) " & _
                    "SELECT """ & strQName & """ AS Query_name, """ & _
                    strTimeframe & """ AS Time_frame, " & intScope & _
                    " AS Data_scope, """ & strQType & _
                    """ AS Query_type, 0 AS Is_done;"
                ' Run the SQL code
                CurrentDb.Execute strSQL
            End If

            ' Run the following query if all queries are being updated, or if the current
            '   query matches the selected query in the form
            If blnUpdateAll Or strQName = strSingleQName Then
                ' Look up the number of records returned on the last run, for comparison
                strTResult = DLookup("Query_result", "tbl_QA_Results", _
                    "[Query_name]=""" & strQName & """ AND [Time_frame]=""" _
                    & strTimeframe & """ AND [Data_scope]=" & intScope)
                ' Update existing records to refresh the results
                strQResult = DCount("*", qdf.Name)  ' the number of records currently returned
                ' Create the statement to add the query description and expression
                '   (expression not always present)
                strQDesc = " - none defined - "         ' Default in case of error
                strQDesc = qdf.Properties("Description")    ' Query description
                ' Clean up any double-quotes in the description and change to single quotes
                strQDesc = ChangeDelimiter(strQDesc)
                strQExp = " - none defined - "          ' Default in case of error
                strQExp = DLookup("Expression", "qsys_QA_query_expressions", "[Name]=""" & _
                    strQName & """")
                ' Clean up any double-quotes in the expression and change to single quotes
                strQExp = ChangeDelimiter(strQExp)

                If strQResult = "0" And strQType <> "3" Then
                    ' If the number of records is zero and the query type is not 'information'
                    '   then set the Is_done flag to True
                    strSQL = "UPDATE tbl_QA_Results SET tbl_QA_Results.Query_expression = """ _
                        & strQExp & """, tbl_QA_Results.Query_description = """ & strQDesc & _
                        """, tbl_QA_Results.Query_result = """ & strQResult _
                        & """, tbl_QA_Results.Query_run_time = #" & dtRunTime & _
                        "#, tbl_QA_Results.Is_done = TRUE " & _
                        "WHERE (((tbl_QA_Results.Query_name)=""" & strQName & _
                        """) AND ((tbl_QA_Results.Time_frame)=""" & strTimeframe & _
                        """) AND ((tbl_QA_Results.Data_scope)=" & intScope & "));"

                ElseIf strTResult <> strQResult Then
                    ' If the number of records has changed then set Is_done flag to False
                    strSQL = "UPDATE tbl_QA_Results SET tbl_QA_Results.Query_expression = """ _
                        & strQExp & """, tbl_QA_Results.Query_description = """ & strQDesc & _
                        """, tbl_QA_Results.Query_result = """ & strQResult _
                        & """, tbl_QA_Results.Query_run_time = #" & dtRunTime & _
                        "#, tbl_QA_Results.Is_done = FALSE " & _
                        "WHERE (((tbl_QA_Results.Query_name)=""" & strQName & _
                        """) AND ((tbl_QA_Results.Time_frame)=""" & strTimeframe & _
                        """) AND ((tbl_QA_Results.Data_scope)=" & intScope & "));"

                Else
                    ' Build the update query without changing the Is_done flag
                    strSQL = "UPDATE tbl_QA_Results SET tbl_QA_Results.Query_expression = """ _
                        & strQExp & """, tbl_QA_Results.Query_description = """ & strQDesc & _
                        """, tbl_QA_Results.Query_result = """ & strQResult _
                        & """, tbl_QA_Results.Query_run_time = #" & dtRunTime & _
                        "#  WHERE (((tbl_QA_Results.Query_name)=""" & strQName & _
                        """) AND ((tbl_QA_Results.Time_frame)=""" & strTimeframe & _
                        """) AND ((tbl_QA_Results.Data_scope)=" & intScope & "));"
                End If
                ' Run the SQL code
                CurrentDb.Execute strSQL
            End If
        End If
    Next qdf

    On Error GoTo Err_Handler
    ' Notify the user if queries are not updating properly
    intNErrors = DCount("*", "qsys_QA_query_errors")
    If intNErrors > 0 Then
        If intNErrors = 1 Then
            MsgBox "There is 1 query not updating properly.", vbCritical, _
                "Validation query error"
        Else
            MsgBox "There are " & intNErrors & " queries not updating properly.", vbCritical, _
                "Validation query error"
        End If
        DoCmd.OpenQuery "qsys_QA_query_errors", , acReadOnly
    End If

    If blnUpdateAll Then
        ' Pause for a second before proceeding
        Dim varPause, varStart
        varPause = 1
        varStart = Timer
        Do While Timer < varStart + varPause
            DoEvents    ' Yield to other processes
        Loop
    End If

Exit_Handler:
    On Error Resume Next
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Close acForm, strProgForm, acSaveNo
    Set frm = Nothing
    DoCmd.Hourglass False
    Set qdfs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateQAResults[mod_QA])"
    End Select
    Resume Exit_Handler

End Function