Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_List
' Level:        Framework module
' Version:      1.01
' Description:  Listview & listbox related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 6/24/2016 - 1.01 - replaced Exit_Function > Exit_Handler
' =================================

' ---------------------------------
'  listview & listbox creation
' ---------------------------------

' =================================
' SUB:          lvwPopulateFromQuery
' Description:  populates listview control from query
' Parameters:   ctrl - listview control
'               strSQL - SQL statement to run for populating listview
'               aryHeadings - heading array for populating values
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from post comment galura.jayar, 4/26/2012
'               http://www.access-programmers.co.uk/forums/showthread.php?t=225070
'               Created 12/10/2014 blc; Last modified 12/10/2014 blc.
' Revisions:    Bonnie Campbell, Dec 10, 2014 - initial version
'               ListView requires Windows Common Control 6.0 (MSCOMCTRL.OCX from c:\windows\system32)
'                   http://support2.microsoft.com/default.aspx?scid=kb;en-us;194784
'                   http://forums.esri.com/Thread.asp?c=93&f=992&t=198775
'               BLC, 4/30/2015 - added error handling & moved from mod_Common_UI to mod_List
' =================================
Public Sub lvwPopulateFromQuery(ctrl As MSComctlLib.ListView, strSQL As String, aryFields As Variant)
On Error GoTo Err_Handler
    Dim dbs As Database
    Dim rs As Recordset
    Dim item As ListItem
    Dim i As Integer
    
    On Error Resume Next
    
    ctrl.ListItems.Clear

    Set dbs = CurrentDb
    Set rs = dbs.OpenRecordset(strSQL, dbOpenSnapshot)

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do Until rs.EOF
            Set item = ctrl.ListItems.Add(, , rs(aryFields(i)))
            For i = 1 To UBound(aryFields)
              item.SubItems(i) = rs(aryFields(i))
            Next
            On Error Resume Next 'continue even in error
            rs.MoveNext
            Set item = Nothing
        Loop
    End If

    Set rs = Nothing

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lvwPopulateFromQuery[mod_List])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          PopulateListHeaders
' Description:  Populate the headers for listbox controls
' Assumptions:  headers are the same as recordset field names
'               sfrms acting as listboxes have static headers already present
' Parameters:   ctrl - listbox control
'               rs   - recordset containing list headers
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/19/2015 - converted to generic to handle listbox-like controls & documentation update
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
' ---------------------------------
Public Sub PopulateListHeaders(ctrl As Control, rs As Recordset)

On Error GoTo Err_Handler

    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer
    Dim frm As Form
    Dim strItem As String, strColHeads As String, aryColWidths() As String

    'exit if subform control (hdrs are static & present on sfrm)
    If ctrl.ControlType = 112 Then
        GoTo Exit_Handler
    End If

    Set frm = ctrl.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.Count
    
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Handler
    End If
    
    'fetch column widths
    aryColWidths = Split(ctrl.ColumnWidths, ";")
    
    'populate column names (if desired)
    If ctrl.ColumnHeads = True Then
        strColHeads = ""
        For i = 0 To cols - 1
            If CInt(aryColWidths(i)) > 0 Then
                strColHeads = strColHeads & rs.Fields(i).Name & ";"
            End If
        Next i
        ctrl.AddItem strColHeads
    End If

    'save headers
    TempVars.Add "lbxHdr", strColHeads

Exit_Handler:
    'leave rs for remaining values
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateListHeaders[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  listview & listbox properties
' ---------------------------------

' =================================
' SUB:          lbxConditionalColor
' Description:  sets lbx text fore color
' Parameters:   ctrl - listbox control
'               tgtCol - column that determines which row(s) fore color should be set to altColor
'               normVal - determining column value for tgtCol  (if tgtCol = normVal then color is set to normColor)
'               altVal - alternate column value for tgtCol (if tgtCol = altVal then color is set to altColor)
'               normColor - string representation of normal listbox row text fore color (vbBlack, vbBlue...)
'               altColor - string representation of color to change listbox row text fore color (vbBlue, vbRed...)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from post comment, 8/2005
'               http://www.tek-tips.com/faqs.cfm?fid=6027
'               Created 12/9/2014 blc; Last modified 12/9/2014 blc.
' Revisions:    Bonnie Campbell, Dec 9, 2014 - initial version
'               ListItem requires Windows Common Control 6.0
'                   http://support2.microsoft.com/default.aspx?scid=kb;en-us;194784
'                   http://forums.esri.com/Thread.asp?c=93&f=992&t=198775
'               BLC, 4/30/2015 - added error handling & moved from mod_Common_UI to mod_List
' =================================
Public Sub lbxConditionalColor(ctrl As ListBox, tgtCol As Integer, normVal As String, altVal As String, normColor As Long, altColor As Long)
On Error GoTo Err_Handler
    Dim Counter As Long
    Dim col As Integer
    
    For Counter = 0 To ctrl.ListCount - 1
        With ctrl
            If CStr(.Column(tgtCol, Counter)) = normVal Then
                For col = 0 To .ColumnCount - 1
                    .Column(col, Counter).ForeColor = normColor
                Next col
            ElseIf CStr(.Column(tgtCol, Counter)) = altVal Then
                For col = 0 To .ColumnCount - 1
                    .Column(col, Counter).ForeColor = altColor
                Next col
            End If
        End With
    Next Counter
    
    'ctrl.refresh

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxConditionalColor[mod_List])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' FUNCTION:     IsListDuplicate
' Description:  Check if item is already on the list
' Assumptions:  -
' Parameters:   lbx - listbox control to check (listbox object)
'               col - column which would hold the item being checked (integer)
'               item - name of item to be checked (string)
' Returns:      boolean - true, if item in list is a duplicate of an existing value in the list
'                         false, if item is not a duplicate
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - updated documentation
' ---------------------------------
Public Function IsListDuplicate(lbx As ListBox, col As Integer, item As String) As Boolean
On Error GoTo Err_Handler
    
    Dim isDupe As Boolean
    Dim i As Integer
    
    'set default
    isDupe = False
    
    'iterate through listbox (use .Column(col,i) vs .ListIndex(i) which results in error 451 property let not defined, property get...)
    For i = 0 To lbx.ListCount
        'check if item exists in listbox
        If lbx.Column(col, i) = item Then
            'duplicate, so exit
            isDupe = True
            GoTo Exit_Handler
        End If
    Next

Exit_Handler:
    IsListDuplicate = isDupe
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsListDuplicate[mod_List])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  listview & listbox item actions
' ---------------------------------

' ---------------------------------
' SUB:          SortList
' Description:  Sorts the listbox item rows alphabetically
' Assumptions:  -
' Parameters:   lbx - listbox to sort
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' MajP, March 22, 2012
' http://www.tek-tips.com/viewthread.cfm?qid=1677888
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
' ---------------------------------
Public Sub SortList(lbx As ListBox) ', orderCol As Integer)

On Error GoTo Err_Handler
  
  Dim strTemp As String
  Dim i As Integer, iHdr As Integer
  Dim j As Integer
  
  'skip first row if lbx has headers
  iHdr = 0
  If Len(TempVars.item("lbxHdr")) > 0 Then
    iHdr = 1
  End If
  
  For i = iHdr To lbx.ListCount - 1
    For j = i + 1 To lbx.ListCount - 1
      If lbx.ItemData(i) > lbx.ItemData(j) Then
        strTemp = lbx.ItemData(i)
        lbx.RemoveItem (i)
        lbx.AddItem lbx.ItemData(j - 1), i
        lbx.RemoveItem (j)
        lbx.AddItem strTemp, j - 1
       End If
     Next j
   Next i

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortList[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetListCount
' Description:  Retrieve the number of items in a list
' Assumptions:  -
' Parameters:   lbx - listbox control to count
'               hdr - if there is a header or not for the listbox (decrements count by 1)
' Returns:      count - number of items in listbox (integer)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/10/2015 - initial version
' ---------------------------------
Public Function GetListCount(lbx As ListBox, HasHeaders As Boolean) As Integer
On Error GoTo Err_Handler

Dim i As Integer

    'Set counts
    i = 0
    If lbx.ListCount > 0 Then
        i = lbx.ListCount - 1
    End If
    
    GetListCount = i

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetListCount[mod_List])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     CountArrayValues
' Description:  count the number of times a specific item is found in an array
' Assumptions:  -
' Parameters:   ary - array to inspect (variant)
'               val - specific value to check for in array (variant)
' Returns:      count - number of items in array (integer)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
' ---------------------------------
Public Function CountArrayValues(ary As Variant, val As Variant) As Integer

On Error GoTo Err_Handler
    
    Dim i As Integer, numItems As Integer

    'default
    numItems = 0
    
    If IsArray(ary) Then
    
        For i = LBound(ary) To UBound(ary)
            If ary(i) = val Then
                numItems = numItems + 1
            End If
        Next
        
    End If
    
    CountArrayValues = numItems

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountArrayValues[mod_List])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          SaveListToTable
' Description:  Save list items to table
' Assumptions:  -
' Parameters:   ctrl - control to iterate through
'               tbl - table being populated
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/8/2015  - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
' ---------------------------------
Public Sub SaveListToTable(ctrl As Control, tbl As String, tblFields As Variant, blnSelectedOnly As Boolean)

On Error GoTo Err_Handler
    
    Dim strSQL As String, strFields As String
    Dim i As Integer, iRow As Integer, jCol As Integer
    
    strSQL = "INSERT INTO " & tbl & " " & tblFields & "VALUES ("
    
    ' prepare fields
    strFields = ""
    For i = 0 To UBound(tblFields)
    
        Select Case tblFields(1, i)
            Case "Integer"
            Case "VarChar"
        End Select
        strFields = strFields
    
    Next

    'iterate through items
    For iRow = 0 To ctrl.ListCount - 1
    
            For jCol = 0 To ctrl.ColumnCount - 1
            
            strSQL = strSQL & "'" & ctrl.Column(jCol, iRow) & "'"
             
            CurrentDb.Execute strSQL, dbFailOnError
            
            Next
    Next 'iRow

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SaveListToTable[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetListRecordset
' Description:  Create a recordset from list items
'               This creates a temporary table for creating the recordset via DAO.
' Assumptions:  -
' Parameters:   lbx - listbox control to get records from (listbox)
'               blnHeaders - true if listbox has headers, false if not (boolean)
'               aryFields - fields (headers & data) from listbox data (array)
'               aryFieldTypes - field types from listbox data (array)
'               tblName - temporary table name (string)
'               blnReplace - true = replace records in the temp table (if it exists)
'                            false = append to records in the temp table (if it exists)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/21/2015 - initial version
'   BLC - 5/26/2015 - revised to SetListRecordset saving listbox rows to temp table
'   BLC - 5/27/2015 - added blnReplace to handle adding additional records to the temp
'                     table from a list
' ---------------------------------
Public Sub SetListRecordset(lbx As ListBox, blnHeaders As Boolean, _
                aryFields As Variant, aryFieldTypes As Variant, tblName As String, _
                blnReplace As Boolean, Optional rsList As DAO.Recordset)
On Error GoTo Err_Handler

Dim iRow As Integer, iStart As Integer, iCol As Integer
Dim strSQL As String, aryFieldNames() As String
Dim aryRecord() As String
Dim aryData() As String
Dim rsProcess As DAO.Recordset
Dim tdf As DAO.TableDef
Dim blnTableExists As Boolean

    'set default table exists
    blnTableExists = False

    'Set start row
    iStart = 0
    If blnHeaders Then iStart = 1
    
    'remove existing table unless it has records
    If TableExists(tblName) Then
    
        'Append Records --> do nothing
        If HasRecords(tblName) And blnReplace = False Then
'            MsgBox "Sorry for the inconvenience, but the table " & tblName & "already exists and has records." & vbCrLf & _
'                "Please check the table's records and remove them (or remove the table)." & vbCrLf & _
'                "Then return here and recreate your list.", _
'                vbCritical, "Oops! " & tblName & " Already Exists!"
'            GoTo Exit_Handler
            
        'Replace Records --> delete existing records
        ElseIf HasRecords(tblName) And blnReplace = True Then
        
            strSQL = "DELETE * FROM " & tblName & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (strSQL)
            DoCmd.SetWarnings True
        
        End If
        blnTableExists = True
    End If

    'create fields for table
    aryFieldNames = Split(CStr(aryFields(0)), ";")

    'handle empty listbox (aside from header record)
    If UBound(aryFields) = 0 Then GoTo Exit_Handler

    'prepare data arrays
    ReDim Preserve aryData(0 To UBound(aryFields) - 1, 0 To UBound(aryFieldNames))
    
    'prepare @ listbox row
    For iRow = 1 To UBound(aryFields)
        
        'get record array
        aryRecord = Split(aryFields(iRow), ";")
       
        'prepare @ listbox field
        For iCol = 0 To UBound(aryFieldNames)
            
            aryData(iRow - 1, iCol) = aryRecord(iCol)
        Next

    Next

    If lbx.ListCount > 0 Then
            
        If Not blnTableExists Then
            
            'create temporary table (if it doesn't exist)
            Set tdf = CurrentDb.CreateTableDef(tblName)
                    
            aryFieldNames = Split(CStr(aryFields(0)), ";")
                
            For iRow = 0 To UBound(aryFieldNames)
                With tdf
                    'add table fields
                    .Fields.Append .CreateField(aryFieldNames(iRow), aryFieldTypes(iRow)) 'GetFieldTypeName(CInt(aryFieldTypes(iRow))))
                
                    'create table & fetch recordset
                    If iRow = UBound(aryFieldNames) Then '- 1 Then
                            
                        ' add table to tabledefs
                        CurrentDb.TableDefs.Append tdf
                                        
                    End If
                    
                End With
            Next
        End If
                
        ' create recordset for the blank table
        Set rsProcess = CurrentDb.OpenRecordset(tblName, dbOpenDynaset)
                
        'add records
        For iRow = 0 To UBound(aryData)
            rsProcess.AddNew
            
            'add each field (second element of aryData)
            For iCol = 0 To UBound(aryData, 2) ' - 1
                
                'add record field values for each record (aryFields - 1, row 0 = field names)
                    rsProcess(aryFieldNames(iCol)).Value = aryData(iRow, iCol)

            Next
            
            rsProcess.Update
                                
        Next
        
        rsProcess.Close
        
    End If

Exit_Handler:
    Set tdf = Nothing
    Set rsProcess = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetListRecordset[mod_List])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' SUB:          SetListRecordset
' Description:  Add list items to existing records in a list recordset table via DAO.
' Assumptions:  Recordset contains the same number and type of fields as the list recordset table.
' Parameters:   tblName - temporary table name (string)
'               rsList - listbox recordset (DAO.recordset)
'               aryFieldNames - table fields (string array)
'               aryFieldTypes - field types (variant array)
'               blnReplace - true = replace records in the temp table (if it exists)
'                            false = append to records in the temp table (if it exists)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 26, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub AddListRecordset(tblName As String, rsList As DAO.Recordset, strFieldNames As String, _
                aryFieldTypes As Variant, blnReplace As Boolean)
On Error GoTo Err_Handler

Dim iRow As Integer, iStart As Integer, iCol As Integer
Dim strSQL As String, aryFieldNames() As String
Dim aryRecord() As String
Dim aryData() As String
Dim rsProcess As DAO.Recordset
Dim tdf As DAO.TableDef
Dim blnTableExists As Boolean

    'set default table exists
    blnTableExists = False

    'Set start row
    iStart = 0
    
    'prepare field names
    aryFieldNames = Split(strFieldNames, ";")
    
    'remove existing table unless it has records
    If TableExists(tblName) Then
            
        'Replace Records --> delete existing records
        If HasRecords(tblName) And blnReplace = True Then
        
            strSQL = "DELETE * FROM " & tblName & ";"
            DoCmd.SetWarnings False
            DoCmd.RunSQL (strSQL)
            DoCmd.SetWarnings True
                
        End If
                
        'Append Records --> do nothing
        
        blnTableExists = True
    End If

    rsList.MoveLast
    If rsList.RecordCount > 0 Then
        
        rsList.MoveFirst
        
        'Create Table
        If Not blnTableExists Then
            
            'create temporary table (if it doesn't exist)
            Set tdf = CurrentDb.CreateTableDef(tblName)

            For iRow = 0 To UBound(aryFieldNames)
                With tdf
                    'add table fields
                    .Fields.Append .CreateField(aryFieldNames(iRow), aryFieldTypes(iRow))
                
                    'create table & fetch recordset
                    If iRow = UBound(aryFieldNames) - 1 Then
                            
                        ' add table to tabledefs
                        CurrentDb.TableDefs.Append tdf
                                        
                    End If
                    
                End With
            Next
        End If
                
        ' create recordset for the blank table
        Set rsProcess = CurrentDb.OpenRecordset(tblName, dbOpenDynaset)
                
        'add records
        For iRow = 0 To rsList.RecordCount - 1 'UBound(aryData)

            rsProcess.AddNew
            
            'add each field (second element of aryData)
            For iCol = 0 To UBound(aryFieldNames) ' - 1
            
                'add record field values for each record (aryFields - 1, row 0 = field names)
                rsProcess(aryFieldNames(iCol)).Value = rsList(aryFieldNames(iCol)).Value

'                iCol = iCol + 1
            Next

            rsProcess.Update
            rsList.MoveNext
            'iRow = iRow + 1
        Next

        rsProcess.Close
        
    End If

Exit_Handler:
    Set tdf = Nothing
    Set rsProcess = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddListRecordset[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetListRecordset
' Description:  Create a recordset from list items saved in temp table
' Assumptions:  Records have already been saved to table via SetListRecordset
' Parameters:   tblName - name of table to check
' Returns:      rs - recordset from list items (or empty recordset), (nothing if no table exists)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 26, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/26/2015 - initial version
' ---------------------------------
Public Function GetListRecordset(tblName As String) As DAO.Recordset
On Error GoTo Err_Handler
    
    'check for table
    If TableExists(tblName) Then

        ' create recordset for the blank table
        Set GetListRecordset = CurrentDb.OpenRecordset(tblName, dbOpenDynaset)
    Else
        'nothing if there isn't a table
        'GetListRecordset = vbNull
    End If
        

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetListRecordset[mod_List])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  listview & listbox item moves
' ---------------------------------

' ---------------------------------
' SUB:          MoveSingleItem
' Description:  moves single list item from one control to another
' Assumptions:  assumes controls are on the same form
' Parameters:   frm - control parent form
'               strSourceControl - name of source control (listbox/listview)
'               strTargetControl - name of destination control (listbox/listview)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 3/5/2015 - added ability to remove from list w/o adding to target if strSourceControl = strTargetControl
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - updated documentation
' ---------------------------------
Public Sub MoveSingleItem(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim strItem As String
    Dim intColumnCount As Integer
    
    'if source = target, just remove the item
    If strSourceControl = strTargetControl Then
        RemoveSelectedItems frm.Controls(strSourceControl)
        GoTo Exit_Handler
    End If
    
    'check for control type
    If frm.Controls(strSourceControl).ControlType = acSubform Then
    'MsgBox frm.Controls(strSourceControl).ControlType, vbOKOnly, "ctrltype"
        'subform control is a continuous form
        Call frm.Controls(strSourceControl).Form.tbxCode_DblClick(False)
        GoTo Exit_Handler
    End If
    
    'check for at *least* one selected item
    If frm.Controls(strSourceControl).ItemsSelected.Count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation, "Oops!"
        GoTo Exit_Handler
    End If
    
    If frm.Controls(strSourceControl).ItemsSelected.Count > 1 Then
        MoveSelectedItems frm, strSourceControl, strTargetControl
        GoTo Exit_Handler
    End If
    
    For intColumnCount = 0 To frm.Controls(strSourceControl).ColumnCount - 1
        strItem = strItem & frm.Controls(strSourceControl).Column(intColumnCount) & ";"
    Next
    
    'remove extra semi-colon (;)
    strItem = Left(strItem, Len(strItem) - 1)

    'Check the length to make sure something is selected
    ' -------------------------------------------------------------------------
    '  NOTE: ListIndex is zero based, so add 1 to remove proper item
    ' -------------------------------------------------------------------------
    If Len(strItem) > 0 Then
        frm.Controls(strTargetControl).AddItem strItem
        frm.Controls(strSourceControl).RemoveItem frm.Controls(strSourceControl).ListIndex + 1
    Else
        MsgBox "Please select an item to move."
    End If


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveSingleItem[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          MoveAllItems
' Description:  moves all list items from one control to another
' Assumptions:  assumes controls are on the same form
' Parameters:   frm - control parent form
'               strSourceControl - name of source control (listbox/listview)
'               strTargetControl - name of destination control (listbox/listview)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 3/5/2015 - added ability to remove from list w/o adding to target if strSourceControl = strTargetControl
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - updated documentation
' ---------------------------------
Public Sub MoveAllItems(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim strItem As String
    Dim intColumnCount As Integer, StartRow As Integer
    Dim lngRowCount As Long
    
    'if source = target, just remove the items
    If strSourceControl = strTargetControl Then
        RemoveSelectedItems (frm.Controls(strSourceControl))
        GoTo Exit_Handler
    End If
        
    'check for at *least* one item
    If frm.Controls(strSourceControl).ListCount = 0 Then
        MsgBox "Your list needs at least one item to move.", vbExclamation, "Oops!"
        GoTo Exit_Handler
    End If
    
    StartRow = 0 'default
    'set start row
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        StartRow = 1
    End If
    
    For lngRowCount = StartRow To frm.Controls(strSourceControl).ListCount - 1
        For intColumnCount = 0 To frm.Controls(strSourceControl).ColumnCount - 1
            strItem = strItem & frm.Controls(strSourceControl).Column(intColumnCount, lngRowCount) & ";"
        Next
        strItem = Left(strItem, Len(strItem) - 1)
        frm.Controls(strTargetControl).AddItem strItem
        strItem = ""
    Next
        
    'clear the list
    frm.Controls(strSourceControl).RowSource = ""
    
    'add back the headers
    ' -------------------------------------------------------------------------
    ' NOTE: target lbx will already have headers, so only add back to source
    ' -------------------------------------------------------------------------
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        frm.Controls(strSourceControl).AddItem TempVars("lbxHdr")
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveAllItems[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          MoveSelectedItems
' Description:  move items selected to another list
' Assumptions:  -
' Parameters:   frm - control parent form (form object)
'               strSourceControl - name of source list (string)
'               strTargetControl - name of destination list (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' ManningFan, January 30,2015
' http://bytes.com/topic/access/answers/765291-populating-1-listbox-another-listbox
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 3/5/2015 - added ability to remove from list w/o adding to target if strSourceControl = strTargetControl
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - updated documentation
' ---------------------------------
Public Sub MoveSelectedItems(frm As Form, strSourceControl As String, strTargetControl As String)
    
On Error GoTo Err_Handler
    
    Dim iRow As Integer, StartRow As Integer, i As Integer, X As Integer, iRemovedItems As Integer
    Dim arySelectedItems() As Integer
    Dim blnDimensioned As Boolean
    Dim strItem As String
    
    'if source = target, just remove the items
    If strSourceControl = strTargetControl Then
        RemoveSelectedItems (frm.Controls(strSourceControl))
        GoTo Exit_Handler
    End If
    
    'check for at *least* one selected item
    If frm.Controls(strSourceControl).ItemsSelected.Count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation, "Oops!"
        GoTo Exit_Handler
    End If
    
    StartRow = 0 'default
    'set start row
    If frm.Controls(strSourceControl).ColumnHeads = True Then
        StartRow = 1
    End If
    
    'add back the header if it doesn't exist
    If frm.Controls(strTargetControl).ColumnHeads = True And frm.Controls(strTargetControl).ListCount = 0 Then
       strItem = TempVars.item("lbxHdr") & strItem
       frm.Controls(strTargetControl).AddItem strItem
    End If
    
    'generate array of selected items
    For iRow = StartRow To frm.Controls(strSourceControl).ListCount - 1
    
        'fetch array of selected items
        '--------------------------------------------------
        ' if > 1 item selected, other selected items
        ' deselected when first source item removed
        '--------------------------------------------------
        If frm.Controls(strSourceControl).Selected(iRow) Then
            
            'Array dimensioned?
            If blnDimensioned = True Then
                      
                'Yes ==> extend array 1 element largee than current upper bound
                '        w/o "Preserve" keyword previous elements erased w/ resizing
                ReDim Preserve arySelectedItems(0 To UBound(arySelectedItems) + 1) As Integer
                      
            Else
                      
                'No ==> dimension it and flag as dimensioned
                ReDim arySelectedItems(0 To 0) As Integer
                blnDimensioned = True
                          
            End If
                  
            'Add to last element in the array.
            arySelectedItems(UBound(arySelectedItems)) = iRow
        End If
    
    Next
    
    'set default
    iRemovedItems = 0
    
    'iterate through selected items
    For X = LBound(arySelectedItems) To UBound(arySelectedItems)
                        
        iRow = arySelectedItems(X) - iRemovedItems
            
        'clear string
        strItem = ""
        
        'add all columns
        For i = 0 To frm.Controls(strSourceControl).ColumnCount
            strItem = strItem & frm.Controls(strSourceControl).Column(i, iRow) & ";"
        Next i
        
        'add to target
        frm.Controls(strTargetControl).AddItem strItem
        
        'remove from source
        frm.Controls(strSourceControl).RemoveItem iRow
            
        'adjust list after removal
        If UBound(arySelectedItems) > 0 Then
            iRemovedItems = iRemovedItems + 1
        End If
    
    Next X

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveSelectedItems[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RemoveSelectedItems
' Description:  Removes selected items from a listbox by re-creating rowsource
' Assumptions:  lbx is a listbox control (not a continuous subform which may act as a listbox control)
' Parameters:   lbx - Listbox to remove selected items from
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' ADezii, April 13, 2010
' http://bytes.com/topic/access/answers/885569-remove-selected-items-list-box-microsoft-access
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
' ---------------------------------
Public Sub RemoveSelectedItems(lbx As ListBox)
On Error GoTo Err_Handler
  
    Dim intRow As Integer, iCol As Integer
    Dim strBuild As String
     
    With lbx
      If .ItemsSelected.Count = 0 Then Exit Sub
     
      For intRow = 0 To .ListCount - 1
        If Not .Selected(intRow) Then
            For iCol = 0 To .ColumnCount - 1
                strBuild = strBuild & .Column(iCol, intRow) & ";"
            Next
        End If
      Next
     
      strBuild = Left$(strBuild, Len(strBuild) - 1)
     
      .RowSource = strBuild
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveSelectedItems[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  listview & listbox item changes
' ---------------------------------
' ---------------------------------
' SUB:          RemoveListDupes
' Description:  Remove listbox duplicate values
' Assumptions:  -
' Parameters:   lbx - listbox to check
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' matsushita, September 27, 2006
' https://social.msdn.microsoft.com/Forums/vstudio/en-US/0799668c-36dd-42d9-9599-3085a6c0581f/how-to-remove-duplicate-values-in-listbox-
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/13/2015 - commented out SortList due to bug which removes headers & values
' ---------------------------------
Public Sub RemoveListDupes(lbx As ListBox)

On Error GoTo Err_Handler

    Dim index As Integer, Count As Integer
    Dim lastItem As String
    
    'sort listbox
 '   SortList lbx
    
    Count = lbx.ListCount

    'check sorted listbox for duplicates & remove
    If Count > 1 Then
    
        lastItem = lbx.ItemData(Count - 1)

        For index = Count - 2 To 0 Step -1
            If lbx.ItemData(index) = lastItem And Len(lbx.ItemData(index)) > 0 Then
                'duplicate
                lbx.RemoveItem (index)
            Else
                lastItem = lbx.ItemData(index)
            End If
        Next
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveListDupes[mod_List])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ClearList
' Description:  Clear values from listbox control
' Assumptions:  -
' Parameters:   lbx - Listbox control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - updated documentation
' ---------------------------------
Public Sub ClearList(lbx As ListBox)

On Error GoTo Err_Handler

    'clear listbox items
    lbx.RowSource = ""

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearList[mod_List])"
    End Select
    Resume Exit_Handler
End Sub