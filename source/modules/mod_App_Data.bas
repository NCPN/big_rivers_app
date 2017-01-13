Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.19
' Description:  data functions & procedures specific to this application
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015  - 1.00 - initial version
'               BLC - 2/18/2015 - 1.01 - included subforms in fillList
'               BLC - 5/1/2015  - 1.02 - integerated into Invasives Reporting tool
'               BLC - 5/22/2015 - 1.03 - added PopulateList()
'               BLC - 6/3/2015  - 1.04 - added IsUsedTargetArea()
'               BLC - 5/5/2016  - 1.05 - added GetRiverSegments(), GetProtocolVersion()
'                                        changed to Exit_Handler vs. Exit_Function
'               BLC - 6/28/2016 - 1.06 - added ToggleIsActive(), revised getParkState() to GetParkState()
'               BLC - 7/26/2016 - 1.07 - added SetRecord(), GetRecords()
'               BLC - 7/28/2016 - 1.08 - added UpsertRecord()
'               BLC - 7/30/2016 - 1.09 - added ToggleSensitive()
'               BLC - 8/8/2016  - 1.10 - updated UpsertRecord() for additional forms
'               BLC - 9/1/2016  - 1.11 - added UploadSurveyFile(), updated UpsertRecord()
'               BLC - 9/13/2016 - 1.12 - added FetchAddlData()
'               BLC - 9/21/2016 - 1.13 - updated SetRecord() i_login parameters
'               BLC - 9/22/2016 - 1.14 - added templates
'               BLC - 10/16/2016 - 1.15 - fixed PopulateCombobox() to properly set recordset
'               BLC - 10/19/2016 - 1.16 - renamed UploadSurveyFile() to UploadCSVFile() to genericize
'               BLC - 10/24/2016 - 1.17 - updated SetRecord(), ToggleIsActive()
'               BLC - 10/28/2016 - 1.18 - updated i_task, TempVars("ContactID") -> TempVars("AppUserID")
'               BLC - 1/9/2017   - 1.19 - revised UpsertRecord from ContactID to ID,
'                                         added GetRecords templates
' =================================

' ---------------------------------
' SUB:          fillList
' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
' Assumptions:  Either a listbox or subform control is being populated
' Parameters:   frm - main form object
'               ctrl - either:
'                      lbx - main form listbox object (for filling a listbox control)
'                      sfrm - subform object (for populating a subform control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/6/2015  - initial version
'   BLC, 2/18/2015 - adapted to include subform as well as listbox controls
'   BLC, 5/1/2015  - integrated into Invasives Reporting tool
' ---------------------------------
Public Sub fillList(frm As Form, ctrlSource As Control, Optional ctrlDest As Control)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    
    'output to form or listbox control?
   
    'determine data source
    Select Case ctrlSource.Name
    
        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
            strQuery = "qry_Active_Datasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            
        Case "lbxSpecies", "lbxTgtSpecies", "fsub_Species_Listbox" 'Species
            strQuery = "qry_Plant_Species"
            strSQL = CurrentDb.QueryDefs(strQuery).sql
            
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    If Not ctrlDest Is Nothing Then
        'populate list & headers
        PopulateList ctrlSource, rs, ctrlDest
    Else
        'populate only ctrlSource headers
        PopulateListHeaders ctrlSource, rs
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fillList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          PopulateList
' Description:  Populate listbox and similar controls from recordset
' Assumptions:  -
' Parameters:   ctrlSource - source control (listbox/listview)
'               rs - recordset used to populate control (recordset object)
'               ctrlDest - destination control (listbox/listview)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' krish KM, Aug. 27, 2014
' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/20/2015 - changed from tbxMasterCode to tbxLUCode
'   BLC - 5/22/2015 - moved to mod_App_Data from mod_List
' ---------------------------------
Public Sub PopulateList(ctrlSource As Control, rs As Recordset, ctrlDest As Control)

On Error GoTo Err_Handler

    Dim frm As Form
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
    Dim strItem As String, strColHeads As String, aryColWidths() As String

    Set frm = ctrlSource.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.Count
    
    'address no records
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Handler
    End If
    
    'handle sfrm controls (acSubform = 112)
    If ctrlSource.ControlType = acSubform Then
        Set ctrlSource.Form.Recordset = rs
        
        ctrlSource.Form.Controls("tbxCode").ControlSource = "Code"
        ctrlSource.Form.Controls("tbxSpecies").ControlSource = "Species"
        'ctrlSource.Form.Controls("tbxMasterCode").ControlSource = "Master_PLANT_Code"
        ctrlSource.Form.Controls("tbxLUCode").ControlSource = "LUCode"
        ctrlSource.Form.Controls("tbxTransectOnly").ControlSource = "Transect_Only"
        ctrlSource.Form.Controls("tbxTgtAreaID").ControlSource = "Target_Area_ID"
        
        'set the initial record count (MoveLast to get full count, MoveFirst to set display to first)
        rs.MoveLast
        ctrlSource.Parent.Form.Controls("lblSfrmSpeciesCount").Caption = rs.RecordCount & " species"
        rs.MoveFirst
        
        GoTo Exit_Handler
    End If
    
    'fetch column widths array
    aryColWidths = Split(ctrlSource.ColumnWidths, ";")
    
    'count number of 0 width elements
    iZeroes = CountArrayValues(aryColWidths, "0")
        
    'clear out existing values
    ClearList ctrlSource
    
    'populate column names (if desired)
    If ctrlSource.ColumnHeads = True Then
        PopulateListHeaders ctrlSource, rs
        
        'populate second listbox headers if present
        If ctrlDest.ColumnHeads = True Then
            ClearList ctrlDest
            PopulateListHeaders ctrlDest, rs
        End If
    End If
    
    'populate data
    Select Case ctrlSource.RowSourceType
        Case "Table/Query"
            Set ctrlSource.Recordset = rs
        Case ""
            
            'initialize
            i = 0
            
            Do Until rs.EOF
            
                'initialize item
                strItem = ""
                    
                'generate item
                For j = 0 To cols - 1
                    'check if column is displayed width > 0
                    If CInt(aryColWidths(j)) > 0 Then
                    
                        strItem = strItem & rs.Fields(j).Value & ";"
                    
                        'determine how many separators there are (";") --> should equal # cols
                        matches = (Len(strItem) - Len(Replace$(strItem, ";", ""))) / Len(";")
                        
                        'add item if not already in list --> # of ; should equal cols - 1
                        'but # in list should only be # of non-zero columns --> cols - iZeroes
                        If matches = cols - iZeroes Then
                            ctrlSource.AddItem strItem
                            'reset the string
                            strItem = ""
                        End If
                    
                    End If
                
                Next
                
                i = i + 1
                
                rs.MoveNext
            Loop
        Case "Field List"
    End Select

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddListToTable
' Description:  Populate table from listbox
' Assumptions:  -
' Parameters:   lbx - listbox control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 3, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/3/2015 - initial version
' ---------------------------------
Public Sub AddListToTable(lbx As ListBox)

On Error GoTo Err_Handler

Dim aryFields() As String
Dim aryFieldTypes() As Variant
Dim strCode As String, strSpecies As String, strLUCode As String
Dim iRow As Integer, iTransectOnly As Integer, iTgtAreaID As Integer
    
    iRow = lbx.ListCount - 1 'Forms("frm_Tgt_Species").Controls("lbxTgtSpecies").ListCount - 1
    
    ReDim Preserve aryFields(0 To iRow)
        
    'header row (iRow = 0)
    aryFields(0) = "Code;Species;LUCode;Transect_Only;Target_Area_ID"   'iRow = 0
    aryFieldTypes = Array(dbText, dbText, dbText, dbInteger, dbInteger)

    'data rows (iRow > 0)
    For iRow = 1 To lbx.ListCount - 1
        
        ' ---------------------------------------------------
        '  NOTE: listbox column MUST have a non-zero width to retrieve its value
        ' ---------------------------------------------------
         strCode = lbx.Column(0, iRow) 'column 0 = Master_PLANT_Code (Code)
         strSpecies = lbx.Column(1, iRow) 'column 1 = Species name (Species)
         strLUCode = lbx.Column(2, iRow) 'column 2 = LU_Code (LUCode)
         iTransectOnly = Nz(lbx.Column(3, iRow), 0) 'column 3 = Transect_Only (TransectOnly)
         iTgtAreaID = Nz(lbx.Column(4, iRow), 0) 'column 4 = Target_Area_ID (TgtAreaID)
        
        aryFields(iRow) = strCode & ";" & strSpecies & ";" & strLUCode & ";" & iTransectOnly & ";" & iTgtAreaID
        
    Next
    
    'save the existing records to temp_Listbox_Recordset & replace any existing records
    SetListRecordset lbx, True, aryFields, aryFieldTypes, "temp_Listbox_Recordset", True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetParkState
' Description:  Retrieve the state associated with a park (via tlu_Parks)
' Assumptions:  Park state is properly identified in tlu_Parks
' Parameters:   parkCode - 4 character park designator
' Returns:      ParkState - 2 character state abbreviation
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
'   BLC - 6/28/2016  - revised to uppercase GetParkState vs getParkState
' ---------------------------------
Public Function GetParkState(ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim State As String, strSQL As String
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 ParkState FROM tlu_Parks WHERE ParkCode LIKE '" & ParkCode & "';"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        State = rs.Fields("ParkState").Value
    End If
   
    'return value
    GetParkState = State
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParkState[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     getListLastModifiedDate
' Description:  Retrieve the last modified date with a park (via tbl_Target_List)
' Assumptions:  -
' Parameters:   tgtYear - 4 digit year of list (integer)
'               parkCode - 4 character park designator (string)
' Returns:      date - last modified date (mmm-d-yyyy H:nn AMPM format) for the specified target list (string)
'                      if NULL (no last modified date) returns empty string
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/10/2015  - initial version
' ---------------------------------
Public Function getListLastModifiedDate(TgtYear As Integer, ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim strCriteria As String

    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Or TgtYear < 2000 Then
        GoTo Exit_Handler
    End If
    
    'set lookup criteria
    strCriteria = "Park_Code LIKE '" & ParkCode & "' AND CInt(Target_Year) = " & CInt(TgtYear)
    
    'Debug.Print strCriteria
        
    'lookup last modified date & return value
    getListLastModifiedDate = Nz(Format(DLookup("Last_Modified", "tbl_Target_List", strCriteria), "mmm-d-yyyy H:nn AMPM"), "")
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getListLastModifiedDate[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsUsedTargetArea
' Description:  Determine if the target area is in use by a target list
' Parameters:   TgtAreaID - target area idenifier (integer)
' Returns:      boolean - true if target area is in use, false if not
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
' ---------------------------------
Public Function IsUsedTargetArea(TgtAreaID As Integer) As Boolean

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    'default
    IsUsedTargetArea = False
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 Target_Area_ID FROM tbl_Target_Species WHERE Target_Area_ID = " & TgtAreaID & ";"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        IsUsedTargetArea = True
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsUsedTargetArea[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:     PopulateTree
' Description:  Populate the treeview control
' Parameters:   TreeType - treeview type (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
' ---------------------------------
Public Sub PopulateTree(TreeType As String)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Select Case TreeType
        Case "ParkSiteFeatureTransectPlot"
            strSQL = "SELECT * FROM qry_Park_Site_Feature_Transect_Plot"
        Case "Photo"
    End Select
                   
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        
        
        
        
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateTree[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          PopulateCombobox
' Description:  Populate priority/status comboboxes
' Parameters:   cbx - combobox control to populate (ComboBox)
'               BoxType - type of combobox, priority or status (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
'  https://msdn.microsoft.com/en-us/library/office/ff845773.aspx
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
'   BLC - 10/12/2016 - fixed to set combobox recordset
' ---------------------------------
Public Sub PopulateCombobox(cbx As ComboBox, BoxType As String)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Select Case BoxType
        Case ""
        Case "priority"
            strSQL = "SELECT ID, Priority FROM Priority ORDER BY Sequence ASC;"
        Case "status"
            strSQL = "SELECT ID, Status FROM Status ORDER BY Sequence ASC;"
    End Select
 
     'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
 
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        Set cbx.Recordset = rs
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateCombobox[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetProtocolVersion
' Description:  Retrieve protocol version, effective & retire dates
' Assumptions:  Assumes only one version of the protocol is active at once
' Parameters:   blnAllVersions - indicator if all versions should be retrieved (boolean)
' Returns:      Protocol name, version, effective & retire dates, last modified date
' Note:         To retrieve values, data must be retrieved from the array:
'                   ary(0,0) = ProtocolName
'                   ary(1,0) = Version
'                   ary(2,0) = EffectiveDate
'                   ary(3,0) = RetireDate
'                   ary(4,0) = LastModified
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
' ---------------------------------
Public Function GetProtocolVersion(Optional blnAllVersions As Boolean = False) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strWhere As String
    Dim Count As Integer
    Dim metadata() As Variant
   
    'handle only appropriate park codes
    If blnAllVersions Then
        strWhere = ""
    Else
        strWhere = "WHERE RetireDate IS NULL"
    End If
    
    'generate SQL
'    strSQL = "SELECT ProtocolName, Version, EffectiveDate, RetireDate, LastModified FROM Protocol " _
'                & strWHERE & ";"
    strSQL = GetTemplate("s_protocol_info", "strWHERE" & PARAM_SEPARATOR & strWhere)
    
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
        
    If rs.BOF And rs.EOF Then GoTo Exit_Handler
        
    With rs
        .MoveLast
        .MoveFirst
        Count = .RecordCount
    
        metadata = rs.GetRows(Count)
 
        .Close
    End With
    
    'return value
    GetProtocolVersion = metadata
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetProtocolVersion[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetSOPMetadata
' Description:  Retrieve SOP metadata (abbreviation code, #, version, effective date)
' Assumptions:  Assumes only one active/effective SOP # for a given area
' Parameters:   area - area covered by the SOP (string)
' Returns:      SOP metadata - Code, SOP #, Version, EffectiveDate
' Note:         To retrieve value, data must be retrieved from the array:
'                   ary(0,0) = SOP #
'               Assuming there is only one matching SOP for each area
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
'   BLC - 5/11/2016 - revised to getting full SOP metadata vs. number only
' ---------------------------------
Public Function GetSOPMetadata(area As String) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
        
    'generate SQL
    '---------------------------------------------------------------------
    ' NOTE: use * vs % for the LIKE wildcard
    '       if it is not used strSQL will work in a query directly,
    '       but will fail to return records via a VBA recordset
    '       So    "...LIKE '" & LCase(area) & "*';"   works
    '       But   "...LIKE '" & LCase(area) & "%';"   does not (except in direct Query SQL)
    '
    ' c.f.  Hans Up, May 17, 2011 & discussion
    '       http://stackoverflow.com/questions/6037290/use-of-like-works-in-ms-access-but-not-vba
    '---------------------------------------------------------------------
    strSQL = GetTemplate("s_sop_metadata", "area" & PARAM_SEPARATOR & LCase(area))
    
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
        
    'return value
    Set GetSOPMetadata = rs
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSOPNum[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetRiverSegments
' Description:  Retrieve the river segments associated with a park
' Assumptions:  River segments are properly associate w/ park
' Parameters:   ParkCode - 4 character park designator
' Returns:      segments - river segments (Green, CAC, GBC, Yampa, CBC, GBC, etc.)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
' ---------------------------------
Public Function GetRiverSegments(ParkCode As String) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim Count As Integer
    Dim segments() As Variant
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
    strSQL = GetTemplate("s_get_river_segments", "ParkCode" & PARAM_SEPARATOR & ParkCode)

            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    Count = rs.RecordCount
    
    segments = rs.GetRows(Count)
 
    rs.Close
    
    'return value
    GetRiverSegments = segments
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRiverSegments[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetParkID
' Description:  Retrieve the ID associated with a park
' Assumptions:  -
' Parameters:   ParkCode - 4 character park designator (string)
' Returns:      ID - unique park identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 1/12/2017  - revised to use GetRecords() vs. GetTemplate()
' ---------------------------------
Public Function GetParkID(ParkCode As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
'    strSQL = GetTemplate("s_park_id", "ParkCode" & PARAM_SEPARATOR & ParkCode)
            
    'fetch data
'    Set db = CurrentDb
    Set rs = GetRecords("s_park_id") 'db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetParkID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParkID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetRiverSegmentID
' Description:  Retrieve the ID associated with a River
' Assumptions:  -
' Parameters:   segment - river segment designator (string)
' Returns:      ID - unique river identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 1/17/2017  - revise to use GetRecords() vs. GetTemplate()
' ---------------------------------
Public Function GetRiverSegmentID(segment As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate River codes
    If Len(segment) < 1 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
'    strSQL = GetTemplate("s_river_segment_id", "waterway" & PARAM_SEPARATOR & segment)
            
    'fetch data
'    Set db = CurrentDb
    Set rs = GetRecords("s_river_segment_id") 'db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetRiverSegmentID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRiverSegmentID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetSiteID
' Description:  Retrieve the ID associated with a site
' Assumptions:  -
' Parameters:   ParkCode - park designator (4-character string)
'               SiteCode - site designator (2-character string)
' Returns:      ID - unique site identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
' ---------------------------------
Public Function GetSiteID(ParkCode As String, SiteCode As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate River codes
    If Len(ParkCode) <> 4 Or Len(SiteCode) <> 2 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
    strSQL = GetTemplate("s_site_id_by_code", _
            "ParkCode" & PARAM_SEPARATOR & ParkCode & _
            "|sitecode" & PARAM_SEPARATOR & SiteCode)
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetSiteID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSiteID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetFeatureID
' Description:  Retrieve the ID associated with a feature
' Assumptions:  -
' Parameters:   ParkCode - park designator (4-character string)
'               Feature - feature designator (2-character string)
' Returns:      ID - unique feature identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 10/4/2016  - update to use parameter query
' ---------------------------------
Public Function GetFeatureID(ParkCode As String, Feature As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate River codes
    If Len(ParkCode) <> 4 Or Len(Feature) < 1 Then
        GoTo Exit_Handler
    End If
    
'    'generate SQL
'    strSQL = GetTemplate("s_feature_id", _
'            "ParkCode" & PARAM_SEPARATOR & ParkCode & _
'            "|feature" & PARAM_SEPARATOR & Feature)
'
'    'fetch data
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset(strSQL)
'
'    If rs.BOF And rs.EOF Then GoTo Exit_Handler
'
'    rs.MoveLast
'    rs.MoveFirst
'
'    If Not rs.BOF And rs.EOF Then
'        ID = rs.GetRows(1)
'    End If
'
'    rs.Close
    
    'return value
    GetFeatureID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFeatureID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          ToggleIsActive
' Description:  Toggle IsActive button click actions
' Assumptions:  -
' Parameters:   Context - form context for the action (string)
'               ID - id of record to toggle (long)
'               IsActive - state to change IsActiveFlag to (Byte), 0 - active, 1 - inactive
'                          optional for ModWentworth scale retire date
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/28/2016 - shifted from ContactList form to mod_App_Data
'   BLC - 10/20/2016 - added ModWentworth retire date toggle
'   BLC - 10/24/2016 - revised to use SetRecord()
' ---------------------------------
Public Sub ToggleIsActive(Context As String, ID As Long, Optional IsActive As Byte)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'
'    Select Case Context
'        Case "Contact"
'            strSQL = GetTemplate("u_contact_isactive_flag", _
'                      "IsActiveFlag" & PARAM_SEPARATOR & IsActive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'        Case "Site"
'            strSQL = GetTemplate("u_site_isactive_flag", _
'                      "IsActiveFlag" & PARAM_SEPARATOR & IsActive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'        Case "ModWentworthScale"
'            strSQL = GetTemplate("u_mod_wentworth_retireyear", _
'                      "RetireDate" & PARAM_SEPARATOR & Date & "|ID" & _
'                      PARAM_SEPARATOR & ID)
'    End Select
'
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (strSQL)
'    DoCmd.SetWarnings True
    
    Dim Template As String
    
    Select Case Context
        Case "Contact"
            Template = "u_contact_isactive_flag"
        Case "Site"
            Template = "u_site_isactive_flag"
        Case "ModWentworthScale"
            Template = "u_mod_wentworth_retireyear"
            
    End Select
    
    Dim Params(0 To 3) As Variant
    
    Params(0) = Template
    Params(1) = ID
    Params(2) = IIf(InStr(Template, "wentworth") > 0, Year(Date), IsActive)
        
    SetRecord Template, Params
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleIsActive[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ToggleSensitive
' Description:  Toggle Sensitive button click actions
' Assumptions:  -
' Parameters:   Context - form context for the action (string)
'               ID - id of record to toggle (long)
'               Sensitive - state to change SensitiveFlag to (Byte), 0 - active, 1 - inactive
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
' ---------------------------------
Public Sub ToggleSensitive(Context As String, ID As Long, Sensitive As Byte)
On Error GoTo Err_Handler
    
    Dim strSQL As String
    
    Select Case Context
        Case "Location"
            strSQL = GetTemplate("u_location_sensitive_flag", _
                      "SensitiveFlag" & PARAM_SEPARATOR & Sensitive & _
                      "|ID" & PARAM_SEPARATOR & ID)
        Case "species"
            strSQL = GetTemplate("u_species_sensitive_flag", _
                      "SensitiveFlag" & PARAM_SEPARATOR & Sensitive & _
                      "|ID" & PARAM_SEPARATOR & ID)
    End Select

    DoCmd.SetWarnings False
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleSensitive[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          GetRecords
' Description:  Retrieve records based on template
' Assumptions:  -
' Parameters:   Template - SQL template name (string)
' Returns:      rs - data retrieved (recordset)
' Throws:       none
' References:
'   user1938742, October 17, 2014
'   http://stackoverflow.com/questions/26422970/run-query-with-parameters-and-display-in-listbox-ms-access-2013
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/22/2016 - added templates
'   BLC - 1/9/2017 - added templates
' ---------------------------------
Public Function GetRecords(Template As String) As DAO.Recordset
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .sql = GetTemplate(Template)
        
            Select Case Template
                        
                Case "s_access_level"
                    '-- required parameters --
                    .Parameters("lvl") = TempVars("tempLvl")
                    
                    'clear the tempvar
                    TempVars.Remove "tempLvl"
                
                Case "s_app_enum_list"
                    '-- required parameters --
                    .Parameters("etype") = TempVars("EnumType")
                
                Case "s_contact_list"
                    '-- required parameters --
                    'N/A
                                
                Case "s_datasheet_defaults_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                
                Case "s_datasheet_defaults_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                                
                Case "s_events_by_feature"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                    .Parameters("feat") = TempVars("Feature")
                                
                Case "s_event_by_park_river_w_location"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                    
                Case "s_events_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_events_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")

                Case "s_events_list_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_feature_id"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_feature_list"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                                        
                Case "s_feature_list_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                                        
                Case "s_get_parks"
                    '-- required parameters --
                                    
                Case "s_location_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_location_by_park_river_segment"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("seg") = TempVars("River")
                
                Case "s_location_list_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_mod_wentworth_for_eventyr"
                    '-- required parameters --
                    'default event year to current year if not passed in
                    .Parameters("eventyr") = Nz(TempVars("EventYear"), Year(Now))
                
                Case "s_park_id"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                
                Case "s_river_segment_id"
                    '-- required parameters --
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_river_list"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                
                Case "s_site_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_site_by_park_river_segment"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("seg") = TempVars("River")
                
                Case "s_site_list_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_site_list_by_park_river_segment"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_site_list_active"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("seg") = TempVars("River")
            
                Case "s_species_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    
                Case "s_top_rooted_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .sql = Replace(Replace(.sql, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_rooted_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .sql = Replace(Replace(.sql, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_understory_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .sql = Replace(Replace(.sql, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_understory_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .sql = Replace(Replace(.sql, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_woody_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .sql = Replace(Replace(.sql, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_woody_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .sql = Replace(Replace(.sql, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_veg_walk_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    
                    'revise TOP X --> 8 is replaced by # blanks to return (from # rows remaining)
                    .sql = Replace(.sql, 8, TempVars("Blanks"))
                
                Case "s_veg_walk_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                    
                    'revise TOP X --> 8 is replaced by # blanks to return (from # rows remaining)
                    .sql = Replace(.sql, 8, TempVars("Blanks"))
                    
                    '-- optional parameters --
                
                Case "s_vegtransect_by_feature"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                    .Parameters("feat") = TempVars("Feature")
                
                Case "s_vegtransect_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_tsys_datasheet_defaults"
                    '-- required parameters --
                    .Parameters("parkID") = TempVars("ParkID")
                
            End Select
            
            Set rs = .OpenRecordset(dbOpenDynaset)
            
        End With
        
    End With
    
    Set GetRecords = rs
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRecords[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     SetRecord
' Description:  Insert/update/delete record based on template
' Assumptions:  -
' Parameters:   template - SQL template name (string)
'               params - array of parameters for template (variant)
' Returns:      id - ID of record inserted, updated, deleted (long integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/21/2016 - updated i_login parameters
'   BLC - 10/24/2016 - added flag templates (contact, site, mod wentworth)
'   BLC - 10/28/2016 - updated TempVars("ContactID") -> TempVars("AppUserID"), updated i_task
' ---------------------------------
Public Function SetRecord(Template As String, Params As Variant) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim SkipRecordAction As Boolean
    Dim ID As Long
    
    'exit w/o values
    If Not IsArray(Params) Then GoTo Exit_Handler
    
    'default
    SkipRecordAction = False
            
    'default ID (if not set as param)
    ID = 0
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .sql = GetTemplate(Template)
            
            '-------------------
            ' set SQL parameters --> .Parameters("") = params()
            '-------------------
            
            '-------------------------------------------------------------------------
            ' NOTE:
            '   param(0) --> reserved for record action RefTable (ReferenceType)
            '   last param(x) --> used as record ID for updates
            '-------------------------------------------------------------------------
            Select Case Template
            
        '-----------------------
        '  INSERTS
        '-----------------------
                Case "i_comment"
                    '-- required parameters --
                    .Parameters("comtype") = Params(1)             'CommentType -> table
                    .Parameters("ctid") = Params(2)                 'TypeID
                    .Parameters("cmt") = Params(3)                  'Comment
                    .Parameters("CID") = Params(4)                  'CommentorID
                    
'                    .Parameters("CreateDate") = Now()
'                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID
'                    .Parameters("LastModified") = Now()
                    .Parameters("LMID") = TempVars("AppUserID")     'LastModifiedByID -> ContactID
        
                Case "i_contact", "i_contact_new"
                    '-- required parameters --
                    .Parameters("First") = Params(1)
                    .Parameters("Last") = Params(2)
                    .Parameters("EmailAddress") = Params(3)
                    .Parameters("Login") = Params(4)
                    .Parameters("Org") = Params(5)
                    .Parameters("MI") = Params(6)
                    .Parameters("Position") = Params(7)
                    .Parameters("Phone") = Params(8)
                    .Parameters("Ext") = Params(9)
                    .Parameters("IsActiveFlag") = Params(10)
                    
                Case "i_contact_access"
                    '-- required parameters --
                    .Parameters("ContactID") = Params(1)
                    .Parameters("AccessID") = Params(2)
                
                    'don't record the action or return ID
                    SkipRecordAction = True
                
                Case "i_cover_species"
                    'set the table name in the template --> handles WCC, URC, ARC species
                    .sql = Replace(.sql, "INTO tbl ", "INTO " & Params(0) & " ")
                                    
                    '-- required parameters --
                    .Parameters("VegPlotID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("PctCover") = Params(3)
                        
'        params(0) = "WoodyCanopySpecies"
'        params(1) = .VegPlotID
'        params(2) = .MasterPlantCode
'        params(3) = .PercentCover

'        params(0) = "RootedSpecies"
'        params(1) = .VegPlotID
'        params(2) = .MasterPlantCode
'        params(3) = .PercentCover
                
                Case "i_event"
                    '-- required parameters --
                    .Parameters("SID") = Params(1)
                    .Parameters("LID") = Params(2)
                    .Parameters("PID") = Params(3)
                    .Parameters("Start") = Params(4)
                                        
                Case "i_feature"
                    '-- required parameters --
                    .Parameters("LocationID") = Params(1)
                    .Parameters("LocationName") = Params(2)
                    .Parameters("Description") = Params(3)
                    .Parameters("Directions") = Params(4)
                
                Case "i_imported_data"
                    '-- required parameters --
                    .Parameters("idate") = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                    .Parameters("sfile") = Params(1)
                    .Parameters("dtbl") = Params(2)
                    .Parameters("nrec") = Params(3)
                    .Parameters("srec") = Params(4)
                    .Parameters("erec") = Params(5)
                    
                Case "i_location"
                    '-- required parameters --
                    .Parameters("CollectionSourceName") = Params(1)
                    .Parameters("LocationType") = Params(2)
                    .Parameters("LocationName") = Params(3)
                    .Parameters("HeadtoOrientDistance") = Params(4)
                    .Parameters("HeadtoOrientBearing") = Params(5)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                                                        
                Case "i_login"
                    '-- required parameters --
                    .Parameters("uname") = Params(1) 'username
                    .Parameters("activity") = Params(2) 'activity
                    .Parameters("version") = TempVars("AppVersion")
                    .Parameters("accesslvl") = TempVars("UserAccessLevelID")
                
                    SkipRecordAction = True
                    
                Case "i_park"
                    '-- required parameters --
                    .Parameters("ParkCode") = Params(1)
                    .Parameters("ParkName") = Params(2)
                    .Parameters("ParkState") = Params(3)
                    .Parameters("IsActiveForProtocol") = Params(4)
                                                        
                Case "i_photo"
                    '-- required parameters --
                    .Parameters("PhotoDate") = Params(1)
                    .Parameters("PhotoType") = Params(2)
                    .Parameters("PhotographerID") = Params(3)
                    .Parameters("FileName") = Params(4)
                    .Parameters("NCPNImageID") = Params(5)
                    .Parameters("DirectionFacing") = Params(6)
                    .Parameters("PhotogLocation") = Params(7)
                    .Parameters("IsCloseup") = Params(8)
                    .Parameters("IsInActive") = Params(9)
                    .Parameters("IsSkipped") = Params(10)
                    .Parameters("IsReplacement") = Params(11)
                    .Parameters("LastPhotoUpdate") = Params(12)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                
                Case "i_record_action"
                    '-- required parameters --
                    .Parameters("RefTable") = Params(0)
                    .Parameters("RefID") = Params(1)
                    .Parameters("ID") = Params(2)
                    .Parameters("Activity") = Params(3)
                    .Parameters("ActionDate") = Params(4)
                    
                    SkipRecordAction = True
                
                Case "i_site"
                    '-- required parameters --
                    .Parameters("parkID") = Params(1)
                    .Parameters("riverID") = Params(2)
                    .Parameters("code") = Params(3)         'SiteCode
                    .Parameters("sname") = Params(4)        'SiteName
                    'use |flag| to force 1/0 values vs. Access False (0) & True (-1)
                    .Parameters("flag") = Abs(Params(5))    'IsActiveForProtocol
                    
                    '-- optional parameters --
                    'NOTE: parameters are limited to 255 char
                    '      dir may be truncated via parameter since it's a MEMO field
                    .Parameters("dir") = Params(6)          'Directions
                    .Parameters("descr") = Params(7)        'Description
                
                Case "i_tagline"
                    '-- required parameters --
                    .Parameters("LineDistSource") = Params(1)
                    .Parameters("LineDistSourceID") = Params(2)
                    .Parameters("LineDistType") = Params(3)
                    .Parameters("LineDistance") = Params(4)
                    .Parameters("HeightType") = Params(5)
                    .Parameters("Height") = Params(6)
                
                Case "i_task"
                    '-- required parameters --
                    .Parameters("descr") = Params(1)         'Task
                    .Parameters("stat") = Params(2)         'Status
                    .Parameters("prio") = Params(3)         'Priority
                    .Parameters("ttype") = Params(4)        'TaskType
                    .Parameters("typeident") = Params(5)    'TaskTypeID
                    .Parameters("RID") = Params(6)          'RequestedByID
                    .Parameters("reqdate") = Params(7)      'RequestDate
                    .Parameters("CID") = Params(8)          'CompletedByID
                    .Parameters("compldate") = Params(9)    'CompleteDate
                
                    '.Parameters("CreateDate") = Now()                  'CreateDate
                    '.Parameters("CreatedByID") = TempVars("ContactID") 'CreatedByID
                    '.Parameters("LastModified") = Now()                'LastModified
                    .Parameters("LMID") = TempVars("AppUserID") 'ContactID")  'lastmodifiedID
                
                Case "i_transducer"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("TransducerType") = Params(2)
                    .Parameters("TransducerNumber") = Params(3)
                    .Parameters("SerialNumber") = Params(4)
                    .Parameters("IsSurveyed") = Params(5)
                    .Parameters("Timing") = Params(6)
                    .Parameters("ActionDate") = Params(7)
                    .Parameters("ActionTime") = Params(8)
                
                Case "i_understory_species"
                    '-- required parameters --
                    .Parameters("VegPlotID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("PercentCover") = Params(3)
                    .Parameters("IsSeedling") = Params(4)
                     
                Case "i_vegplot"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("SiteID") = Params(2)
                    .Parameters("FeatureID") = Params(3)
                    .Parameters("VegTransectID") = Params(4)
                    .Parameters("PlotNumber") = Params(5)
                    .Parameters("PlotDistance") = Params(6)
                    .Parameters("ModalSedimentSize") = Params(7)
                    .Parameters("PercentFines") = Params(8)
                    .Parameters("PercentWater") = Params(9)
                    .Parameters("UnderstoryRootedPctCover") = Params(10)
                    .Parameters("PlotDensity") = Params(11)
                    .Parameters("NoCanopyVeg") = Params(12)
                    .Parameters("NoRootedVeg") = Params(13)
                    .Parameters("HasSocialTrail") = Params(14)
                    .Parameters("FilamentousAlgae") = Params(15)
                    .Parameters("NoIndicatorSpecies") = Params(16)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                
                Case "i_vegtransect"
                    '-- required parameters --
                    .Parameters("LocationID") = Params(1)
                    .Parameters("EventID") = Params(2)
                    .Parameters("TransectNumber") = Params(3)
                    .Parameters("SampleDate") = Params(4)
        
                Case "i_vegwalk"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("CollectionPlaceID") = Params(2)
                    .Parameters("CollectionType") = Params(3)
                    .Parameters("StartDate") = Params(4)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                    
                Case "i_vegwalk_species"
                    '-- required parameters --
                    .Parameters("VegWalkID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("IsSeedling") = Params(3)
                    
                Case "i_waterway"
                    '-- required parameters --
                    .Parameters("ParkID") = Params(1)
                    .Parameters("Name") = Params(2)
                    .Parameters("Segment") = Params(3)
                    
                Case "i_usys_temp_photo"
                    '-- required parameters --
                    .Parameters("ppath") = Params(1)
                    .Parameters("pfile") = Params(2)
                    .Parameters("pdate") = Params(3)
                    .Parameters("ptype") = Params(4)
                
        '-----------------------
        '  UPDATES
        '-----------------------
                Case "u_comment"
                    '-- required parameters --
                    .Parameters("CommentType") = Params(1)
                    .Parameters("TypeID") = Params(2)
                    .Parameters("Comment") = Params(3)
                    .Parameters("CommentorID") = Params(4)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                    
                Case "u_contact"
                    '-- required parameters --
                    .Parameters("First") = Params(1)
                    .Parameters("Last") = Params(2)
                    .Parameters("EmailAddress") = Params(3)
                    .Parameters("Login") = Params(4)
                    .Parameters("Org") = Params(5)
                    .Parameters("MI") = Params(6)
                    .Parameters("Position") = Params(7)
                    .Parameters("Phone") = Params(8)
                    .Parameters("Ext") = Params(9)
                    .Parameters("IsActiveFlag") = Params(10)
                    .Parameters("ContactID") = Params(11)
                    ID = Params(11)
                
                Case "u_contact_access"
                    '-- required parameters --
                    .Parameters("ContactID") = Params(1)
                    .Parameters("AccessID") = Params(2)
                    ID = Params(1)
                
                Case "u_contact_isactive_flag"
                    '-- required parameters --
                    .Parameters("cid") = Params(1)
                    .Parameters("flag") = Params(2)
                
                Case "u_cover_species"
                    'set the table name in the template --> handles WCC, URC, ARC species
                    .sql = Replace(.sql, " tbl ", " " & Params(0) & " ")
                                    
                    '-- required parameters --
                    .Parameters("VegPlot_ID") = Params(1)
                    .Parameters("Master_PLANT_Code") = Params(2)
                    .Parameters("PctCover") = Params(3)
                
                Case "u_event"
                    '-- required parameters --
                    .Parameters("SID") = Params(1)
                    .Parameters("LID") = Params(2)
                    .Parameters("PID") = Params(3)
                    .Parameters("Start") = Params(4)
                    .Parameters("EID") = Params(5)
                    ID = Params(5)
                    
                Case "u_feature"
                    '-- required parameters --
                    .Parameters("LocationID") = Params(1)
                    .Parameters("LocationName") = Params(2)
                    .Parameters("Description") = Params(3)
                    .Parameters("Directions") = Params(4)
                    
                Case "u_location"
                    '-- required parameters --
                    .Parameters("CollectionSourceName") = Params(1)
                    .Parameters("LocationType") = Params(2)
                    .Parameters("LocationName") = Params(3)
                    .Parameters("HeadtoOrientDistance") = Params(4)
                    .Parameters("HeadtoOrientBearing") = Params(5)
                    
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                
                Case "u_mod_wentworth_retireyear"
                    '-- required parameters --
                    .Parameters("mwsid") = Params(1)
                    .Parameters("yr") = Params(2)
                
                Case "u_park"
                    '-- required parameters --
                    .Parameters("ParkCode") = Params(1)
                    .Parameters("ParkName") = Params(2)
                    .Parameters("ParkState") = Params(3)
                    .Parameters("IsActiveForProtocol") = Params(4)
                        
                Case "u_photo"
                    '-- required parameters --
                    .Parameters("PhotoDate") = Params(1)
                    .Parameters("PhotoType") = Params(2)
                    .Parameters("PhotographerID") = Params(3)
                    .Parameters("FileName") = Params(4)
                    .Parameters("NCPNImageID") = Params(5)
                    .Parameters("DirectionFacing") = Params(6)
                    .Parameters("PhotogLocation") = Params(7)
                    .Parameters("IsCloseup") = Params(8)
                    .Parameters("IsInActive") = Params(9)
                    .Parameters("IsSkipped") = Params(10)
                    .Parameters("IsReplacement") = Params(11)
                    .Parameters("LastPhotoUpdate") = Params(12)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                
                Case "u_site"
                    '-- required parameters --
                    .Parameters("ParkID") = Params(1)
                    .Parameters("RiverID") = Params(2)
                    .Parameters("Code") = Params(3)
                    .Parameters("Name") = Params(4)
                    .Parameters("IsActiveForProtocol") = Params(5)
                    
                    '-- optional parameters --
                    .Parameters("Directions") = Params(6)
                    .Parameters("Description") = Params(7)
                
                Case "u_site_isactive_flag"
                    '-- required parameters --
                    .Parameters("sid") = Params(1)
                    .Parameters("flag") = Params(2)
                
                Case "u_tagline"
                    '-- required parameters --
                    .Parameters("LineDistSource") = Params(1)
                    .Parameters("LineDistSourceID") = Params(2)
                    .Parameters("LineDistType") = Params(3)
                    .Parameters("LineDistance") = Params(4)
                    .Parameters("HeightType") = Params(5)
                    .Parameters("Height") = Params(6)
                
                Case "u_task"
                    '-- required parameters --
                    .Parameters("Task") = Params(1)
                    .Parameters("Status") = Params(2)
                    .Parameters("RequestedByID") = Params(3)
                    .Parameters("RequestDate") = Params(4)
                    .Parameters("CompletedByID") = Params(5)
                    .Parameters("CompleteDate") = Params(6)
                
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                
                Case "u_transducer"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("TransducerType") = Params(2)
                    .Parameters("TransducerNumber") = Params(3)
                    .Parameters("SerialNumber") = Params(4)
                    .Parameters("IsSurveyed") = Params(5)
                    .Parameters("Timing") = Params(6)
                    .Parameters("ActionDate") = Params(7)
                    .Parameters("ActionTime") = Params(8)
                
                Case "u_template"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                
                Case "u_tsys_datasheet_defaults"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                    .Parameters("pid") = Params(2)
                    .Parameters("rid") = Params(3)
                    .Parameters("cover") = Params(4)
                    .Parameters("species") = Params(5)
                    .Parameters("blanks") = Params(6)
                    
                    '-- optional parameters --
                
                Case "u_usys_temp_photo"
                    '-- required parameters --
                    .Parameters("iid") = Params(1)
                    .Parameters("ptype") = Params(4)
                
                Case "u_vegtransect"
                    '-- required parameters --
                    .Parameters("LocationID") = Params(1)
                    .Parameters("EventID") = Params(2)
                    .Parameters("TransectNumber") = Params(3)
                    .Parameters("SampleDate") = Params(4)
                
                Case "u_vegwalk"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("CollectionPlaceID") = Params(2)
                    .Parameters("CollectionType") = Params(3)
                    .Parameters("StartDate") = Params(4)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") '"ContactID")
                
                Case "u_vegwalk_species"
                    '-- required parameters --
                    .Parameters("VegWalkID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("IsSeedling") = Params(3)
                    
                Case "u_understory_species"
                    '-- required parameters --
                    .Parameters("VegPlotID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("PercentCover") = Params(3)
                    .Parameters("IsSeedling") = Params(4)
                
                Case "u_waterway"
                    '-- required parameters --
                    .Parameters("ParkID") = Params(1)
                    .Parameters("Name") = Params(2)
                    .Parameters("Segment") = Params(3)
                
            End Select
            
            .Execute dbFailOnError
                
    ' -------------------
    '  Record Action
    ' -------------------
            'handle unrecorded actions & those which don't generate an ID
            If SkipRecordAction Then GoTo Exit_Handler
            
            If ID = 0 Then
                'retrieve identity
                ID = db.OpenRecordset("SELECT @@IDENTITY;")(0)
            End If
            
            'set record action
            .sql = GetTemplate("i_record_action")
                                            
            '-- required parameters --
            .Parameters("RefTable") = Params(0)
            .Parameters("RefID") = ID
            .Parameters("ID") = TempVars("AppUserID") 'TempVars("ContactID")
            .Parameters("Activity") = "DE"
            .Parameters("ActionDate") = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                                
            .Execute dbFailOnError
            
            'cleanup
            .Close
        
        End With

        SetRecord = ID
    End With
                
Exit_Handler:
    'cleanup
    Set qdf = Nothing
    Set db = Nothing

    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     UpsertRecord
' Description:  Handle insert/update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   gecko_1, February 10, 2005
'   http://www.access-programmers.co.uk/forums/showthread.php?t=81221
'   Khinsu, August 19, 2013
'   http://stackoverflow.com/questions/18317059/how-to-test-if-item-exists-in-recordset
'   HansUp, April 4, 2013
'   http://stackoverflow.com/questions/15823687/findfirst-vba-access2010-unbound-form-runtime-error
' Source/date:  Bonnie Campbell, July 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/28/2016 - initial version
'   BLC - 9/1/2016  - added vegwalk, photo
'   BLC - 10/4/2016 - added template, adjusted for form w/o list
'   BLC - 10/14/2016 - updated to accommodate non-users for contacts
'   BLC - 1/9/2017 - revised retrieve ID from ContactID to ID, revised i_event to use TempVar("SiteID")
' ---------------------------------
Public Sub UpsertRecord(ByRef frm As Form)
On Error GoTo Err_Handler
    
' ----------------------------------------------------------------------------------
'    1) Click to edit
'       a) populates form fields
'       b) tbxID is set
'
'       c) change values --> i) compare against existing values
'                           ii) no existing values match ==> update
'                           iii) existing values match ==> message no change
'
'   2) Enter new values
'       a) enables save button
'       b) click save -->   i) compare against existing values
'                           ii) no existing values match ==> insert
'                           iii) existing values match ==> message no change
' ----------------------------------------------------------------------------------
    
    Dim DoAction As String, strCriteria As String, strTable As String
    Dim NoList As Boolean
    Dim obj As Object
    
    'use generic object to handle multiple obj types
    With obj
    
        'default
        NoList = False
        strTable = frm.Name
    
        Select Case frm.Name
            Case "Contact"
                Dim p As New Person
    
                With p
                    'values passed into form
                            
                    'form values
                    .LastName = frm!tbxLast.Value
                    .FirstName = frm!tbxFirst.Value
                    If Not IsNull(frm!tbxMI.Value) Then p.MiddleInitial = frm!tbxMI.Value  'FIX EMPTY STRING
                    .Email = frm!tbxEmail.Value
                    '.Username = frm!tbxUsername.Value
                    If Not IsNull(frm!tbxUsername.Value) Then p.Username = frm!tbxUsername.Value
                    If Not IsNull(frm!tbxOrganization.Value) Then p.Organization = frm!tbxOrganization.Value
                    If Not IsNull(frm!tbxPosition.Value) Then .PosTitle = frm!tbxPosition.Value
                    If Not IsNull(frm!tbxPhone.Value) And Len(frm!tbxPhone.Value) > 0 Then
                        .WorkPhone = RemoveChars(frm!tbxPhone.Value, True) 'remove non-numerics
                    Else
                        .WorkPhone = Null
                    End If
                    If Not IsNull(frm!tbxExtension.Value) And Len(frm!tbxExtension.Value) > 0 Then
                        .WorkExtension = RemoveChars(frm!tbxExtension.Value, True) 'remove non-numerics
                    Else
                        .WorkExtension = Null
                    End If
                    If Not IsNull(frm!cbxUserRole.Column(1)) Then .AccessRole = frm!cbxUserRole.Column(1)
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[FirstName] = '" & .FirstName _
                                    & "' AND [LastName] = '" & .LastName _
                                    & "' AND [MiddleInitial] = '" & .MiddleInitial _
                                    & "' AND [Email] = '" & .Email & "'"
                    
                    'set the generic object --> Contact
                    Set obj = p
                    
                    'cleanup
                    Set p = Nothing
                End With

            Case "Events"
                Dim ev As New EventVisit
                strTable = "Event"
                
                With ev
                    'values passed into form
                    
                    'form values
                    .LocationID = frm!cbxLocation.Column(0)
                    .ProtocolID = 1 ' assumes this is for big rivers protocol
                    .SiteID = TempVars("SiteID") 'frm!cbxSite.Column(0)
                    
                    .StartDate = frm!tbxStartDate.Value
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    
                    strCriteria = "[Site_ID] = " & .SiteID & " AND [Location_ID] = " & .LocationID & " AND [StartDate] = " & Format(.StartDate, "YYYY-mm-dd")
                    
                    'set the generic object --> EventVisit
                    Set obj = ev
                    
                    'cleanup
                    Set ev = Nothing
                End With
            
            Case "Feature"
                    Dim f As New Feature
    
                    With f
                        'values passed into form
                                
                        'form values
                        .LocationID = frm!cbxLocation.Column(0)
                        .Name = frm!tbxFeature.Value
                        
                        If Not IsNull(frm!tbxFeatureDirections.Value) Then f.Directions = frm!tbxFeatureDirections.Value
                        If Not IsNull(frm!tbxDescription.Value) Then .Directions = frm!tbxDescription.Value
                        .ID = frm!tbxID.Value '0 if new, edit if > 0
                    
                        strCriteria = "[Location_ID] = " & .LocationID & " AND [Feature] = '" & .Name & "'"
                        
                        'set the generic object --> Feature
                        Set obj = f
                    
                        'cleanup
                        Set f = Nothing
                    End With

            Case "Location"
                    Dim loc As New Location
                    
                    With loc
                        'values passed into form
                        Dim loctype As String
                        
                        Select Case frm.CallingForm
                            Case "Feature"
                                loctype = "F"
                            Case "VegPlot"
                                loctype = "P"
                            Case "Transect"
                                loctype = "T"
                        End Select
                        
                        'CollectionSourceName is the ID from passed in form
                        'collection feature ID (A, B, C...) or Transect number (1-8)
                        .CollectionSourceName = frm.CallingRecordID '"T"
                        
'                        .CreateDate = ""
'                        .CreatedByID = 0
                        .LastModified = Now()
                        .LastModifiedByID = 0
                        
                        '.ProtocolID = 1
                        '.SiteID = 1
                        
                        'form values
                        .LocationName = frm!tbxName.Value 'Collection feature ID (A, B, C, ...) or Transect number (1-8)
                        .LocationType = loctype '"T" 'cbxLocationType.SelText 'F- feature, T- transect, P - plot
                
                        .HeadtoOrientDistance = frm!tbxDistance.Value
                        .HeadtoOrientBearing = frm!tbxBearing.Value
                        
                        .ID = frm!tbxID.Value '0 if new, edit if > 0

                        strCriteria = "[LocationName] = '" & .LocationName _
                                    & "' AND [LocationType] = '" & .LocationType _
                                    & "' AND [CollectionSourceName] = '" & .CollectionSourceName _
                                    & "' AND [HeadtoOrientDistance_m] = " & .HeadtoOrientDistance _
                                    & " AND [HeadtoOrientBearing] = " & .HeadtoOrientBearing '_
'                                    & " AND [LastModified] = " & .LastModified _
'                                    & " AND [LastModifiedBy_ID] = " & .LastModifiedByID
                    
                        'set the generic object --> Location
                        Set obj = loc
                        
                        'cleanup
                        Set loc = Nothing
                    End With
                                        
            Case "Photo"
                Dim ph As New Photo
                
                With ph
                    'values passed into form
                
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                                
                    'set the generic object --> Location
                    Set obj = p
                    
                    'cleanup
                    Set ph = Nothing
                End With
                                        
            Case "SetObserverRecorder"
                Dim ra As New RecordAction
                
                With ra
                    'values passed into form
                    .RefTable = frm.RefTable
                    .RefID = frm.RefID
                    .ContactID = frm.RAContactID
                    .RefAction = frm.RAAction
                    '.ActionType = frm.RAAction
                    .ActionDate = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                                
                    strCriteria = "[Contact_ID] = " & .ContactID _
                                & " AND [Activity] = '" & .RefAction _
                                & "'"

                    'set the generic object --> Location
                    Set obj = ra
                    
                    'cleanup
                    Set ra = Nothing
                End With
            
            
            Case "Site"
                Dim s As New Site
                
                With s
                    'values passed into form
                    .Park = TempVars("ParkCode")
                    .River = TempVars("River")
                    
                    'form values
                    .Code = frm!tbxSiteCode.Value
                    .Name = frm!tbxSiteName.Value
                    .Directions = Nz(frm!tbxSiteDirections.Value, "")
                    .Description = Nz(frm!tbxDescription.Value, "")
                    
                    'assumed
                    .IsActiveForProtocol = 1 'all sites assumed active when added
        
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[SiteCode] = '" & .Code & "' AND [SiteName] = '" & .Name & "'"
                
                    'set the generic object --> Site
                    Set obj = s
                    
                    'cleanup
                    Set s = Nothing
                End With
                
            Case "Template"
                Dim tpl As New Template
                
                With tpl
                    .IsSupported = 1
                    .Context = ""
                    .EffectiveDate = Date
                    .Remarks = ""
                    .TemplateName = ""
                    .Version = ""
                    .TemplateSQL = ""
                    .Syntax = ""
    
                End With
                
                'set the generic object --> Template
                Set obj = tpl
                
                'cleanup
                Set tpl = Nothing
            
            Case "Task"
                Dim tk As New Task
                
                With tk
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    .RequestDate = frm!tbxRequestDate.Value
                    .RequestedByID = frm!cbxRequestedBy.Column(0)
                    .Status = frm!cbxStatus.Column(0)
                    .Priority = frm!cbxPriority.Column(0)
                    .Task = frm!tbxTask.Value
                    .TaskType = frm.ContextType
                    
                    
'                    strCriteria = "[TransducerNumber] = " & .TransducerNumber _
'                                & " AND [Timing] = '" & .Timing _
'                                & "' AND [SerialNumber] = '" & .SerialNumber _
'                                & "' AND [ActionDate] = " & .ActionDate
                
                    'set the generic object --> Transducer
                    Set obj = tk
                    
                    'cleanup
                    Set tk = Nothing
                End With
                
            Case "TemplateAdd"
                'Dim tpl As New Template
                
                With tpl
                    .TemplateName = frm!tbxTemplate
                    .Context = .TemplateName
                    .IsSupported = 1
                    .Version = frm!tbxVersion
                    .Syntax = frm!cbxSyntax
                    .TemplateSQL = frm!tbxTemplateSQL
                    .EffectiveDate = frm!tbxEffectiveDate
                    .Params = GetParamsFromSQL(.TemplateSQL)
                    .Remarks = frm!tbxRemarks
                    .ContactID = TempVars("AppUserID")
                    
                    'set the generic object --> Transducer
                    Set obj = tpl
                    
                    'cleanup
                    Set tpl = Nothing
                End With
                
                'inserts only, no ID?
                NoList = True
                
            Case "Transducer"
                Dim t As New Transducer
        
                With t
                    'values passed into form
                    .EventID = 1
                            
                    'form values
                    .TransducerType = ""
                    .TransducerNumber = frm!cbxTransducer.SelText
                    .SerialNumber = frm!tbxSerialNo.Value
                    .IsSurveyed = frm!chkSurveyed.Value
                    .Timing = frm!cbxTiming.SelText
                    .ActionDate = Format(frm!tbxSampleDate.Value, "YYYY-mm-dd")
                    .ActionTime = Format(frm!tbxSampleTime.Value, "hh:mm.ss")
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[TransducerNumber] = " & .TransducerNumber _
                                & " AND [Timing] = '" & .Timing _
                                & "' AND [SerialNumber] = '" & .SerialNumber _
                                & "' AND [ActionDate] = " & .ActionDate
                
                    'set the generic object --> Transducer
                    Set obj = t
                    
                    'cleanup
                    Set t = Nothing
                End With
            
            Case "Transect"
                Dim vt As New VegTransect
                strTable = "VegTransect"
                
                With vt
                    'values passed into form
                    .Park = TempVars("ParkCode")
                    .LocationID = 1
                    .EventID = 1
                            
                    'form values
                    .TransectNumber = frm!tbxNumber.Value
                    .SampleDate = Format(frm!tbxSampleDate.Value, "YYYY-mm-dd")
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    
                    strCriteria = "[TransectNumber] = " & .TransectNumber _
                                & "' AND [SampleDate] = " & .SampleDate
                
                    'set the generic object --> VegTransect
                    Set obj = vt
                    
                    'cleanup
                    Set vt = Nothing
                End With
            
            Case "UserRole"
                Dim u As New Person
                    
                With u
                    'values passed into form
            '        .EventID = 1
                            
                    'form values
            '        .UserRoleType = ""
            '        .UserRoleNumber = cbxUserRole.SelText
            '        .SerialNumber = tbxSerialNo.value
            '        .IsSurveyed = chkSurveyed.value
            '        .Timing = cbxTiming.SelText
            '        .ActionDate = Format(tbxSampleDate.value, "YYYY-mm-dd")
            '        .ActionTime = Format(tbxSampleTime.value, "hh:mm.ss")
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    'strCriteria = "[UserRoleNumber] = " & .UserRoleNumber
                
                    'set the generic object --> Location
                    Set obj = u
                    
                    'cleanup
                    Set u = Nothing
                End With

            Case "VegWalk"
                Select Case frm.FormContext
                    Case "AllRootedSpecies"
                        Dim ars As New RootedSpecies
                        
                        With ars
                            'values passed into form
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            'set the generic object --> Woody Canopy Species
                            Set obj = ars
                            
                            'cleanup
                            Set ars = Nothing
                        End With
                    
                    Case "UnderstoryRootedSpecies"
                        Dim ucs As New UnderstoryCoverSpecies
                        
                        With ucs
                            'values passed into form
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            'set the generic object --> Woody Canopy Species
                            Set obj = ucs
                            
                            'cleanup
                            Set ucs = Nothing
                        End With

                    Case "VegWalk"
                        Dim vw As New VegWalk
                        
                        With vw
                            'values passed into form
                        
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                                        
                            'set the generic object --> Location
                            Set obj = vw
                            
                            'cleanup
                            Set vw = Nothing
                        End With
                    
                    Case "WoodyCanopySpecies"
                        Dim wcs As New WoodyCanopySpecies
                        
                        With wcs
                            'values passed into form
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            'set the generic object --> Woody Canopy Species
                            Set obj = wcs
                            
                            'cleanup
                            Set wcs = Nothing
                        End With
                        
                End Select
            

            Case Else
                GoTo Exit_Handler
        End Select
                
        'set insert/update based on whether its an edit or new entry
        DoAction = IIf(frm!tbxID.Value > 0, "u", "i")
        
        If NoList Then
                    
            'form doesn't contain list subform or message/icon fields
            'so cut to the chase -> do nothing here
            
        Else
        
            'check if the record already exists by checking event list form records
            'event list form pulls active records for park, river segment
            Dim rs As DAO.Recordset
            
            Set rs = frm!list.Form.RecordsetClone
            rs.FindFirst strCriteria
            
            If rs.NoMatch Then
                ' --- INSERT ---
                frm!lblMsg.ForeColor = lngLime
                frm!lblMsgIcon.ForeColor = lngLime
                frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                frm!lblMsg.Caption = IIf(DoAction = "i", "Inserting new record...", "Updating record...")
            Else
                ' --- UPDATE ---
                'record already exists & ID > 0
                
                'retrieve ID
                If frm!tbxID.Value = rs("ID") Then 'rs("Contact.ID") Then
                    'IDs are equivalent, just change the data
                    frm!lblMsg.ForeColor = lngLime
                    frm!lblMsgIcon.ForeColor = lngLime
                    frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                    frm!lblMsg.Caption = "Updating record..."
                Else
                    'prevent duplicate record entries
                    frm!lblMsg.ForeColor = lngYellow
                    frm!lblMsgIcon.ForeColor = lngYellow
                    frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                    frm!lblMsg.Caption = "Oops, record already exists."
                    GoTo Exit_Handler
                End If
                
            End If
        End If
        
        'T/F refers to whether the record is an update (T) or insert (F)
        obj.SaveToDb IIf(DoAction = "i", False, True)
        
        'add the action record --> DONE via SaveToDb (thru SetRecord)
        
        'set the tbxID.value ==> tbxID is a bound control, can't set it this way
        'tbxID = .ID
        'frm!tbxID.Value = obj.ID
        'frm.Controls("tbxID").Value = obj.ID
    End With
    
    'clear values & refresh display
    frm.ReadyForSave 'Application defined error? --> ensure ReadyForSave is Public Sub
    'Forms!frm.ReadyForSave
    
    'handle situations where Access is saving same record
    
    'save record changes from form first to avoid "Write Conflict" errors
    'where form & SQL are attempting to save record
    'frm.Dirty = False
    
    If frm.Dirty Then
        Debug.Print "UpsertRecord " & frm.Name & " DIRTY"
        'frm.Dirty = False
        
        frm!lblMsg.ForeColor = lngYellow
        frm!lblMsgIcon.ForeColor = lngYellow
        frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
        frm!lblMsg.Caption = "** DIRTY **" 'UNSAVED CHANGES! **"
        
    Else
        Debug.Print "UpsertRecord " & frm.Name & " CLEAN"
    End If
        
' CHECK IF POPULATING FORM IS THE ISSUE...
'    PopulateForm frm, frm!tbxID.Value
    
'    'refresh list
'    frm!list.Requery
    
    frm.Requery
    
    'clear messages & icon
    frm!lblMsgIcon.Caption = ""
    frm!lblMsg.Caption = ""
    
    'refresh list
    frm!list.Requery
    
    'exit
    GoTo Exit_Handler
    
Form_Without_List:
    DoAction = "i"
    Resume Next

Exit_Handler:
    'cleanup
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpsertRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SetObserverRecorder
' Description:  Sets data observer & recorder
' Assumptions:  -
' Parameters:   obj - object to set observer/recorder on (object)
'               tbl - name of table being modified (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 9, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/9/2016 - initial version
' ---------------------------------
Public Sub SetObserverRecorder(obj As Object, tbl As String)
On Error GoTo Err_Handler
    
    'handle record actions
    Dim act As New RecordAction
    With act
    
    'Recorder
        .RefAction = "R"
        .ContactID = obj.RecorderID
        .RefID = obj.ID
        .RefTable = tbl
        .SaveToDb
        
    'Observer
        .RefAction = "O"
        .ContactID = obj.ObserverID
        .RefID = obj.ID
        .RefTable = tbl
        .SaveToDb
        
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetObserverRecorder[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          UploadCSVFile
' Description:  Uploads data into database from CSV file
' Assumptions:  -
' Parameters:   strFilename - name of file being uploaded (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/1/2016 - initial version
'   BLC - 10/19/2016 - renamed to UploadCSVFile from UploadSurveyFile to genericize
' ---------------------------------
Public Sub UploadCSVFile(strFilename As String)
On Error GoTo Err_Handler
    
    'import to table
    ImportCSV strFilename, "usys_temp_csv", True, True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UploadCSVFile[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' Function:          FetchAddlData
' Description:  Retrieves additional data field(s)
' Assumptions:
'               fields are delimited w/ a pipe (|)
' Parameters:   tbl - name of table to retrieve from (string)
'               field(s) - name of field to retrieve (string)
'               id - record to retrieve's ID (long)
' Returns:      field value(s) for record (DAO.Recordset)
' Throws:       none
' References:
'   Steven Thomas, November 28, 2011
'   https://blogs.office.com/2011/11/28/display-real-time-information-with-the-controltip-property/
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Public Function FetchAddlData(tbl As String, Fields As String, ID As Long) As DAO.Recordset
On Error GoTo Err_Handler
    
    'values are required --> exit if not
    If Len(tbl) = 0 Or Len(Fields) = 0 Or Not (ID > 0) Then GoTo Exit_Handler
    
    'begin retrieval
    Dim field As String
    Dim strFields As String
    Dim strSQL As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
            
            'check for multiple fields
            If InStr(Fields, "|") > 0 Then
                Dim aryFlds() As String
                Dim i As Integer
                
                aryFlds = Split(Fields, "|")
                
                For i = 0 To UBound(aryFlds)
                    strFields = aryFlds(i) & ","
                Next
                
                'remove extra comma
                strFields = IIf(Right(strFields, 1) = ",", RTrim(strFields), strFields)
            
            Else
                
                strFields = Fields
            End If
            
            'base
            strSQL = "SELECT " & strFields & " FROM " & tbl & " WHERE ID = " & ID & ";"
            
            'update the query SQL
            .sql = strSQL
            
            Dim rs As DAO.Recordset

            Set rs = .OpenRecordset
                        
            'send results
            Set FetchAddlData = rs
            
            'cleanup
            Set rs = Nothing
            Set qdf = Nothing
            Set db = Nothing

        End With
    End With
    

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FetchAddlData[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          GetHierarchyLevel
' Description:  Determine the hierarchy level set
' Assumptions:  -
' Parameters:   -
' Returns:      lvl - maximum level set in the application (string)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 1, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2017 - initial version
' ---------------------------------
Public Function GetHierarchyLevel() As String
On Error GoTo Err_Handler
    
    Dim lvl As String
    
    'default
    lvl = ""
    
    If Not TempVars("ParkCode") Is Nothing Then
        lvl = "park"
        If Not TempVars("River") Is Nothing Then
            lvl = "river"
            If Not TempVars("SiteCode") Is Nothing Then
                lvl = "site"
                If Not TempVars("Feature") Is Nothing Then
                    lvl = "feature"
                End If
            End If
        End If
    End If

    GetHierarchyLevel = lvl

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetHierarchyLevel[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function