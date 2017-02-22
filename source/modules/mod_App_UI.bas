Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.21
' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added PopulateSpeciesPriorities function from mod_Species
'               BLC, 11/19/2015 - 1.02 - added CreateEnums call to initApp
'               BLC, 4/26/2016  - 1.03 - added ClickAction() for handling various app actions
'               BLC, 6/24/2016 - 1.04 - replaced Exit_Function > Exit_Handler
'               BLC, 7/5/2016  - 1.05 - added ClearFields() to support Species Search
'               BLC, 8/8/2016 - 1.06 - revised to use default table name in PopulateForm()
'               BLC, 8/29/2016 - 1.07 - revised to use usys_temp_qdf & Contact_ID in PopulateForm()
'                                       for Contact form
'               BLC, 8/30/2016 - 1.08 - added Batch Upload Photos to ClickAction()
'               BLC, 9/13/2016 - 1.09 - added SortList()
'               BLC, 10/14/2016 - 1.10 - added SetContext()
'               BLC, 10/19/2016 - 1.11 - revised to use UploadCSVFile() vs. UploadSurveyFile()
'               BLC, 10/24/2016 - 1.12 - added modwentworth form
'               BLC, 10/25/2016 - 1.13 - added originForm TempVar for species seach
'               BLC, 12/9/2016 -  1.14 - added PopulateCSVFields()
'               BLC, 12/13/2016 - 1.15 - added SetCurrentPseudoRecord()
'               BLC, 1/9/2017   - 1.16 - revised ClickAction() to use SetTempVar()
'               BLC, 1/12/2017 - 1.17 - revised to VegTransect vs. Transect form
'               BLC, 1/31/2017 - 1.18 - adjusted SortListForm() to accommodate template list form
'               BLC, 2/2/2017  - 1.19 - commented CreateEnums call in Initialize(),
'                                       most/not all enums handled through calls to AppEnum table
'               BLC, 2/14/2017 - 1.20 - added Task form to PopulateForm()
'               BLC, 2/21/2017 - 1.21 - adjusted SortListForm() to accommodate Contact list form,
'                                       revised to use Photo vs. Tree form
' =================================

' =================================
'   Functions
' =================================


' =================================
'   Methods
' =================================

' =================================
' SUB:     PopulateInsetTitle
' Description:  Sets inset title on form
' Assumptions:
' Parameters:   frm - form holding crumb labels
'               aryCrumbs - breadcrumb array
'               separator - non-clickable value between crumbs, default = >
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - initial version
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
'               BLC, 5/26/2015 - Added error handling
' =================================
Public Sub PopulateInsetTitle(ctrl As Control, strContext As String)
On Error GoTo Err_Handler
    
    Dim strTitle As String
    
    Select Case strContext
        Case "Create" ' Create main
            strTitle = "Choose what you'd like to create"
        Case "CreateTgtLists" ' Create species target lists
            strTitle = "Create > Species Target Lists"
        Case "AddTgtArea" ' Add target areas
            strTitle = "Create > Add Target Area"
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates"  ' QA/QC > Outliers etc.
            strContext = Replace(Replace(strContext, "Suspect", "Suspect "), "Missing", "Missing ")
            strTitle = "Data Validation > " & strContext
        Case "Data Validation" ' QA/QC analysis project selection
            strTitle = "Data Validation > Field > Duplicates (NFV)" '<<<<< Make this so it ties back to the selected analysis
        Case "View" ' View main
            strTitle = "View"
        Case "Reports" ' Reports main
            strTitle = "Reports"
        Case "CrewVegWalk" ' Reports > Field Crew Species List
            strTitle = "Reports > Field Crew Species List"
        Case "VegWalkByPark" ' Reports > Species List By Park
            strTitle = "Reports > Species List By Park"
        Case "TgtListAnnualSummary" ' Reports > Annual Species List Summary
            strTitle = "Reports > Annual Species List Summary"
        Case "Precision", "Effectiveness", "Bias", "Stage", "Flow" ' Reports > Precision etc.
            strTitle = "Reports > " & strContext
        Case "Export" ' Export main
            strTitle = "Export"
        Case "UtahLab" ' Exports > Utah Lab etc.
            strContext = Replace(strContext, "Lab", " Lab")
            strTitle = "Exports > " & strContext
        Case "DB Admin" ' DB Admin main
            strTitle = ""
    End Select
    
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strTitle
        If strContext <> "DbAdmin" Then
            ctrl.Visible = True
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateInsetTitle[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' SUB:     PopulateInstructions
' Description:  Sets form instruction strings
' Assumptions:  -
' Parameters:   strTab - tab for instruction string
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - initial version
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
'               BLC, 5/26/2015 - Added error handling
' =================================
Public Sub PopulateInstructions(ctrl As Control, strContext As String)
On Error GoTo Err_Handler
    Dim strInstructions As String
    
    'MsgBox strContext
    
    Select Case strContext
        Case "Create" ' Create main
            strInstructions = "Choose what you would like to create."
        Case "CreateTgtLists" ' Create > Species Target Lists
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your list."
        Case "AddTgtArea" ' Create > Add Target Area
            strInstructions = "" '"Choose the park and year for your target area. Click 'Continue' to create your area."
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates" ' QA/QC main
            strInstructions = "Complete the fields to define the data set or subset you are validating. " _
                    & "Leave the fields blank if you are validating all data. Click 'Run' to validate."
        Case "View" ' View main
            strInstructions = "The view menu is currently not in use for this application."
            'strInstructions = "Log your modifications to data within the edit log. " _
            '        & "Be as complete as possible to aid others in tracing data changes."
        Case "Reports" ' Reports main
            strInstructions = "Choose the report you would like to run."
        Case "CrewVegWalk" ' Reports > Field Crew Species List
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "VegWalkByPark" ' Reports > Species List By Park
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "TgtListAnnualSummary"
            strInstructions = "Choose the year(s) for your list. Click 'Continue' to prepare your report." & vbCrLf & vbCrLf & _
                            "This report may take a minute to create and display. " & vbCrLf & _
                            "Calculated summary values will display once the report has finished rendering. " & vbCrLf & vbCrLf & _
                            "Your patience is appreciated."
        Case "Precision", "Effectiveness", "Bias", "Stage", "Flow" ' Reports > Precision etc.
            strInstructions = "Complete the fields to define the data set or subset you are reporting. " _
                    & "Leave the fields blank if you are reporting on all data. Click 'Run' to validate."
        Case "Export" ' Export main
            strInstructions = "After opening a report from the report tab, use the Export menu above in the application menu to export reports to your desired format."
        Case "UtahLab" ' Exports > Utah Lab etc.
            strInstructions = "Choose the export you would like to run."
        Case "DbAdmin" ' DB Admin main
            strInstructions = "The database administration tab is currently not in use for this application."
            'strInstructions = ""
    End Select
    
    'populate caption & display instructions
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strInstructions
        If strContext <> "DbAdmin" Then
            ctrl.Visible = True
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateInstructions[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     PopulateSpeciesPriorities
' Description:  Populate species priority values from species priority concatenation
' Assumptions:  Park priority textboxes are named tbxPARKPriority (e.g. tbxZIONPriority)
' Parameters:   parkCode - 4 character park code (string)
'               priorities - species priority string concatenation for all parks (e.g. "BLCA-1|COLM-Transect|FOBU-1")
' Returns:      Priority - value for park species priority (string)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/9/2015 - initial version
'   BLC - 5/26/2015 - moved from mod_Species to mod_App_UI
' ---------------------------------
Public Function PopulateSpeciesPriorities(ParkCode As String, priorities As String) As String

On Error GoTo Err_Handler

Dim ParkPriorities As Variant
Dim i As Integer

    'check if parkCode is in priorities string
    If Len(priorities) > Len(Replace(priorities, ParkCode, "")) Then
    
        'prepare the Park Priority values
        ParkPriorities = Split(priorities, "|")
        
        'set park priority values
        For i = 0 To UBound(ParkPriorities)
            'does Park have a priority value?
            If ParkCode = Left(ParkPriorities(i), 4) Then
                PopulateSpeciesPriorities = Replace(ParkPriorities(i), ParkCode + "-", "")
            End If
        Next
        
    Else
        'not listed
        PopulateSpeciesPriorities = "X"
    
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateSpeciesPriorities[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          Initialize
' Description:  initialize application values
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/19/2015 - added dynamic getParkState() & standard error handling
'   BLC - 3/4/2015  - shifted colors to mod_Color, removed setting of park, state, tgtYear TempVars
'   BLC - 5/13/2015 - stub only
'   BLC - 11/19/2015 - added CreateEnums call to create application specific Enums,
'                      updated documentation to reflect mod_App_UI vs. mod_Init
'   BLC - 2/2/2017  - comment: most enums handled through calls to AppEnum table
'                     however some require CreateEnums()
' ---------------------------------
Public Sub Initialize()
On Error GoTo Err_Handler

    'create the enums specific to this application from the Enums table &
    'mod_App_Enum stub module
    'CreateEnums requires BOTH mod_Enum & mod_App_Enum files to be re-imported to
    'the database
    CreateEnums

    'set application UI display
'     SetStartupOptions "AppTitle", dbText, "NCPN Big Rivers"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Initialize[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ================================ Big Rivers ===========================
' ---------------------------------
' SUB:          SetHints
' Description:  Sets hints for form actions
' Assumptions:  -
' Parameters:   frm - form object to set hints on (form)
'               strForm - form name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/27/2015  - initial version
' ---------------------------------
Public Sub SetHints(frm As Form, strForm As String)
On Error GoTo Err_Handler

' Forms!Mainform!Subform1.Form!
 
    With frm!fsub.Form
 
        Select Case strForm
 
            Case "fsub_Photo_FTOR_Details"
 
                !lblCloseupHint.Caption = "Is the photo a closeup?"
                !lblReplacementHint.Caption = "Does photo replace another?"
                !lblCommentHint.Caption = ""
 
                Select Case TempVars("phototype")
                    Case "R" 'reference
                        !lblPhotogLocHint.Caption = "from river, 10m upstream, etc."
                        !lblSubjectLocHint.Caption = "CP1, RM2, etc."
                    Case "O" 'overview
                        !lblPhotogLocHint.Caption = ""
                        !lblSubjectLocHint.Caption = "O1, O2, etc."
                    Case "T" 'transect
                        !lblPhotogLocHint.Caption = "T + transect# - order# (T2-1)"
                        !lblSubjectLocHint.Caption = ""
                    Case "F" 'feature
                        !lblPhotogLocHint.Caption = "F + transect# - order# " & vbCrLf & "(F3/4-2)"
                        !lblSubjectLocHint.Caption = ""
                End Select
 
            Case "fsub_Photo_Other_Details"
                !lblDescriptionHint.Caption = ""
            Case Else
 
        End Select
 
        !lblPhotoNumHint.Caption = "P + Month" & vbCrLf & "(Jan-Sep=0-9,Oct-Dec=A-C) + day(01-31) + " & vbCrLf & "4-digit camera seq# " & vbCrLf & "(PA010300 = Jan 1, #300)"
                  
      End With
      
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHints[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     SetCurrentPseudoRecord
' Description:  sets a pseudo current record # based on the combobox w/ current focus
' Assumptions:  -
' Parameters:   -
' Returns:      current # for combobox (integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/13/2016 - initial version
' ---------------------------------
Public Function SetCurrentPseudoRecord(ctrl As Control) As Integer
On Error GoTo Err_Handler
              
    'set psuedo current record <- set the # of the cbx
'    'MsgBox ActiveControl.Name
'    MsgBox Screen.ActiveForm.Name & " is the active form."
'    MsgBox Screen.ActiveControl.Name & "is the active control."
'    If InStr(Me.ActiveControl, "cbxColumnName") Then
'        Me.Parent.Controls("tbxCSVRecord").Value = Replace(Me.ActiveControl.Name, "cbxColumnName", "")
'    End If

    If InStr(ctrl.Name, "cbxColumnName") Then
        ctrl.Parent.Form.Parent.Form.Controls("tbxCSVRecord").Value = Replace(ctrl.Name, "cbxColumnName", "")
        ChangeBackColor ctrl, lngYelLime
        Call Forms("ImportMap").tbxCSVRecord_Change
    End If
              
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetCurrentPseudoRecord[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ClickAction
' Description:  Handles click events for various form links
' Assumptions:  Link caption and tag text matches action text values.
'               If a link caption &/or tag changes, the corresponding action must change
'               here too.
' Parameters:   action - concatenated link label caption & tag (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/26/2016  - initial version
'   BLC - 8/30/2016  - added Batch Upload Photos
'   BLC - 10/19/2016 - revised to use UploadCSVFile() vs. UploadSurveyFile()
'   BLC - 10/25/2016 - revised species search to add originForm TempVar, callingform oArg
'   BLC - 1/9/2017   - revised to use SetTempVar()
'   BLC - 2/21/2017  - revised to use Photo vs. Tree form
' ---------------------------------
Public Sub ClickAction(action As String)
On Error GoTo Err_Handler

    Dim fName As String, rName As String, oArgs As String
    Dim StartFolder As String, strPics As String, strPath As String
    
    action = LCase(Nz(Trim(action), ""))
    
    'defaults
    fName = ""
    rName = ""
    oArgs = ""
    
    Select Case Trim(action)
        'Where?
        Case "site"
            fName = "Site"
        Case "feature"
            fName = "Feature"
        Case "transect"
            fName = "VegTransect"
            oArgs = ""
        Case "plot"
            fName = "VegPlot"
        Case "location"
            fName = "Location"
        'Sampling
        Case "event"
            fName = "Events"
            oArgs = "" 'site & protocol IDs
        Case "vegplots"
            fName = "VegPlot"
            oArgs = "" 'site & protocol IDs
        Case "locations"
            fName = "Location"
            oArgs = "" 'collection source name - feature (A-G), transect #(1-8) &
        Case "people"
            fName = "Contact"
            oArgs = "Main"
        'VegetationS
        Case "veg plots"
            fName = "VegPlot"
        Case "woody canopy cover"
            fName = "VegWalk" '"WoodyCanopyCover"
            oArgs = "" '"1|2016|WCC"
        Case "understory cover"
        Case "vegetation walk"
            fName = "VegWalk"
        Case "species"
        Case "unknowns"
            fName = "Unknown"
        Case "species search"
            fName = "SpeciesSearch"
            oArgs = "Main"
            'disable double click events
            SetTempVar "originForm", "DisableDoubleClick"
'            If Not IsNull(TempVars("originForm")) Then
'                TempVars("originForm") = "DisableDoubleClick"
'            Else
'                TempVars.Add "originForm", "DisableDoubleClick"
'            End If
        'Observations
        Case "photos"
            fName = "Photo" '"Tree"
        Case "transducers"
            fName = "Transducer"
            oArgs = ""
        Case "Survey Files"
            fName = "SurveyFile"
            oArgs = ""
        Case "Upload Survey File"
            fName = ""
            oArgs = ""
            
            'handle upload
            StartFolder = GetSpecialFolderPath("FOLDERID_Recent")
            
            strPath = BrowseFolder("Select survey file to upload", "Confirm File", _
                                    StartFolder, , msoFileDialogFilePicker, "Survey files-CSV")
            
            If Len(strPath) > 0 Then
                'open data form before upload
                DoCmd.OpenForm "SurveyFile", acNormal, , , , , strPath
                
                'upload survey file
'                UploadCSVFile strPath
            End If
            
            'restore Main
'            ToggleForm "Main", 0
            
        Case "batch upload photos"
            fName = ""
            oArgs = ""
            
            'handle upload
            
            StartFolder = GetSpecialFolderPath("FOLDERID_Recent")
            
            strPath = BrowseFolder("Select directory with photos to upload", "Confirm Directory", _
                                    StartFolder)
            
            If Len(strPath) > 0 Then
                'ingest photos as "U" - unclassified
                IngestPhotos strPath, "U"
'            Else
'                MsgBox "Oops. Missed the directory the photos are in. " _
'                        & "Please re-select it.", vbOKOnly, "Missing Directory"
            End If
        
            'restore Main
            ToggleForm "Main", 0
        'Trip Prep
        Case "vegplot"
            rName = "VegPlot"
            oArgs = ""
        Case "vegwalk"
            rName = "VegWalk"
            oArgs = ""
        Case "photo"
            rName = "Photo"
            oArgs = ""
        Case "sediment class settings"
            fName = "ModWentworth"
        Case "sheet settings"
            fName = "SetDatasheetDefaults"
        Case "transducer"
            rName = "Transducer"
        Case "tasks"
            fName = "Task"
        'Reports
        Case "# Plots"
            rName = "NumPlots"
        Case "VegPlot - Species #s"
            rName = "NumSpecies"
        Case "VegPlot - Species"
            rName = "SpeciesCommon"
        Case "VegWalk - Species #s"
            rName = "NumSpeciesCommon"
        Case "VegWalk - Species"
            rName = "SpeciesUnique"
        Case "More..."
            fName = "AppReport"
    End Select

    If Len(fName) > 0 Then
        Forms("Main").Visible = False
        DoCmd.OpenForm fName, acNormal, OpenArgs:=oArgs
    ElseIf Len(rName) > 0 Then
        'print preview mode - acViewPreview
        DoCmd.OpenReport rName, acViewPreview
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClickAction[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          GetParks
' Description:  Retrieves list of parks from database
' Assumptions:  -
' Parameters:   active - flag if park is currently being sampled, 1-active, 0-inactive (boolean)
' Returns:      parks - list of park codes separated by "|" (string)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 18, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/18/2016  - initial version
' ---------------------------------
Public Function GetParks() As String
On Error GoTo Err_Handler

    'defaults
        

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParks[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          PopulateForm
' Description:  Populate a form using a specific record for edits
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 6/2/2016 - moved from forms (EventsList, TaglineList)
'   BLC - 8/8/2016 - revised to use default table name
'   BLC - 8/29/2016 - adjusted for Contact form (requires both Contact, Contact_Access data)
'                     using usys_temp_qdf & adjusting ID to Contact_ID in final SQL
'   BLC - 10/24/2016 - added ModWentworth form
'   BLC - 1/12/2017 - code cleanup
'   BLC - 2/14/2017 - added Task form
' ---------------------------------
Public Sub PopulateForm(frm As Form, ID As Long)
On Error GoTo Err_Handler
    Dim strSQL As String, strTable As String

    With frm
        'default
        strTable = .Name
        
        'find the form & populate its controls from the ID
        Select Case .Name
            Case "Contact"
                'requires Contact & Contact_Access data
                Dim qdf As DAO.QueryDef
                CurrentDb.QueryDefs("usys_temp_qdf").sql = GetTemplate("s_contact_access")
                
                strTable = "usys_temp_qdf"
                'set form fields to record fields as datasource
                'contact data
                .Controls("tbxID").ControlSource = "c.ID"
                .Controls("tbxFirst").ControlSource = "FirstName"
                .Controls("tbxMI").ControlSource = "MiddleInitial"
                .Controls("tbxLast").ControlSource = "LastName"
                .Controls("tbxEmail").ControlSource = "Email"
                .Controls("tbxUsername").ControlSource = "Username"
                .Controls("tbxOrganization").ControlSource = "Organization"
                .Controls("tbxPhone").ControlSource = "WorkPhone"
                .Controls("tbxPosition").ControlSource = "PositionTitle"
                .Controls("tbxExtension").ControlSource = "WorkExtension"
                'contact_access data
                .Controls("cbxUserRole").ControlSource = "Access_ID"
            Case "Events"
                strTable = "Event"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxSite").ControlSource = "Site_ID"
                .Controls("cbxLocation").ControlSource = "Location_ID"
                .Controls("tbxStartDate").ControlSource = "StartDate"
                .Controls("lblMsgIcon").Caption = ""
                .Controls("lblMsg").Caption = ""
            Case "Feature"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxFeature").ControlSource = "Feature"
                '.Controls("cbxLocation").ControlSource = ""
            Case "Location"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxName").ControlSource = "CollectionSourceName"
                .Controls("tbxDistance").ControlSource = "HeadToOrientDistance_m"
                .Controls("tbxBearing").ControlSource = "HeadToOrientBearing"
                .Controls("tbxNotes").ControlSource = "LocationNotes"
            Case "ModWentworth"
                strTable = "ModWentworthScale"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxClass").ControlSource = "Label"
                .Controls("tbxCode").ControlSource = "Code"
                .Controls("tbxSizeRange").ControlSource = "DiameterRange_mm"
                .Controls("tbxEffectiveDate").ControlSource = "ActiveYear"
                .Controls("tbxRetireDate").ControlSource = "RetireYear"
            Case "SetDatasheetDefaults"
                strTable = "tsys_Datasheet_Defaults"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxSpecies").ControlSource = "SpeciesRows"
                .Controls("tbxBlanks").ControlSource = "BlankRows"
            Case "Site"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxSiteCode").ControlSource = "SiteCode"
                .Controls("tbxSiteName").ControlSource = "SiteName"
                .Controls("tbxDescription").ControlSource = "SiteDescription"
                .Controls("tbxSiteDirections").ControlSource = "SiteDirections"
            Case "Tagline"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxCause").ControlSource = "HeightType"
                .Controls("tbxDistance").ControlSource = "LineDistance_m"
                .Controls("tbxHeight").ControlSource = "Height_cm"
            Case "Task"
                'set form fields to record fields as datasource
                .Controls("tbxType").ControlSource = "TaskType"
                .Controls("tbxTypeID").ControlSource = "TaskType_ID"
                '.Controls("lblTaskContext").Caption = .Controls("tbxType") & " (" & .Controls("tbxTypeID") & ")"
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxStatus").ControlSource = "Status_ID"
                .Controls("cbxPriority").ControlSource = "Priority_ID"
                .Controls("tbxTask").ControlSource = "Task"
                .Controls("cbxRequestedBy").ControlSource = "RequestedBy_ID"
                .Controls("tbxRequestDate").ControlSource = "RequestDate"
            Case "Transducer"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxTiming").ControlSource = "Timing"
                .Controls("cbxTransducer").ControlSource = "TransducerNumber"
                .Controls("tbxSerialNo").ControlSource = "SerialNumber"
                .Controls("tbxSampleDate").ControlSource = "ActionDate"
                .Controls("tbxSampleTime").ControlSource = "ActionTime"
                .Controls("chkSurveyed").ControlSource = "IsSurveyed"
            Case "VegPlot"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxNumber").ControlSource = "PlotNumber"
                .Controls("tbxDistance").ControlSource = "Distance"
                .Controls("cbxModalSedSize").ControlSource = "ModalSedSize"
                .Controls("tbxPctFines").ControlSource = "PctFines"
                .Controls("tbxPctWater").ControlSource = "PctWater"
                .Controls("tbxPctURC").ControlSource = "PctURC"
                .Controls("tbxPlotDensity").ControlSource = "PlotDensity"
                .Controls("chkNoCanopyVeg").ControlSource = "NoCanopyVeg"
                .Controls("chkNoRootedVeg").ControlSource = "NoRootedVeg"
                .Controls("chkNoIndicatorSpecies").ControlSource = "NoIndicatorSpecies"
                .Controls("chkHasSocialTrails").ControlSource = "HasSocialTrails"
            Case "VegTransect"
            
        End Select
    
'        'save record changes from form first to avoid "Write Conflict" errors
'        'where form & SQL are attempting to save record
'        'frm.Dirty = False
'
'        If frm.Dirty Then
'            MsgBox frm.Name & " DIRTY"
'            frm.Dirty = False
'        Else
'            MsgBox frm.Name & " CLEAN"
'        End If
        
        strSQL = GetTemplate("s_form_edit", "tbl" & PARAM_SEPARATOR & strTable & "|id" & PARAM_SEPARATOR & ID)
        
        'alter to retrieve proper ID
        Select Case .Name
            Case "Contact"
                strSQL = Replace(strSQL, " ID = ", " c.ID = ")
        End Select
        
        .RecordSource = strSQL
        
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateForm[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateCSVFields
' Description:  CSV field combobox populating actions
' Assumptions:  Control OnChange event = PopulateCSVFields([Screen].[ActiveControl])
'               where Screen.ActiveControl passes in the proper combobox
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Jeremy Cook, September 13, 2013
'   http://stackoverflow.com/questions/8787979/how-do-i-reference-the-current-form-in-an-expression-in-microsoft-access
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2016 - initial version
' ---------------------------------
'Public Sub PopulateCombobox(ctrl As ComboBox)
Public Function PopulateCSVFields(ctrl As Control) 'frm As Form) 'strName As String) 'ByRef ctrl As ComboBox)
On Error GoTo Err_Handler
    
    'Dim ctrl As ComboBox
    
    'Set ctrl = Forms("ImportMap").Controls("listCSV").Form.Controls("cbxColumnName2")
    
'    Set ctrl = Me.ActiveControl
    
'    'set displayed title
'    lblTitle.Caption = "CSV fields"
    
    'retrieve field info
    Dim aryFieldInfo() As Variant 'string
    
    aryFieldInfo = FetchDbTableFieldInfo("usys_temp_csv")
    
    'clear table
    ClearTable "usys_temp_rs2"
    
    'populate w/ table data
    Dim rs2 As DAO.Recordset
    Dim aryRecord() As String
    Dim i As Integer
    
    Set rs2 = CurrentDb.OpenRecordset("usys_temp_rs2", dbOpenDynaset)
    
    'add the "None" value
    rs2.AddNew
    rs2.Fields(0) = "None"
    rs2.Update
    
    For i = 0 To UBound(aryFieldInfo)
    
        'create new record
        rs2.AddNew
        
        aryRecord = Split(aryFieldInfo(i), "|")
        
        'rs!Column = aryRecord(0)
        rs2.Fields(0) = aryRecord(0)
    
        'add the new record
        rs2.Update
        
    Next
    
    Set ctrl.Recordset = rs2 '<--ERROR #5302
    
    Debug.Print "mod_App_UI PopulateCSVFields rs2 count = " & rs2.RecordCount
    
'    Dim strControl As String
'
'Debug.Print Me.NumColumns
'
'    'expose & populate the proper # of dropdowns
'    For i = 1 To Me.NumColumns 'CInt(Me.Records.RecordCount)
'        strControl = "cbxColumnName" & i
'Debug.Print strControl
'
''FIX HERE!
'        If i = 30 Then
'            Debug.Print "30"
'        End If
'
'        Me.Controls(strControl).Visible = True
''        Set Me.Controls(strControl).Recordset = rs2 '<--ERROR #5302
'        'Me.Controls(strControl).AddItem item:="None", index:=0
'
'        'set "None" to red --> Conditional formmating = "None"
'
'        'requery to refresh displayed controls
'        Me.Controls(strControl).Requery
'Debug.Print Me.Controls(strControl).ListRows
'    Next
'
'    If Me.NumColumns > 0 Then
'        'set detail to proper height
'        Me.Detail.Height = Me.Controls(strControl).Height * Me.NumColumns 'Me.Records.RecordCount
'    End If
'
''    Set Me.Recordset = rs
'
''    Set cbxColumnName.Recordset = rs2
'
'    'set the # of repeats of the cbx
''    Set Me.Recordset = rs

Exit_Handler:
    'cleanup
    Set rs2 = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateCSVFields[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          DeleteRecord
' Description:  Delete a specific record from a table
' Assumptions:  Assumes tbl name is properly capitalized & matches db table name
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 6/2/2016 - moved from forms (TaglineList, EventsList) to mod_App_UI
'   BLC - 6/27/2016- revised to match
' ---------------------------------
Public Sub DeleteRecord(tbl As String, ID As Long)
On Error GoTo Err_Handler
    Dim strSQL As String

    'find the form & populate its controls from the ID
    strSQL = GetTemplate("d_form_record", "tbl" & PARAM_SEPARATOR & tbl & "|id" & PARAM_SEPARATOR & ID)
    
    If IsNull(strSQL) Or Len(strSQL) = 0 Then GoTo Exit_Handler
Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    
    'show deleted record message & clear
    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
        tbl & PARAM_SEPARATOR & ID & _
        "|Type" & PARAM_SEPARATOR & "info"
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeleteRecord[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ClearFields
' Description:  initialize application values
' Assumptions:  -
' Parameters:   frm - Form whose fields should be cleared
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
'   BLC - 5/18/2015  - fixed error documentation ClearFields vs. ITIS_Click, mod_Forms vs. frm_SpeciesSearch
'   BLC - 6/30/2015  - moved to mod_App_UI
'   BLC - 7/5/2016   - added from Invasives Reporting mod_App_UI to support Species Search
' ---------------------------------
Public Sub ClearFields(frm As Form)
On Error GoTo Err_Handler

    Select Case frm.Name
    
        Case "frm_Species_Search"
            frm.Controls("cbxCO").DefaultValue = False
            frm.Controls("cbxUT").DefaultValue = False
            frm.Controls("cbxWY").DefaultValue = False
            frm.Controls("cbxITIS").DefaultValue = False
            frm.Controls("cbxCommon").DefaultValue = False
            frm.Controls("tbxSearchFor").Value = ""
    End Select
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearFields[mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          DisplayIcons
' Description:  Prepare icon set for reports
' Assumptions:  -
' Parameters:   icons - icons to display (delimited string)
'               delimiter - character splitting icons (string)
' Returns:      icon display translating icons field delimited string to display (string)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 24, 2016 - for NCPN tools
' Revisions:
'   BLC - 8/24/2016  - initial version
' ---------------------------------
Public Function DisplayIcons(icons As String, delimiter As String)
On Error GoTo Err_Handler

    Dim dDocIcons As Dictionary
    Dim ary() As String
    Dim strIcons As String
    Dim i As Integer
    
    Set dDocIcons = CreateObject("scripting.dictionary")
    
    'setup dictionary
    dDocIcons.Add "uDocument", uDocument
    dDocIcons.Add "uPDF", uNotepad ' uPDF
    
    'default
    strIcons = ""
    
    ary = Split(icons, delimiter)
    
    For i = LBound(ary) To UBound(ary)
    
        strIcons = strIcons & StringFromCodepoint(dDocIcons(ary(i))) & Space(2)
    
    Next
    
    DisplayIcons = strIcons
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisplayIcons[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          SortListForm
' Description:  form label sort on click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
'   Allen Browne, June 28, 2006
'   https://bytes.com/topic/access/answers/506322-using-orderby-multiple-fields
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/19/2017 - initial version
'   BLC - 1/31/2017 - adjusted to accommodate templates list
'   BLC - 2/21/2017 - adjusted to accommodate Contact list
' ---------------------------------
Public Sub SortListForm(frm As Form, ctrl As Control)
On Error GoTo Err_Handler

    Dim strSort As String
    
    'default
    strSort = ""
    
    'set sort field
    Select Case Replace(ctrl.Name, "lbl", "")
        Case "Email"
            strSort = "Email"
        Case "HdrID"
            strSort = "ID"
            Select Case frm.Name
                Case "ContactList"
                    strSort = "c.ID"
            End Select
        Case "Name"
            strSort = "LastName"
        Case "Template"
            strSort = "TemplateName"
        Case "SOPNum"
            strSort = "SOPNumber"
        Case "SOP"
            strSort = "FullName"
        Case "Syntax"
            strSort = "Syntax"
        Case "Version"
            strSort = "Version"
        Case "EffectiveDate"
            strSort = "EffectiveDate"
        Case ""
    End Select

    'set the sort
    If InStr(frm.OrderBy, strSort) = 0 Then
        frm.OrderBy = strSort
    ElseIf Right(frm.OrderBy, 4) = "Desc" Then
        frm.OrderBy = strSort
    Else
        frm.OrderBy = strSort & " Desc"
    End If
    
    frm.OrderByOn = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortListForm[mod_App_UI form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' Sub:          FilterListForm
' Description:  form filter click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Filter-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Public Sub FilterListForm(frm As Form, ctrl As Control)
On Error GoTo Err_Handler

    Dim strFilter As String
    
    'default
    strFilter = ""
    
    'set Filter field
    Select Case Replace(ctrl.Name, "lbl", "")
        Case "HdrID"
            strFilter = "ID"
        Case "Version"
            strFilter = "Version"
        Case "Template"
            strFilter = "Template"
        Case "Remarks"
            strFilter = "Remarks"
        Case "EffectiveDate"
            strFilter = "EffectiveDate"
        Case ""
    End Select

    'set the Filter
    If InStr(frm.OrderBy, strFilter) = 0 Then
        frm.OrderBy = strFilter
    ElseIf Right(frm.OrderBy, 4) = "Desc" Then
        frm.OrderBy = strFilter
    Else
        frm.OrderBy = strFilter & " Desc"
    End If
    
    frm.OrderByOn = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterListForm[mod_App_UI form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetContext
' Description:  set the context based on tempvars
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, October 14, 2015 - for NCPN tools
' Revisions:
'   BLC - 10/14/2016  - initial version
' ---------------------------------
Public Function GetContext() As String
On Error GoTo Err_Handler

    Dim strContext
    
    strContext = Nz(TempVars("ParkCode"), "") & Space(2) & ">" & Space(2) & _
                 Nz(TempVars("River"), "-") & Space(2) & ">" & Space(2) & _
                 Nz(TempVars("SiteCode"), "-")

    Select Case Nz(TempVars("ParkCode"), "")
    
        Case "BLCA"
            'add feature
            strContext = strContext & Space(2) & ">" & Space(2) & _
                 Nz(TempVars("Feature"), "-")
        Case "CANY"
            'site level
        Case "DINO"
            'site level
    End Select
    
    GetContext = strContext

Exit_Sub:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetContext[mod_App_UI])"
    End Select
    Resume Exit_Sub
End Function