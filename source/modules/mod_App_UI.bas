Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.04
' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added PopulateSpeciesPriorities function from mod_Species
'               BLC, 6/1/2015  - 1.02 - changed View to Search tab
'               BLC, 6/12/2015 - 1.03 - added EnableTargetTool button
'               BLC, 6/30/2015 - 1.04 - added ClearFields()
'               BLC, 7/27/2015 - 1.05 - added SetHints()
' =================================

' =================================
' SUB:     PopulateInsetTitle
' Description:  Sets inset title on form
' Assumptions:
' Parameters:   ctrl - control whose text is being set (control)
'               strContext - identifies which title to use,
'                            specifies the context for the title (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - initial version
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
'               BLC, 5/26/2015 - Added error handling
'               BLC, 6/4/2015 - Changed View to Search tab, added "or modify" for create tab
' =================================
Public Sub PopulateInsetTitle(ctrl As Control, strContext As String)
On Error GoTo Err_Handler
    
    Dim strTitle As String
    
    Select Case strContext
        Case "Create" ' Create main
            strTitle = "Choose what you'd like to create or modify"
        Case "CreateTgtLists" ' Create species target lists
            strTitle = "Create > Species Target Lists"
        Case "AddTgtArea" ' Add target areas
            strTitle = "Create > Add Target Area"
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates"  ' QA/QC > Outliers etc.
            strContext = Replace(Replace(strContext, "Suspect", "Suspect "), "Missing", "Missing ")
            strTitle = "Data Validation > " & strContext
        Case "Data Validation" ' QA/QC analysis project selection
            strTitle = "Data Validation > Field > Duplicates (NFV)" '<<<<< Make this so it ties back to the selected analysis
        Case "Search" ' Search main
            strTitle = "Species Search"
        Case "Reports" ' Reports main
            strTitle = "Reports"
        Case "CrewSpeciesList" ' Reports > Field Crew Species List
            strTitle = "Reports > Field Crew Species List"
        Case "SpeciesListByPark" ' Reports > Species List By Park
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
        Case "DbAdmin" ' DB Admin main
            strTitle = "Db Admin"
    End Select
    
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strTitle
        If strContext <> "DbAdmin" Or DB_ADMIN_CONTROL = False Then
            ctrl.visible = True
        End If
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateInsetTitle[mod_App_UI])"
    End Select
    Resume Exit_Sub
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
'               BLC, 6/4/2015  - Changed View to Search
' =================================
Public Sub PopulateInstructions(ctrl As Control, strContext As String)
On Error GoTo Err_Handler
    Dim strInstructions As String
    
    'MsgBox strContext
    
    Select Case strContext
        Case "Create" ' Create main
            strInstructions = "Choose what you would like to create/modify."
        Case "CreateTgtLists" ' Create > Species Target Lists
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your list." & vbCrLf & vbCrLf & _
                    "Only existing lists for the current or future years may be modified." & vbCrLf & vbCrLf & _
                    "Please contact the project lead or data management if a prior year list must be modified."
        Case "AddTgtArea" ' Create > Add Target Area
            strInstructions = "" '"Choose the park and year for your target area. Click 'Continue' to create your area."
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates" ' QA/QC main
            strInstructions = "Complete the fields to define the data set or subset you are validating. " _
                    & "Leave the fields blank if you are validating all data. Click 'Run' to validate."
        Case "Search" ' Search main
            strInstructions = "Search for species family, name, codes. " & _
                    "Latin, common, and state specific (UT, CO, WY) genus species names " & _
                    "and lookup (6-letter) and ITIS codes are included." & vbCrLf & vbCrLf & _
                    "Searches can be made across all or only a few species names/codes."
            'strInstructions = "Log your modifications to data within the edit log. " _
            '        & "Be as complete as possible to aid others in tracing data changes."
        Case "Reports" ' Reports main
            strInstructions = "Choose the report you would like to run."
        Case "CrewSpeciesList" ' Reports > Field Crew Species List
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "SpeciesListByPark" ' Reports > Species List By Park
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
    End Select
    
    'populate caption & display instructions
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strInstructions
        If strContext <> "DbAdmin" Or DB_ADMIN_CONTROL = False Then
            ctrl.visible = True
        End If
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateInstructions[mod_App_UI])"
    End Select
    Resume Exit_Sub
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
Public Function PopulateSpeciesPriorities(parkCode As String, priorities As String) As String

On Error GoTo Err_Handler

Dim ParkPriorities As Variant
Dim i As Integer

    'check if parkCode is in priorities string
    If Len(priorities) > Len(Replace(priorities, parkCode, "")) Then
    
        'prepare the Park Priority values
        ParkPriorities = Split(priorities, "|")
        
        'set park priority values
        For i = 0 To UBound(ParkPriorities)
            'does Park have a priority value?
            If parkCode = Left(ParkPriorities(i), 4) Then
                PopulateSpeciesPriorities = Replace(ParkPriorities(i), parkCode + "-", "")
            End If
        Next
        
    Else
        'not listed
        PopulateSpeciesPriorities = "X"
    
    End If
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateSpeciesPriorities[mod_App_UI])"
    End Select
    Resume Exit_Function
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
' ---------------------------------
Public Sub Initialize()
On Error GoTo Err_Handler


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Initialize[mod_Init])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          EnableTargetTool
' Description:  enable the target tool button
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/4/2015  - initial version
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Public Sub EnableTargetTool(ctrl As Control)
On Error GoTo Err_Handler
    
    'enable button if connected
    If TempVars("Connected") Then
        ctrl.Enabled = True
    Else
        ctrl.Enabled = False
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableTargetTool[mod_Init])"
    End Select
    Resume Exit_Sub
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
' ---------------------------------
Public Sub ClearFields(frm As Form)
On Error GoTo Err_Handler

    Select Case frm.name
    
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
            "Error encountered (#" & Err.Number & " - ClearFields[form_Forms])"
    End Select
    Resume Exit_Sub
End Sub

' ================================ Big Rivers ===========================

' ---------------------------------
' SUB:          SetHints
' Description:  set field hints for form
' Assumptions:  -
' Parameters:   frm - form where fields reside(form object)
'               strForm - name of subform (string)
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

'Forms!Mainform!Subform1.Form!
    
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
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHints[form_Forms])"
    End Select
    Resume Exit_Sub
End Sub