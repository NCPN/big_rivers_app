Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.00
' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added PopulateSpeciesPriorities function from mod_Species
'               BLC, 11/19/2015 - 1.02 - added CreateEnums call to initApp
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
        Case "DB Admin" ' DB Admin main
            strTitle = ""
    End Select
    
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strTitle
        If strContext <> "DbAdmin" Then
            ctrl.visible = True
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
            'strInstructions = ""
    End Select
    
    'populate caption & display instructions
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strInstructions
        If strContext <> "DbAdmin" Then
            ctrl.visible = True
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
'   BLC - 11/19/2015 - added CreateEnums call to create application specific Enums,
'                      updated documentation to reflect mod_App_UI vs. mod_Init
' ---------------------------------
Public Sub Initialize()
On Error GoTo Err_Handler

    'create the enums specific to this application from the Enums table & mod_App_Enum stub module
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

Public Function SetStartupOptions(propertyname As String, _
    propertytype As Variant, propertyvalue As Variant) _
    As Boolean
  Dim dbs As Object
  Dim prp As Object
  Set dbs = Application.CurrentDb
  On Error Resume Next
  dbs.Properties(propertyname) = propertyvalue
  If Err.Number = 3270 Then
    Set prp = dbs.CreateProperty(propertyname, _
        propertytype, propertyvalue)
    dbs.Properties.Append prp
    Application.RefreshTitleBar
  Else
    SetStartupOptions = False
  End If
  Set dbs = Nothing
  Set prp = Nothing
End Function