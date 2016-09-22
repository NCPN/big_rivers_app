Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_User
' Level:        Framework module
' Version:      1.07
' Description:  Access related functions & subroutines
'
' Source/date:  Bonnie Campbell, May 2014
' Revisions:    BLC, 6/13/2014 - 1.00 - initial version
'               BLC, 8/6/2014  - 1.01 - renamed from mod_Access_Utilities to mod_User
'                                to accommodate additional user functions/subs
'               BLC, 4/30/2015 - 1.02 - added level & version info, move blnRunQueries & blnUpdateAll to mod_Initialize_App
'               BLC, 5/18/2015 - 1.03 - removed fxn prefixes
'               BLC, 6/12/2015 - 1.04 - replaced TempVars.item(... with TempVars("...
'               BLC, 6/30/2015 - 1.05 - updated cmd button prefixes to btn
'               BLC, 4/4/2016  - 1.06 - changed Exit_Handler/Exit_Procedure > Exit_Handler
'               BLC, 4/20/2016 - 1.07 - added getADCommonName() to retrieve person's name
'               BLC, 6/30/2016 - 1.08 - adjusted to use GetTemplate
' =================================

' ---------------------------------
' GLOBALS:      global variables
' Description:  variables provide project-wide globally accessible values
' References:   -
' Source/date:  Bonnie Campbell, August 2014
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version
'               BLC, 4/30/2015 - moved blnRunQueries & blnUpdateAll to mod_Initialize_App
' ---------------------------------
' N/A

' ---------------------------------
' FUNCTION:     getDbUserAccess
' Description:  Retrieves user access level from tsysUser & sets TempVars.Item("UserAccessLevel")
'               If User_role exists, this value is set otherwise
'               TempVars.Item("UserAccessLevel") defaults to "read only"
' Parameters:   -
' Returns:      UserAccessLevel - string (read only, data entry, power user, admin)
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, August, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 8/21/2014 - initial version
'               BLC, 4/22/2015 - handle global USER_ACCESS_CONTROL setting to enable full access
'                                for apps w/o user access controls
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 4/4/2016  - changed Exit_Function > Exit_Handler, dbCurrent to CurrentDb
'               BLC, 6/30/2016 - revised to use GetTemplate()
' ---------------------------------
Public Function getDbUserAccess() As String
On Error GoTo Err_Handler
Dim strSQL As String
Dim rs As DAO.Recordset

    '-------------------------------
    ' set defaults
    '-------------------------------

    If USER_ACCESS_CONTROL Then
    
'        strSQL = "SELECT User_role FROM tsys_User_Roles WHERE User_name = '" & Environ("Username") & "';"
'Debug.Print strSQL
'        'fetch User role & set UserAccessLevel
'        'Set rs = dbCurrent.OpenRecordset(strSQL)
'        Set rs = CurrentDb.OpenRecordset(strSQL)
        
        Dim db As DAO.Database
        Dim qdf As DAO.QueryDef
        
        Set db = CurrentDb
        
        With db
            Set qdf = .QueryDefs("usys_temp_qdf")
                    
            With qdf
                .SQL = GetTemplate("s_user_access")
                
                '-- required parameters --
                .Parameters("appuser") = Environ("Username")
                
                '-- optional parameters --
                
                Set rs = .OpenRecordset(dbOpenDynaset)
                '.Execute dbFailOnError
                
        
        If Not rs.BOF And Not rs.EOF Then
            'db user role
            TempVars("UserAccessLevel") = CStr(rs!AccessLevel)
        Else
            'default
            TempVars("UserAccessLevel") = "read only"
        End If
                
                
                'cleanup
                .Close
                Set rs = Nothing
            End With
            
        End With
        
        
'        If Not rs.BOF And Not rs.EOF Then
'            'db user role
'            TempVars("UserAccessLevel") = CStr(rs!User_role)
'        Else
'            'default
'            TempVars("UserAccessLevel") = "read only"
'        End If
        
    Else
        'default for apps w/o user access controls
        TempVars("UserAccessLevel") = "admin"
    End If
    
    getDbUserAccess = TempVars("UserAccessLevel")
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getUserAccess[mod_User])"
    End Select
    Resume Exit_Handler

End Function

' ---------------------------------
' SUB:          setUserAccess
' Description:  Sets form control access/functionality based on global user access level - TempVars.Item("UserAccessLevel")
' Parameters:   frm  - form referencing the function
'               flag - optional string value for identifying specific form control change subsets
'                      for example for the Switchboard form's Update condition
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, July, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 7/29/2014 - initial version
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2015 - updated cmd button prefixes to btn
'               BLC, 4/4/2016  - changed Exit_Procedure > Exit_Handler
'               BLC, 6/30/2016 - revised to use GetTemplate()
' ---------------------------------
Public Sub setUserAccess(frm As Form, Optional flag As String)
On Error GoTo Err_Handler
    
    '-------------------------------
    ' set defaults
    '-------------------------------
Debug.Print "useraccesslvl=" & TempVars("UserAccessLevel")

    '-------------------------------
    ' get db user access level
    '-------------------------------
    getDbUserAccess
    
'-------------------------------
' form control access
'-------------------------------
    With frm
    
    '-------------------------------
    ' set functionality based on access level
    '-------------------------------
        Select Case TempVars("UserAccessLevel")

        '-------------------------------
        ' admin (full control)
        '-------------------------------
            Case "admin"
                Select Case .Name
                    Case "frm_App_Releases"
                        With .subBugs
                            .Locked = False
                            .Form.AllowAdditions = True
                        End With
                    Case "frm_Set_Defaults"
                        .tbxTimeframe.Locked = False
                        .tbxTimeframe.backstyle = 1
                        .tbxProject.Enabled = True
                        .tbxProject.Locked = False
                        .tbxProject.backstyle = 1
                        .tbxProject.SpecialEffect = 2
                        .btnNewUser.Enabled = True
                    Case "frm_Switchboard"
                        'handle updates
                        If flag = "update" Then
                            With !fsub_DbAdmin.Form
                                !btnEditLog.Enabled = True
                                !btnSetRoles.Enabled = True
                                !btnDbWindow.Enabled = True
                                !pgSettings.Enabled = True
                                !btnChangeDbInfo.Enabled = True
                                
                                '!tbxAppMode.Locked = False
                                '!tbxAppMode.Enabled = True
                                !tbxAppMode = TempVars("UserAccessLevel")
                                '!tbxAppMode.Enabled = False
                                '!tbxAppMode.Locked = True
                            End With
                            CurrentDb.Properties("AllowFullMenus") = True
                            CurrentDb.Properties("AllowShortcutMenus") = True
                            CurrentDb.Properties("AllowBuiltInToolbars") = True
                        End If
                    Case "DbAdmin", _
                         "fsub_DbAdmin", "frm_Data_Gateway", "frm_Edit_Log", "frm_Lookups", "frm_QA_Tool", _
                         "frm_Data_Browser"
                        'admin & power user settings are same
                        GoTo Admin_PowerUser:
                    Case ""
                End Select
        
        '-------------------------------
        ' power user
        '-------------------------------
            Case "power user"
                Select Case .Name
                    Case "frm_App_Releases"
                        'power user, data entry & read only the same
                        GoTo Catch_All_Else:
                    Case "frm_Set_Defaults"
                        .tbxTimeframe.Locked = False
                        .tbxTimeframe.backstyle = 1
                        .tbxProject.Enabled = False
                        .tbxProject.Locked = True
                        .tbxProject.backstyle = 0
                        .tbxProject.SpecialEffect = 0
                        .btnNewUser.Enabled = True
                    Case "frm_Switchboard"
                        'handle updates
                        If flag = "update" Then
                            With !fsub_DbAdmin.Form
                                !btnEditLog.Enabled = True
                                !btnSetRoles.Enabled = True
                                !btnDbWindow.Enabled = True
                                !pgSettings.Enabled = True
                                !btnChangeDbInfo.Enabled = False
                                '!tbxAppMode.Locked = False
                                !tbxAppMode = TempVars("UserAccessLevel")
                                '!tbxAppMode.Locked = True
                            End With
                            CurrentDb.Properties("AllowFullMenus") = True
                            CurrentDb.Properties("AllowShortcutMenus") = True
                            CurrentDb.Properties("AllowBuiltInToolbars") = True
                        End If
                    Case "DbAdmin", _
                        "fsub_DbAdmin", "frm_Data_Gateway", "frm_Edit_Log", "frm_Lookups", "frm_Data_Browser"
                        'admin & power user settings are same
                        GoTo Admin_PowerUser:
                    Case ""
                End Select
        
        '-------------------------------
        ' data entry
        '-------------------------------
            Case "data entry"
                Select Case .Name
                    Case "frm_App_Releases"
                        'power user, data entry & read only the same
                        GoTo Catch_All_Else:
                    Case "frm_Set_Defaults"
                        .tbxTimeframe.Locked = True
                        .tbxTimeframe.backstyle = 0
                        .tbxProject.Enabled = False
                        .tbxProject.Locked = True
                        .tbxProject.backstyle = 0
                        .tbxProject.SpecialEffect = 0
                        .btnNewUser.Enabled = True
                    Case "fsub_DbAdmin"
                        .btnEditLog.Enabled = False   ' automatically opens if certified data are edited
                        .btnSetRoles.Enabled = False
                        .btnDbWindow.Enabled = False
                        .pgSettings.Enabled = True
                        .btnChangeDbInfo.Enabled = False
                    Case "frm_Switchboard"
                        With !fsub_DbAdmin.Form
                            !btnEditLog.Enabled = False   ' automatically opens if certified data are edited
                            !btnSetRoles.Enabled = False
                            !btnDbWindow.Enabled = False
                            !pgSettings.Enabled = True
                            !btnChangeDbInfo.Enabled = False
                            '!tbxAppMode.Locked = False
                            !tbxAppMode = TempVars("UserAccessLevel")
                            '!tbxAppMode.Locked = True
                        End With
                        ' Turn off options (only apparent after the next time app is opened)
                        CurrentDb.Properties("AllowFullMenus") = False
                        CurrentDb.Properties("AllowShortcutMenus") = False
                        CurrentDb.Properties("AllowBuiltInToolbars") = False
                    Case "frm_Data_Gateway"
                        .optgCertifiedMode.Enabled = False    ' no cert data to minimize confusion
                        .optgScheduledMode.Enabled = True     ' user can enter data for unscheduled locs
                        .cmbYearFilter.Locked = True          ' current year only to minimize confusion
                        .togFilterByYear.Enabled = False
                        .txtLoc_code.ForeColor = 0            ' black - no direct link to data browser
                        .btnDeleteRec.Enabled = True
                        .btnNewLoc.Enabled = True
                        .btnNewRareObs.Enabled = True
                    Case "frm_Edit_Log"
                        ' Disable delete, edit, and filter controls
                        .DataEntry = True
                        .AllowDeletions = False
                        .cmbYearFilter.Enabled = False
                        .togFilterByYear.Enabled = False
                        .cmbTypeFilter.Enabled = False
                        .togFilterByType.Enabled = False
                        .cmbUserFilter.Enabled = False
                        .togFilterByUser.Enabled = False
                    Case "frm_Lookups"
                        .btnNewTaxon.Enabled = True
                        .optgFormMode.Enabled = True
                        .subLookupTables.Locked = True
                    Case "frm_QA_Tools"
                        .pgDataTables.Visible = False
                        .btnDesignView.Enabled = False
                        .optgScope.Enabled = False
                        .cmbTimeframe.Enabled = True
                        .optgMode.Enabled = False
                        .btnRefresh.Enabled = True
                        .btnRequery.Enabled = True
                        .btnOpenRecord.Enabled = True
                        .btnAutoFix.Enabled = False
                        .selObject.Enabled = False
                        ' Run the queries if the user selects Yes
                        If MsgBox("Would you like to run the QA queries now?" & vbCrLf & _
                            "'No' opens the form without running queries ...", _
                            vbYesNo, "Quality Assurance Data Checks") = vbYes Then
                            blnRunQueries = True
                            UpdateQAResults
                        End If
                    Case "frm_Data_Browser"
                        .subSite.Form.AllowEdits = False
                        .subSchedule.Form.AllowEdits = False
                        .subSchedule.Form.AllowAdditions = False
                        .subSchedule.Form.AllowDeletions = False
                        .subTarget_coords.Form.AllowEdits = False
                        .subCoordinates.Form.AllowEdits = False
                        .subTasks.Form.AllowEdits = True          ' exception
                        .subTasks.Form.AllowAdditions = True      ' exception
                        .subTasks.Form.AllowDeletions = False
                        .subEvents.Form.AllowEdits = False
                        .subEvents.Form!btnDelete.Enabled = False
                        .subImages.Form.AllowEdits = True         ' exception
                        .subImages.Form.AllowAdditions = True     ' exception
                        .subImages.Form.AllowDeletions = True     ' exception
                        .subMarkers.Form.AllowEdits = True        ' exception
                        .subMarkers.Form.AllowAdditions = True    ' exception
                        .subMarkers.Form.AllowDeletions = False
                        .subFeatures.Form.AllowEdits = True       ' exception
                        .subFeatures.Form.AllowAdditions = True   ' exception
                        .subFeatures.Form.AllowDeletions = True   ' exception
                        .subAnalysisNotes.Form.AllowEdits = False
                        .subAnalysisNotes.Form.AllowAdditions = False
                        .subAnalysisNotes.Form.AllowDeletions = False
                        .subVarianceGroups.Form.AllowEdits = False
                        .subVarianceGroups.Form.AllowAdditions = False
                        .subVarianceGroups.Form.AllowDeletions = False
                        .btnNew.Enabled = False
                    Case ""
                End Select
                
        '-------------------------------
        ' read only (default)
        '-------------------------------
            Case Else
                Select Case .Name
                    Case "frm_App_Releases"
                        'power user, data entry & read only the same
                        GoTo Catch_All_Else:
                    Case "frm_Switchboard"
                        ' Set default application mode to read only, and set up form accordingly in case of error
                        ' requires frm!fsub_DBAdmin.Form format or error #438 object doesn't support this property/method occurs
                        With !fsub_DbAdmin.Form
                            .btnSetRoles.Enabled = False
                            .btnDbWindow.Enabled = False
                            .pgSettings.Enabled = False
                            .btnChangeDbInfo.Enabled = False
                            '!tbxAppMode.Locked = False
                            '!tbxAppMode = TempVars("UserAccessLevel")
                            '!tbxAppMode.Locked = True
                        End With
                        
                        'handle updates
                        If flag = "update" Then
                            !fsub_DbAdmin.Form!btnEditLog.Enabled = True
                            ' Turn off options (only apparent after the next time app is opened)
                            CurrentDb.Properties("AllowFullMenus") = False
                            CurrentDb.Properties("AllowShortcutMenus") = False
                            CurrentDb.Properties("AllowBuiltInToolbars") = False
                        End If
                    Case "frm_Set_Defaults"
                        .tbxTimeframe.Locked = True
                        .tbxTimeframe.backstyle = 0
                        .tbxProject.Enabled = False
                        .tbxProject.Locked = True
                        .tbxProject.backstyle = 0
                        .tbxProject.SpecialEffect = 0
                        .btnNewUser.Enabled = False
                    Case "DbAdmin", "fsub_DbAdmin"
                        .btnEditLog.Enabled = True
                        .btnSetRoles.Enabled = False
                        .btnDbWindow.Enabled = False
                        .pgSettings.Enabled = False
                        .btnChangeDbInfo.Enabled = False
                    Case "frm_Data_Gateway"
                        .optgCertifiedMode.Enabled = True     ' cert data are protected in this mode
                        .optgScheduledMode.Enabled = False
                        .cmbYearFilter.Locked = False         ' enable seeing multiple years
                        .togFilterByYear.Enabled = True
                        .txtLoc_code.ForeColor = 0            ' black - no direct link to data browser
                        .btnDeleteRec.Enabled = False
                        .btnNewLoc.Enabled = False
                        .btnNewRareObs.Enabled = False
                    Case "frm_Edit_Log"
                        .DataEntry = False
                        .AllowAdditions = False
                        .AllowDeletions = False
                        .cmbYearFilter.Enabled = True
                        .togFilterByYear.Enabled = True
                        .cmbTypeFilter.Enabled = True
                        .togFilterByType.Enabled = True
                        .cmbUserFilter.Enabled = True
                        .togFilterByUser.Enabled = True
                    Case "frm_Lookups"
                        .btnNewTaxon.Enabled = False
                        .optgFormMode.Enabled = False
                        .subLookupTables.Locked = True
                    Case "fsub_Events_Browser"
                        .btnEdit.Caption = "View"               ' Set button caption
                    Case "frm_QA_Tools"
                        .pgDataTables.Visible = False
                        .btnDesignView.Enabled = False
                        .optgScope.Enabled = False
                        .cmbTimeframe.Enabled = True
                        .optgMode.Enabled = False
                        .btnRefresh.Enabled = False
                        .btnRequery.Enabled = False
                        .btnOpenRecord.Enabled = False
                        .btnAutoFix.Enabled = False
                        .selObject.Enabled = False
                    Case "frm_Data_Browser"
                        .subSite.Form.AllowEdits = False
                        .subSchedule.Form.AllowEdits = False
                        .subSchedule.Form.AllowAdditions = False
                        .subSchedule.Form.AllowDeletions = False
                        .subTarget_coords.Form.AllowEdits = False
                        .subCoordinates.Form.AllowEdits = False
                        .subTasks.Form.AllowEdits = False
                        .subTasks.Form.AllowAdditions = False
                        .subTasks.Form.AllowDeletions = False
                        .subEvents.Form.AllowEdits = False
                        .subEvents.Form!btnDelete.Enabled = False
                        .subImages.Form.AllowEdits = False
                        .subImages.Form.AllowAdditions = False
                        .subImages.Form.AllowDeletions = False
                        .subMarkers.Form.AllowEdits = False
                        .subMarkers.Form.AllowAdditions = False
                        .subMarkers.Form.AllowDeletions = False
                        .subFeatures.Form.AllowEdits = False
                        .subFeatures.Form.AllowAdditions = False
                        .subFeatures.Form.AllowDeletions = False
                        .subAnalysisNotes.Form.AllowEdits = False
                        .subAnalysisNotes.Form.AllowAdditions = False
                        .subAnalysisNotes.Form.AllowDeletions = False
                        .subVarianceGroups.Form.AllowEdits = False
                        .subVarianceGroups.Form.AllowAdditions = False
                        .subVarianceGroups.Form.AllowDeletions = False
                        .btnNew.Enabled = False
                    Case ""
                End Select
           
        End Select

    End With
    
Exit_Handler:
    Exit Sub
    
'------------------------------------
' admin & power user (both the same)
'------------------------------------
Admin_PowerUser:

    With frm
        Select Case .Name
            Case "DbAdmin", "fsub_DbAdmin"
                .btnEditLog.Enabled = True
                .btnSetRoles.Enabled = True
                .btnDbWindow.Enabled = True
                .pgSettings.Enabled = True
                .btnChangeDbInfo.Enabled = False
            Case "frm_Data_Gateway"
                .optgCertifiedMode.Enabled = True
                .optgScheduledMode.Enabled = True
                .cmbYearFilter.Locked = False
                .togFilterByYear.Enabled = True
                .txtLoc_code.ForeColor = 16711680     ' blue to indicate link to data browser
                .btnDeleteRec.Enabled = True
                .btnNewLoc.Enabled = True
                .btnNewRareObs.Enabled = True
            Case "frm_Edit_Log"
                .DataEntry = False
                .AllowAdditions = True
                .AllowDeletions = True
                ' Enable the filter controls
                .cmbYearFilter.Enabled = True
                .togFilterByYear.Enabled = True
                .cmbTypeFilter.Enabled = True
                .togFilterByType.Enabled = True
                .cmbUserFilter.Enabled = True
                .togFilterByUser.Enabled = True
            Case "frm_Lookups"
                .btnNewTaxon.Enabled = True
                .optgFormMode.Enabled = True
                .subLookupTables.Locked = False
            Case "frm_Schedule"
                .subSchedule.Form.AllowAdditions = True
                .subSchedule.Form.AllowDeletions = True
                .subSchedule.Form.AllowEdits = True
                .btnAutoPopulate.Enabled = True
            Case "frm_QA_Tool"
                .pgDataTables.Visible = True
                .btnDesignView.Enabled = True
                .optgScope.Enabled = True
                .cmbTimeframe.Enabled = True
                .optgMode.Enabled = True
                .btnRefresh.Enabled = True
                .btnRequery.Enabled = True
                .btnOpenRecord.Enabled = True
                .btnAutoFix.Enabled = True
                .selObject.Enabled = True
                ' Run the queries if the user selects Yes
                If MsgBox("Would you like to run the QA queries now?" & vbCrLf & _
                    "'No' opens the form without running queries ...", _
                    vbYesNo, "Quality Assurance Data Checks") = vbYes Then
                    blnRunQueries = True
                    UpdateQAResults
                End If
            Case "frm_Data_Browser"
                .subSite.Form.AllowEdits = True
                .subSchedule.Form.AllowEdits = True
                .subSchedule.Form.AllowAdditions = True
                .subSchedule.Form.AllowDeletions = True
                .subTarget_coords.Form.AllowEdits = True
                .subCoordinates.Form.AllowEdits = True
                .subTasks.Form.AllowEdits = True
                .subTasks.Form.AllowAdditions = True
                .subTasks.Form.AllowDeletions = True
                .subEvents.Form.AllowEdits = True
                .subEvents.Form!btnDelete.Enabled = True
                .subImages.Form.AllowEdits = True
                .subImages.Form.AllowAdditions = True
                .subImages.Form.AllowDeletions = True
                .subMarkers.Form.AllowEdits = True
                .subMarkers.Form.AllowAdditions = True
                .subMarkers.Form.AllowDeletions = True
                .subFeatures.Form.AllowEdits = True
                .subFeatures.Form.AllowAdditions = True
                .subFeatures.Form.AllowDeletions = True
                .subAnalysisNotes.Form.AllowEdits = True
                .subAnalysisNotes.Form.AllowAdditions = True
                .subAnalysisNotes.Form.AllowDeletions = True
                .subVarianceGroups.Form.AllowEdits = True
                .subVarianceGroups.Form.AllowAdditions = True
                .subVarianceGroups.Form.AllowDeletions = True
                .btnNew.Enabled = True
            Case ""
        End Select
    End With
    
    GoTo Exit_Handler

'------------------------------------------
' admin, power user, data entry (same)
'------------------------------------------
Admin_PowerUser_DataEntry:
    With frm
        Select Case .Name
            Case "fsub_Events_Browser"
                .btnEdit.Caption = "Edit"           'Set button caption
            Case ""
        End Select
    End With
    
    GoTo Exit_Handler

'------------------------------------
' data entry & read only (both the same)
'------------------------------------
DataEntry_ReadOnly:
    With frm
        Select Case .Name
            Case "frm_Schedule"
                .subSchedule.Form.AllowAdditions = False
                .subSchedule.Form.AllowDeletions = False
                .subSchedule.Form.AllowEdits = False
                .btnAutoPopulate.Enabled = False
            Case ""
        End Select
    End With
    
    GoTo Exit_Handler

'------------------------------------------
' power user, data entry, read only (same)
'------------------------------------------
Catch_All_Else:
    With frm
        Select Case .Name
            Case "frm_App_Releases"
                With .subBugs
                    .Locked = True
                    .Form.AllowAdditions = False
                End With
            Case ""
        End Select
    End With
    
    GoTo Exit_Handler

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetUserAccess[mod_User])"
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          logUserAction
' Description:  Logs user activities
' Parameters:   frm - form referencing the function
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, August, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 8/8/2014 - initial version
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/30/2016 - revised to use SWITCHBOARD vs. frm_Switchboard and GetTemplate()
' ---------------------------------
Public Sub logUserAction(frm As Form)
On Error GoTo Err_Handler
Dim strSQL As String

    '-------------------------------
    ' set defaults
    '-------------------------------
    With frm
    
        Select Case .Name
            Case SWITCHBOARD
                ' Log the user exit time if the back end is connected and the user has write privileges
                If TempVars("Connected") And TempVars("WritePermission") Then
'                    strSQL = "INSERT INTO tsys_Logins ( User_name, Action_taken ) SELECT '" _
'                        & Environ("Username") & "' AS User, 'close' AS Action;"
'                    strSQL = GetTemplate("tsys_login_action")
                    Dim db As DAO.Database
                    Dim qdf As DAO.QueryDef
                    
                    Set db = CurrentDb
                    
                    With db
                        Set qdf = .QueryDefs("usys_temp_qdf")
                                
                        With qdf
                            .SQL = GetTemplate("tsys_login_action")
                            
                            '-- required parameters --
                            .Parameters("environ_username") = Environ("Username")
                            .Parameters("action") = "close"
                            
                            '-- optional parameters --
                            
                            .Execute dbFailOnError
                            
                            'cleanup
                            .Close
                        End With
                        
                    End With
                
                End If
        End Select
    End With

    'SQL check
'    If Len(strSQL) > 0 Then
'        DoCmd.SetWarnings False
'        DoCmd.RunSQL strSQL
'    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - logUserAction[mod_db_Utilities])"
    End Select
    Resume Exit_Handler

End Sub

' =================================
' FUNCTION:     UserName
' Description:  Returns the current user name
' Parameters:   none
' Returns:      string of the user login
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 12/31/2009
' Revisions:    JRB, 12/31/2009 - initial version
'               BLC, 4/30/2015  - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 4/4/2016  - changed Exit_Function > Exit_Handler
' =================================
Public Function Username() As String
    On Error GoTo Err_Handler

    Username = "Unknown"
    Username = Environ("Username")

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UserName[mod_User])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     GetADCommonName
' Description:  Returns the currently logged in user's common name
' Note:
'       AD/LDAP directories use objects referenced by
'       Distinguished Name (DN)
'       DN parts are delimited by commas & represent where in
'       the AD hierarchy objects exist
'       DN monikers include:
'           Common name (cn) - John Doe
'           Organizational unit (ou) - Marketing
'           Domain component (dc) - MyCompany
'
' Parameters:   none
' Returns:      common name (cn) - Joe User (string)
' Throws:       none
' References:
'   Andrey V Artemyev & h2fsell, February 22, 2011
'   https://social.msdn.microsoft.com/Forums/office/en-US/1b38c289-af3f-471a-8bb6-d04d6e281964/os-compatible-way-to-get-the-full-user-name-via-vba?forum=accessdev
' Source/date:  Bonnie Campbell, April 20, 2016
' Revisions:    BLC, 4/20/2016 - initial version
' =================================
Public Function GetADCommonName() As String
On Error GoTo Err_Handler

    Dim ADSI As Object, tmp As Variant
 
    Set ADSI = CreateObject("ADSystemInfo")
    
    tmp = Split(ADSI.Username, ",")
 
    'retrieve first & last name
    GetADCommonName = Right(tmp(0), Len(tmp(0)) - 3)
  
' Automation Error
'No mapping between account names and security IDs was done.
'Print Err.Number
'-2147023564

 
Exit_Handler:
    'cleanup
    Set ADSI = Nothing

    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UserName[mod_User])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     AccessID
' Description:  Returns the access level ID from the access level name
' Assumptions:  Access levels include:
'                   admin, power user, data entry, read only
' Parameters:   AccessLevel - name of the level (string)
' Returns:      access ID (long)
' Throws:       none
' References:   none
' Source/date:  B. Campbell, 6/7/2016
' Revisions:    BLC, 6/7/2016 - initial version
' =================================
Public Function AccessID(AccessLevel As String) As Long
    On Error GoTo Err_Handler

    Dim ID As Long, i As Long
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    strSQL = GetTemplate("s_access_level", "lvl" & PARAM_SEPARATOR & LCase(AccessLevel))
    
    Set rs = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If rs.RecordCount > 0 Then
        AccessID = rs("ID")
    Else
        'default --> 0 = no access, exit application
        AccessID = 0
    End If

Exit_Handler:
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetAccessID[mod_User])"
    End Select
    Resume Exit_Handler
End Function