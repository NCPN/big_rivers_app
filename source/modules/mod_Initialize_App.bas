Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Initialize_App
' Level:        Framework module
' Version:      1.03
' Description:  Standard module for setting initial app & database values/settings & global variables
' Source/date:  Bonnie Campbell, July 2014
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - 1.00 - initial version
'               BLC, 8/6/2014  - 1.01 - merged in mod_Global_Variables (see history below)
'               BLC, 4/22/2015 - adapted to generic tools (NCPN Invasives Reporting Tool) by adding
'                                USER_ACCESS_CONTROL (False - gives users full control in apps w/o controls,
'                                                     True - relies on user access control settings)
'                                DB_SYS_TABLES & APP_SYS_TABLES (handle table arrays for the database/
'                                   application)
'                                WQ Utilities tool constants removed (WATER_YEAR_START & WATER_YEAR_END)
'               BLC, 4/30/2015 - 1.02 - shifted USER_ACCESS_CONTROL, DB_SYS_TABLES, APP_SYS_TABLES to mod_App_Settings
'                                since these are application vs. framework specific, added Level & Version #
'                                added blnRunQueries & blnUpdateAll from mod_User
'               BLC, 7/7/2015  - 1.03 - added SafeStart() to set error trapping for the application
'                                to "Break in Class Module"
' =================================
' HISTORY:
' MERGED MODULE: mod_Global_Variables (merged with mod_Initialize_App)
' Description:   Standard module for dimensioning global variables
' Source/date:   John R. Boetsch, May 2005
' Adapted:       Bonnie Campbell, May 2014
' Revisions:     JRB, 5/26/2006 - updated gvar names, added gvarConnected
'                JRB, 7/7/2009  - removed gvarParentForm; added gvarWritePermission,
'                                 gvarHasAccessBE
'                --------------------------
'                BLC, 6/18/2014 - added public constants WATER_YEAR_START & WATER_YEAR_END
'                BLC, 7/31/2014 - changed db & user gvars to TempVars & initialized values
'                BLC, 8/6/2014  - switched order of setting globals & constants before sub
'                                 to ensure these load upon module being called for initGlobalTempVars
'                                 merged into mod_Initialize_App
'                --------------------------
' =================================

' ---------------------------------
' GLOBALS:      global variables
' Description:  variables provide globally accessible references for forms, controls
'               used to refresh objects after popup form updates
' References:   -
' Source/date:  John R. Boetsch, May 2005
' Adapted:      Bonnie Campbell, May 2014
' Revisions:    BLC, 7/31/2014 - initial version
' ---------------------------------
Public gvarRefForm As Form          ' referring form object
Public gvarRefCtl As Control        ' specific control on referring form
Public gvarRefTaxonCtl As Control   ' specific taxon control
Public gvarRefContactCtl As Control ' specific contacts control
Public blnRunQueries As Boolean     ' flag to indicate whether to run the queries upon opening
Public blnUpdateAll As Boolean      ' flag to indicate whether to run all queries

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, May 2014
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version (NCPN WQ Utilities Tool, WATER_YEAR_START & WATER_YEAR_END)
'               BLC, 4/22/2015 - adapted to generic tools (NCPN Invasives Reporting Tool) by adding
'                                USER_ACCESS_CONTROL (False - gives users full control in apps w/o controls,
'                                                     True - relies on user access control settings)
'                                DB_SYS_TABLES & APP_SYS_TABLES (handle table arrays for the database/
'                                   application)
'               BLC, 4/30/2015 - shifted to mod_App_Settings
' ---------------------------------

' ---------------------------------
' SUB:          SafeStart
' Description:  Sets error trapping/handling for database to ensure clear error trapping.
' Note:         Trapping is set to "Break in Class Module" (1) vs. "Break on Unhandled Errors" (1) since
'               the latter breaks on class calling code vs. class code. "Break on All Errors" (0) is not
'               used since this breaks even on handled errors.
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Luke Chung, date unkown
'               http://www.fmsinc.com/tpapers/vbacode/debug.asp
' Adapted:      Bonnie Campbell, July 7, 2015 for NCPN WQ Utilities tool
' Revisions:    BLC, 7/7/2015 - initial version
' ---------------------------------
Sub SafeStart()
On Error GoTo Err_Handler

  Application.SetOption "Error Trapping", 1
  

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SafeStart[mod_Initialize_App])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          initGlobalTempVars
' Description:  Initializes database TempVars which cannot be initialized outside of sub/function
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, July 31, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version
' ---------------------------------
Public Sub initGlobalTempVars()
On Error GoTo Err_Handler
Dim aryStdVars() As Variant
Dim i As Integer

    ' Global variables
    TempVars.Add "Connected", False     'Boolean flag -> back-end db connection is valid or not
    TempVars.Add "HasAccessBE", False   'Boolean flag -> app has one or more Access back-ends or not
    
    ' User access global variables
    TempVars.Add "WritePermission", False   'Boolean flag -> user has write privileges to the back-end db or not

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - initGlobalTempVars[mod_Initialize_App])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          initApp
' Description:  Initializes application variables
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, July 31, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Public Sub initApp()
On Error GoTo Err_Handler:

    ' Initialize global TempVars that require function
    initGlobalTempVars

    ' Application option settings
    Application.SetOption "Default Font Name", "Arial"
    Application.SetOption "Default Font Size", 9
    Application.SetOption "Auto Compact", True

    If DEV_MODE = False Then
        ' Turn off options (only apparent after the next time app is opened)
        CurrentDb.Properties("AllowFullMenus") = False
        CurrentDb.Properties("AllowShortcutMenus") = False
        CurrentDb.Properties("AllowBuiltInToolbars") = False
    End If
    
    'Check for missing tables
    If SysTablesExist("db") = False Then Exit Sub

    ' Verify the back-end database connections, and run the setup function if okay
    VerifyConnections
    If TempVars("Connected") Then AppSetup

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - initApp[mod_Initialize_App])"
    End Select
    Resume Exit_Procedure
End Sub

' =================================
' FUNCTION:     AppSetup
' Description:  Confirms required tables, determines if application version is current, and
'                   reset the switchboard / application mode based on user privileges upon
'                   first opening the application and just after relinking the back-end dbs
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   BEUpdates, IsODBC, LinkedDatabase, TableExists,
'                   TestODBCConnection
' Source/date:  John R. Boetsch, 7/9/2009
' Revisions:    JRB, 7/27/2009 - added a check on whether the application version was added
'                   by fxnBEUpdates, reordered caption setting statements
'               JRB, 12/14/2009 - changed to allow db window access for power users
'               JRB, 1/11/2010 - added a line to make cmdBackup visible if Access back-end
'               BLC, 6/12/2014 - revised to set TempVars.Item("UserAccessLevel") vs. cAppMode
'                                TempVars available throughout app w/o setting cAppMode subform control
'               BLC, 7/31/2014 - changed gvars to TempVars, moved to mod_Initialize_App,
'                                revised to iterate missing system table check
'               BLC, 8/6/2014  - moved switchboard control settings based on user access to setUserAccess
'                                removed unused varRole
'               BLC, 8/25/2014 - added setUserAccess "update" flag for refreshing UI settings
'               BLC, 4/30/2015 - added DB_ADMIN_CONTROL and MAIN_APP_FORM checks for handling apps w/o full Db_Admin subform
'                                to set strReleaseID and strAddress values
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 5/28/2015 - added MAIN_APP_FORM open check to prevent Error #2450 where
'                                frm_Tgt_List_Tool is not found on exit from frm_Connect_Dbs
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
' =================================
Public Function AppSetup()
    On Error GoTo Err_Handler

    Dim frm As Form
    Dim strSysTable As String, strAddress As String, strUser As String, strRelease As String
    Dim strSQL As String, strCaption As String, strReleaseID As String

    If Not FormIsOpen(MAIN_APP_FORM) Then
        DoCmd.OpenForm MAIN_APP_FORM, acNormal, , , , acHidden
    End If

    Set frm = Forms(MAIN_APP_FORM) 'Forms!frm_Switchboard
    TempVars("WritePermission") = False
    
    If DB_ADMIN_CONTROL Then
        strReleaseID = APP_RELEASE_ID
        strAddress = APP_URL
    Else
        strReleaseID = frm!Release_ID
        strAddress = frm!Web_address
    End If
    
    ' Check for required system tables
    If SysTablesExist("app") = False Then GoTo Exit_Procedure

    ' Confirm that the application version is supported
    Select Case DLookup("Is_supported", "tsys_App_Releases", _
            "[Release_ID] = """ & strReleaseID & """")
      Case 0    ' Application not supported
        If MsgBox("This version of the front-end application is out of date ... " _
            & vbCrLf & " ... a more recent version is available!" _
            & vbCrLf & vbCrLf & "Would you like to download the most recent version now?", _
            vbYesNo, "Database Application Update Available") = vbYes Then
            
            If IsNull(strAddress) Then
                MsgBox "Web address not found - contact the Data Manager"
            Else
                Application.FollowHyperlink strAddress, , True, False
                MsgBox "You may replace this front-end file with the new download ..."
            End If
        End If
        ' Exit the application as it is not supported
        DoCmd.Quit acQuitSaveNone

      Case 1    ' Application is supported but not the most current release
        If MsgBox("An updated version of the front-end application is available!" _
            & vbCrLf & vbCrLf & "Would you like to download the most recent version now?", _
            vbYesNo, "Database Application Update Available") = vbYes Then
            
            If IsNull(strAddress) Then
                MsgBox "Web address not found - contact the Data Manager"
            Else
                Application.FollowHyperlink strAddress, , True, False
                MsgBox "You may replace this front-end file with the new download ..."
                ' Exit the application only if they download a new copy
                DoCmd.Quit acQuitSaveNone
            End If
        End If

      Case Else  ' Application is current, do nothing
    End Select

    ' Determine the application mode (user access level) according to the user role
    setUserAccess frm, "update"

    ' Log the user, login time, release number, and application mode in the systems table
    strRelease = Left(strReleaseID, 8) & " / " & TempVars("UserAccessLevel")
    If IsODBC("tsys_Logins") Then
        ' Use a pass-through query to test the connection for write privileges
        strSQL = "INSERT INTO dbo.tsys_Logins " & _
            "SELECT GETDATE() AS Time_stamp, '" & strUser & "' AS User_name, '" & _
            strRelease & "' AS Action_taken"
        TempVars("WritePermission") = TestODBCConnection("tsys_Logins", , strSQL, False)
        ' Notify the user if their back-end privileges are insufficient to use the application
        If TempVars("WritePermission") = False And TempVars("UserAccessLevel") <> "read only" Then
            MsgBox "Your login does not have modify privileges to the database." & _
                vbCrLf & "Notify the database administrator before using this application." _
                & vbCrLf & vbCrLf & "User: " & strUser & vbCrLf & "Db:   " & _
                LinkedDatabase("tsys_Logins")
        End If
    Else
        TempVars("WritePermission") = True
        strSQL = "INSERT INTO tsys_Logins ( User_name, Action_taken ) SELECT '" _
            & strUser & "' AS User, """ & strRelease & """ AS Action;"
        DoCmd.SetWarnings False
        DoCmd.RunSQL strSQL     ' Will throw a trapped error if no write permissions
        DoCmd.SetWarnings True
    End If

    ' If the current front-end release is not listed in the back-end file, run fxn to update
    '   Note: Needed where there are one or more back-end copies at remote locations that
    '   cannot be updated with new release information by the developer
    If DCount("*", "tsys_App_Releases", "[Release_ID]=""" & strReleaseID & """") = 0 Then
        If TempVars("WritePermission") Then BEUpdates (True)
        ' Check once more to make sure that the release was added properly - if not notify
        If DCount("*", "tsys_App_Releases", "[Release_ID]=""" & strReleaseID & """") = 0 Then
            MsgBox "Unable to determine the application version." & vbCrLf & vbCrLf & _
                "Please notify the database administrator.", , "Application error"
            ' Skip the code to set the caption
            GoTo Update_Settings
        End If
    ' Or run updates only on new update lines (avoids issuing a new version for minor updates)
    ElseIf DCount("*", "tsys_BE_Updates", "[Is_done]=False") > 0 Then
        If TempVars("WritePermission") Then BEUpdates (False)
    End If

    ' Set the table-driven caption of the switchboard
    strCaption = DLookup("[Database_title]", "tsys_App_Releases", "[Release_ID] = '" _
        & frm!Release_ID & "'")
    frm.Caption = strCaption

Exit_Procedure:
    DoCmd.SetWarnings True
    Exit Function

Update_Settings:
    ' Update the switchboard settings according to application mode
    setUserAccess frm, "update"

    'if DbAdmin subform is complete, then continue
    If DB_ADMIN_CONTROL Then
        ' If there is an Access back-end, open the always-open form (to maintain a connection
        '   to the back-end and avoid unnecessary create/delete/updates to its .ldb lock file)
        If TempVars("HasAccessBE") Then DoCmd.OpenForm "frm_Lock_BE", , , , , acHidden
    
        ' If there is an Access back-end, make the backups button visible
        frm!fsub_DbAdmin.Form!cmdBackup.visible = TempVars("HasAccessBE")
    
        ' Requery the control that shows the linked back-ends
        frm!lbxLinkedDbs.Requery
    
        Resume Exit_Procedure
    End If

Err_Handler:
    Select Case Err.Number
      Case 3073 ' Operation must use updateable query - back-end is read-only
        MsgBox "The back-end file is set to read-only, or is located in" & vbCrLf & _
            "a folder for which you do not have modify privileges." & vbCrLf & vbCrLf & _
            "Close the application and uncheck the read-only box in the" & vbCrLf & _
            "file properties window before using this application.", vbCritical, _
            "Application error (#" & Err.Number & " - AppSetup[mod_Initialize_App])"
        TempVars("WritePermission") = False
      Case 3078   ' Can't find the system table
        MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "Application error (#" & Err.Number & " - AppSetup[mod_Initialize_App])"
      Case 2001   ' Field name in DLookup improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, _
            "Application error (#" & Err.Number & " - AppSetup[mod_Initialize_App])"
      Case 94    ' Missing information in the systems table
        MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
            vbCrLf & "the database administrator before using this application.", _
            vbCritical, "Application error (#" & Err.Number & " - AppSetup[mod_Initialize_App])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AppSetup[mod_Initialize_App])"
    End Select
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     SysTablesExist
' Description:  Checks if select system tables exist
' Assumptions:  -
' Parameters:   tblType - string value representing the group of tables to check
'                         either "db" -> backend data tables, links & app defaults
'                         or    "app" -> release version, bugs, user roles & logins
' Returns:      True if all tables exist, false if any do not
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, July 31, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version
'               BLC, 4/22/2015 - shifted default arrays of sys db & app tables to globals
'                                DB_SYS_TABLES & APP_SYS_TABLES to accommodate & expose settings for
'                                multiple apps (NCPN Invasives Reporting tool), some that do not
'                                contain same/all tables
' ---------------------------------
Public Function SysTablesExist(tblType As String) As Boolean
On Error GoTo Err_Handler:
Dim sysTables As Variant
Dim i As Integer
Dim missingTable As String

    Select Case tblType
            
        Case "db"
            ' Confirm certain system tables exist --> if not, close the application
            '-----------------------------------------------------------------------
            '   tsys_App_Defaults -> default application settings
            '   tsys_BE_Updates   -> updates to post to remot back-end copies
            '   tsys_Link_Dbs     -> info about linked back-end dbs
            '   tsys_Link_Tables  -> info about linked tables
            '-----------------------------------------------------------------------
            sysTables = Split(DB_SYS_TABLES, ",")

        Case "app"
            ' Confirm certain backend system tables exist --> if not, set connected to false
            '-----------------------------------------------------------------------
            '   tsys_App_Releases -> list of application releases
            '   tsys_Bug_Reports  -> tracking for known issues
            '   tsys_Logins       -> system use monitoring
            '   tsys_User_Roles   -> assign user access priviledges
            '-----------------------------------------------------------------------
            sysTables = Split(APP_SYS_TABLES, ",")
        Case ""
    End Select
        
    For i = 0 To UBound(sysTables)
        If TableExists("tsys_" & Trim(sysTables(i))) = False Then
            missingTable = sysTables(i)
            GoTo Missing_Table:
        End If
    Next
    
    'return result
    SysTablesExist = True
    
Exit_Function:
    Exit Function
    
Missing_Table:
    Dim strMsg As String
    strMsg = "Unable to find the system table: " & vbCrLf & vbCrLf & vbTab & _
                sysTables(i) & vbCrLf & vbCrLf

    Select Case missingTable
        Case "App_Defaults", "BE_Updates", "Link_Dbs", "Link_Tables", "Link_Files"
            strMsg = strMsg & "Notify the database administrator."
            DoCmd.SetWarnings True
            DoCmd.Quit acQuitSaveNone
        Case "App_Releases", "Bug_Reports", "Logins", "User_Roles"
            ' Close the application if missing one or more systems tables
            strMsg = strMsg & _
                "Either link to the correct back-end or quit and notify the" & vbCrLf & _
                "database administrator before using this application."
            TempVars("Connected") = False
        Case ""
    End Select
    
    'display messages
    MsgBox strMsg, vbCritical, "Application error - Missing system table"
    
    'return result
    SysTablesExist = False

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SysTablesExist[mod_Initialize_App])"
    End Select
    Resume Exit_Function
End Function