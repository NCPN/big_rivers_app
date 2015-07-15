Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Settings
' Level:        Application module
' Version:      1.00
' Description:  Application-wide related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - initial version
' =================================

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
'               BLC, 4/30/2015 - add DB_ADMIN_CONTROL flag to handle applications w/o full DbAdmin subform & controls
'                                add MAIN_APP_FORM constant to handle applications where frm_Switchboard is NOT the main form
'                                add APP_RELEASE_ID constant to handle application release ID w/o full DbAdmin subfrom & controls
'               BLC, 5/1/2015  - add DEV_MODE constant to enable menus typically off during use
'               BLC, 5/13/2015 - shifted UI enable/disabled colors from TempVars set in initialize (mod_App_UI) to constants
'               BLC, 5/19/2015 - added FIX_LINKED_DBS flag to handle applications which require updates of tbl_Dbs via FixLinkedDb
'                                (usually when DbAdmin is not fully implemented)
'               BLC, 5/28/2015 - added MAIN_APP_MENU to handle applications w/ main menu forms (not tabbed switchboards)
' ---------------------------------
Public Const USER_ACCESS_CONTROL As Boolean = False             'Boolean flag -> db includes user access control or not
Public Const DB_ADMIN_CONTROL As Boolean = False                'Boolean flag -> db does not include DbAdmin subform & controls
Public Const FIX_LINKED_DBS As Boolean = True                   'Boolean flag -> db requires tbl_Dbs to be updated via FixLinkedDb (usually when DbAdmin is not fully implemented)
Public Const MAIN_APP_FORM As String = "frm_Tgt_List_Tool"      'String -> main tabbed form (frm_Switchboard, etc.)
Public Const MAIN_APP_MENU As String = "frm_Main_Menu"          'String -> main tabbed form (frm_Switchboard, etc.)
Public Const APP_RELEASE_ID As String = ""                      'String -> release ID (tsys_App_Release.Release_ID) for current release
                                                                '          used when db doesn't include full DbAdmin subform & controls, otherwise NULL
Public Const APP_URL As String = "science.nature.nps.gov/im/units/ncpn/datamanagement.cfm"
                                                                'String -> website URL for application
                                                                '          used when db doesn't include full DbAdmin subform & controls, otherwise NULL
Public Const DEV_MODE As Boolean = True                         'Boolean flag -> enable menus when typically they'd be OFF

'-----------------------------------------------------------------------
' Database System Tables
'-----------------------------------------------------------------------
'   Array("App_Defaults", "BE_Updates", "Link_Dbs", "Link_Tables")
'   tsys_App_Defaults -> default application settings
'   tsys_BE_Updates   -> updates to post to remot back-end copies
'   tsys_Link_Dbs     -> info about linked back-end dbs
'   tsys_Link_Tables  -> info about linked tables
'-----------------------------------------------------------------------
' Application Backend System Tables
'-----------------------------------------------------------------------
'   Array("App_Releases", "Bug_Reports", "Logins", "User_Roles")
'   tsys_App_Releases -> list of application releases
'   tsys_Bug_Reports  -> tracking for known issues
'   tsys_Logins       -> system use monitoring
'   tsys_User_Roles   -> assign user access priviledges
'-----------------------------------------------------------------------
' SEE ALSO >>>> SysTablesExist() function
'-----------------------------------------------------------------------
Public Const DB_SYS_TABLES As String = "App_Defaults, Link_Files, Link_Tables"
Public Const APP_SYS_TABLES As String = ""

'-----------------------------------------------------------------------
' User Interface Colors
'-----------------------------------------------------------------------
'std control colors
Public Const CTRL_DISABLED As Long = lngLtGray
Public Const CTRL_ADD_ENABLED As Long = lngLime
Public Const CTRL_REMOVE_ENABLED As Long = lngLtOrange
Public Const TEXT_ENABLED As Long = lngBlue
Public Const TEXT_DISABLED As Long = lngGray

Public Const PROGRESS_BAR As Long = lngLime