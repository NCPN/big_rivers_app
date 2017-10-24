Option Compare Database
Option Explicit

' =================================
' MODULE:       App_Settings
' Level:        Application module
' Version:      1.01
' Description:  Application-wide related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, September 2017
' Revisions:    BLC, 9/19/2015  - 1.00 - initial version
'               BLC, 10/18/2017 - 1.01 - added CREATE_ENUMS for turning ON/OFF enum creation from table
' =================================

' ---------------------------------
' GLOBALS:      global values set for application
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, June 2016
' Adapted:      -
' Revisions:    BLC, 6/6/2016 - initial version (NCPN WQ Big Rivers App, App_Templates)
' ---------------------------------

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, September 2017
' Adapted:      -
' Revisions:    BLC, 9/19/2017 - initial version
'               BLC, 10/18/2017 - added CREATE_ENUMS for turning ON/OFF enum creation from table
' ---------------------------------

'-----------------------------------------------------------------------
' Application
'-----------------------------------------------------------------------
Public Const app As String = "Big_Rivers"                       'String -> application

'-----------------------------------------------------------------------
' Reference Loading
'-----------------------------------------------------------------------
Public Const LOAD_REFERENCES As Boolean = True                  'Boolean -> whether references should
                                                                '           be loaded into the current db on open

'-----------------------------------------------------------------------
' Enum Creation
'-----------------------------------------------------------------------
Public Const CREATE_ENUMS As Boolean = False                    'Boolean -> whether enums should
                                                                '           be created from the enum table