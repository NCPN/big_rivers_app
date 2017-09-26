Option Compare Database
Option Explicit

' =================================
' MODULE:       App_Settings
' Level:        Application module
' Version:      1.00
' Description:  Application-wide related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, September 2017
' Revisions:    BLC, 9/19/2015  - 1.00 - initial version
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
' ---------------------------------

'-----------------------------------------------------------------------
' Application
'-----------------------------------------------------------------------
Public Const APP As String = "Big_Rivers"                       'String -> application

'-----------------------------------------------------------------------
' Reference Loading
'-----------------------------------------------------------------------
Public Const LOAD_REFERENCES As Boolean = True                  'Boolean -> whether references should
                                                                '           be loaded into the current db on open