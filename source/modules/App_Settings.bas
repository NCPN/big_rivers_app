Option Compare Database
Option Explicit

' =================================
' MODULE:       App_Settings
' Level:        Application module
' Version:      1.02
' Description:  Application-wide related values, functions & subroutines
'
' Source/date:  Bonnie Campbell, September 2017
' Revisions:    BLC, 9/19/2015  - 1.00 - initial version
'               BLC, 10/18/2017 - 1.01 - added CREATE_ENUMS for turning ON/OFF enum creation from table
'               BLC, 11/2/2017 - 1.02 - added MAX_PLOT_NUMBER, MAX_TRANSECT_NUMBER for handling max
'                                       numbers retrieved
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
'               BLC, 11/2/2017 - added MAX_PLOT_NUMBER, MAX_TRANSECT_NUMBER for handling max
'                                       numbers retrieved
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

'-----------------------------------------------------------------------
' Enum Creation
'-----------------------------------------------------------------------
Public Const CREATE_ENUMS As Boolean = False                    'Boolean -> whether enums should
                                                                '           be created from the enum table

'-----------------------------------------------------------------------
' Maximum Numbers
'-----------------------------------------------------------------------
Public Const MAX_PLOT_NUMBER As Integer = 250                   'Integer -> highest plot # in protocol
Public Const MAX_TRANSECT_NUMBER As Integer = 8                 'Integer -> highest transect # in protocol