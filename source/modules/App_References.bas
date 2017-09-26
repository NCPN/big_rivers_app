Option Compare Database
Option Explicit

' =================================
' MODULE:       App_References
' Level:        Framework module
' Version:      1.00
' Description:  generic reference loading functions & procedures
'
' Source/date:  Bonnie Campbell, 9/19/2017
' Revisions:    BLC - 9/19/2017 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------

' ---------------------------------
' Function:     PowerShellExists
' Description:  Indicates whether powershell is installed or not
' Assumptions:  The current machine is the location to check. Remote locations not supported.
' Parameters:   -
' Returns:      True or False depending upon whether PowerShell is installed on the system
' Throws:       none
' References:
'   Mark K, November 3, 2015
'   https://www.access-programmers.co.uk/forums/showthread.php?t=282167
' Source/date:  Bonnie Campbell, September 12, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/12/2017 - initial version
' ---------------------------------
Public Function PowerShellExists() As Boolean
On Error GoTo Err_Handler

    Const HKEY_LOCAL_MACHINE = &H80000002
    Const PS_KEY As String = "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine"
    
    Dim oReg As Object
    Dim sComputer As String
    Dim sValue As String
    
    sComputer = "."
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sComputer & "\root\default:StdRegProv")
    oReg.GetStringValue HKEY_LOCAL_MACHINE, PS_KEY, "RuntimeVersion", sValue
    
    MsgBox sValue


Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PowerShellExists[mod_PowerShell])"
    End Select
    Resume Exit_Handler
End Function