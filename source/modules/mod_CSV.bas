Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_CSV
' Level:        Framework module
' Version:      1.00
' Description:  Framework-wide related mathematical values, functions & subroutines
'
' Source/date:  Bonnie Campbell, September 30, 2016 for NCPN tools
' Revisions:    BLC, 9/30/2016 - initial version
' =================================

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:   -
' Source/date:  Bonnie Campbell, September 30, 2016 for NCPN tools
' Adapted:      -
' Revisions:    BLC, 9/30/2016 - initial version
' ---------------------------------

'-----------------------------------------------------------------------
' Constants
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' Declarations
'-----------------------------------------------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

'-----------------------------------------------------------------------
' Functions
'-----------------------------------------------------------------------

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------

' ---------------------------------
' SUB:          ImportCSV
' Description:  CSV import actions
' Assumptions:  If DeleteExistingTable is False, data will append to table if possible
' Parameters:   strPath - CSV full file path (string)
'               strTable - table to insert data into (string)
'               HasHeaders - whether CSV first row is a header row
'                            (boolean, optional, default = true)
'               DeleteExistingTable - whether table should be deleted first
'                                     (boolean, optional, default = true)
' Returns:      -
' Throws:       none
' References:   -
'
'
' Source/date:  Bonnie Campbell, September 30, 2016 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 9/30/2015 - initial version
' ---------------------------------
Public Sub ImportCSV(strPath As String, strTable As String, _
                    Optional HasHeaders As Boolean = True, _
                    Optional DeleteExistingTable As Boolean = True)
On Error GoTo Err_Handler

    'remove existing table --> otherwise append
    If DeleteExistingTable Then
        If TableExists(strTable) Then _
            DoCmd.DeleteObject acTable, strTable
    End If

    DoCmd.TransferText acImportDelim, , strTable, strPath, HasHeaders

Exit_Handler:
    'cleanup
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ImportCSV[mod_CSV])"
    End Select
    Resume Exit_Handler
End Sub