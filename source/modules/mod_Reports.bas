Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Reports
' Level:        Framework module
' Version:      1.00
' Description:  generic report functions & procedures
'
' Source/date:  Bonnie Campbell, 5/25/2016
' Revisions:    BLC - 5/25/2016 - 1.00 - initial version
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
' Function:     NoData
' Description:  report actions when no data is found
' Assumptions:  -
' Parameters:   rpt - report being referenced
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 10, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/10/2015 - initial version
' ---------------------------------
Public Function NoData(rpt As Report)
On Error GoTo Err_Handler

    'Purpose: Called by report's NoData event.
    'Usage: =NoData([Report])
    Dim strCaption As String   'Caption of report.
    
    strCaption = rpt.Caption
    If strCaption = vbNullString Then
        strCaption = rpt.Name
    End If
    
    DoCmd.CancelEvent
    MsgBox "There are no records to include in report """ & _
        strCaption & """.", vbInformation, "No Data..."


Exit_Function:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NoData[mod_Reports])"
    End Select
    Resume Exit_Function
End Function