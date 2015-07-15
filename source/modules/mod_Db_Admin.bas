Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Db_Admin
' Level:        Framework module
' Version:      1.00
' Description:  Database admin related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
' =================================

' ---------------------------------
' SUB:          initializeControls
' Description:  set initial control values
' Parameters:   frm - form to initialize (form object)
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, Sept 2014 for NCPN tools
' Adapted:      -
' Revisions:    BLC - 9/01/2014 - initial version
' ---------------------------------
Public Sub initializeControls(frm As Form)
    On Error GoTo Err_Handler
    Dim aryCtrls() As Variant
    Dim ctrlName As String, tgtCtrlName As String
    Dim i As Integer

    With frm
        Select Case .name
            Case "frm_Set_Defaults"
                'TempVars not yet populated -> use fsub_DbAdmin control defaults
                aryCtrls = Array("User", "Project", "GPS_model", "Park", "Datum", "Declination", "Timeframe", "Project")
                For i = 0 To UBound(aryCtrls)
                    ctrlName = "tbx" & aryCtrls(i)
                    If aryCtrls(i) = "Declination" Or _
                       aryCtrls(i) = "Timeframe" Or _
                       aryCtrls(i) = "Project" Then
                        tgtCtrlName = "tbx" & aryCtrls(i)
                    Else
                        tgtCtrlName = "cbx" & aryCtrls(i)
                    End If
                    .Controls(tgtCtrlName) = Forms!frm_Switchboard.fsub_DbAdmin.Form.Controls(ctrlName).Value
                Next
        End Select
    End With
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - initializeControls[mod_Common_UI])"
    End Select
    Resume Exit_Procedure
End Sub