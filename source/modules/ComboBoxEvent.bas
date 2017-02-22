Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        ComboBoxEvent
' Level:        Framework class
' Version:      1.00
'
' Description:  ComboBoxEvent object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 2/16/2017
' References:   -
' Revisions:    BLC - 2/16/2017 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private WithEvents ComboBoxEvent As Office.CommandBarComboBox
Attribute ComboBoxEvent.VB_VarHelpID = -1

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------

'======== Standard Methods ===========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' References:
'   Microsoft, Unknown
'   https://msdn.microsoft.com/en-us/library/office/aa170937(v=office.11).aspx
' Source/Date:  Bonnie Campbell
' Adapted:      -
' Revisions:
'   BLC, 2/16/2017 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_ComboBoxEvent])"
    End Select
    Resume Exit_Handler

End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' References:
'   Microsoft, Unknown
'   https://msdn.microsoft.com/en-us/library/office/aa170937(v=office.11).aspx
' Source/Date:  Bonnie Campbell - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/16/2017 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    Set ComboBoxEvent = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_ComboBoxEvent])"
    End Select
    Resume Exit_Handler

End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SyncBox
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Microsoft, Unknown
'   https://msdn.microsoft.com/en-us/library/office/aa170937(v=office.11).aspx
' Source/Date:  Bonnie Campbell - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/16/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SyncBox(box As Office.CommandBarComboBox)
On Error GoTo Err_Handler
    
    Set ComboBoxEvent = box

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SyncBox[cls_ComboBoxEvent])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          ComboBoxEvent_Change
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Microsoft, Unknown
'   https://msdn.microsoft.com/en-us/library/office/aa170937(v=office.11).aspx
' Source/Date:  Bonnie Campbell - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 2/16/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub ComboBoxEvent_Change(ByVal ctrl As Office.CommandBarComboBox)
On Error GoTo Err_Handler
    
    'Dim strComboText As String
    
    'strComboText = ctrl.text

'    For Each PhotoType In Split(PHOTO_TYPES_MAIN, ",")
'         PhotoType 'Left(PhotoType, 1)
'    Next
'    For Each PhotoType In Split(PHOTO_TYPES_OTHER, ",")
'        '"Other - " & PhotoType '"O" & Left(PhotoType, 1)
'    Next
    Debug.Print "in cbxevent change " & ctrl.Text
    
    Select Case ctrl.Text 'strComboText
        Case "Reference"
        Case "Overview"
        Case "Feature"
        Case "Transect"
        Case "Other - Animal"
        Case "Other - Plant"
        Case "Other - Cultural"
        Case "Other - Disturbance"
        Case "Other - Field Work"
        Case "Other - Scenic"
        Case "Other - Weather"
        Case "Other - Other"
        Case "Unclassified"
    End Select

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - ComboBoxEvent_Change[cls_ComboBoxEvent])"
    End Select
    Resume Exit_Handler
End Sub