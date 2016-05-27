Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Forms
' Level:        Framework module
' Version:      1.03
' Description:  generic form functions & procedures
'
' Source/date:  Bonnie Campbell, 2/19/2015
' Revisions:    BLC - 2/19/2015 - 1.00 - initial version
'               BLC - 5/18/2015 - 1.01 - fixed ClearFields documentation
'               BLC - 6/9/2015  - 1.02 - added CloseFormsReports()
'               BLC - 6/30/2015 - 1.03 - shifted to mod_UI: ChangeBackColor
'                                        shifted from mod_UI: FormIsOpen, FormIsLoaded, SwitchboardIsOpen
'                                        shifted to mod_App_UI: ClearFields
' =================================

'=================================================================
'  References
'=================================================================
' ---------------------------------
'  Access Control Types
' ---------------------------------
' dbtech1, March 13, 2008
' http://www.utteraccess.com/forum/control-type-vba-t1609220.html
' 126 - acAttachment         119 - acCustomControl  114 - acObjectFrame    101 - acRectangle
' 108 - acBoundObjectFrame   103 - acImage          105 - acOptionButton   112 - acSubform
' 106 - acCheckBox           100 - acLabel          107 - acOptionGroup    123 - acTabCtl
' 111 - acComboBox           102 - acLine           124 - acPage           109 - acTextBox
' 104 - acCommandButton      110 - acListBox        118 - acPageBreak      122 - acToggleButton
' ---------------------------------

' ---------------------------------
'  Access Form Sections
' ---------------------------------
'   acDetail        0   (Default) Detail section    acGroupLevel1Footer 6   Group-level 1 footer (reports only)
'   acFooter        2   Form or report footer       acGroupLevel1Header 5   Group-level 1 header (reports only)
'   acHeader        1   Form or report header       acGroupLevel2Footer 8   Group-level 2 footer (reports only)
'   acPageFooter    4   Page footer                 acGroupLevel2Header 7   Group-level 2 header (reports only)
'   acPageHeader    3   Page header
' ---------------------------------

' ---------------------------------
'  Access Backstyle Property
' ---------------------------------
'  Transparent  0           Normal  1
' ---------------------------------

'=================================================================
'  Declarations
'=================================================================
Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As _
     Integer
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As _
     Integer

'=================================================================
'  Properties
'=================================================================


'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
' FUNCTION:     CloseFormsReports
' Description:  close forms, reports
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Susan Harkins, July 21, 2009
'   http://www.techrepublic.com/blog/microsoft-office/automatically-close-all-the-open-forms-and-reports-in-an-access-database/
' Adapted:      Bonnie Campbell, June 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/9/2015  - initial version
' ---------------------------------
Public Function CloseFormsReports()
On Error GoTo Err_Handler

    'Close all open forms
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).Name
    Loop
    
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).Name
    Loop

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CloseFormsReports[mod_Forms])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     FormIsOpen
' Description:  Indicates whether or not the specific form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/5/2006 as fxnSwitchboardIsOpen
' Adapted:      Bonnie Campbell, 4/30/2015 for NCPN tools
' Revisions:    BLC, 4/30/2015 - initial version
' =================================
Public Function FormIsOpen(strFormName As String) As Boolean
    On Error GoTo Err_Handler

    Dim frm As Form

    FormIsOpen = False    ' Default in case of error
 
    'search for form in Forms collection (all open forms)
    For Each frm In Forms
      If frm.Name = strFormName Then
        'check form is in Form view: 0 - Design View, 1 - Form View, 2 - Datasheet View
        If frm.CurrentView = 1 Then
            FormIsOpen = True
            'Exit Function
        End If
      End If
    Next

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsOpen[mod_UI])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     SwitchboardIsOpen
' Description:  Indicates whether or not the switchboard form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/5/2006
' Revisions:    JRB, 5/5/2006 - initial version
'               BLC, 4/30/2015  - moved to mod_Db framework module from mod_Custom_Functions
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function SwitchboardIsOpen() As Boolean
    On Error GoTo Err_Handler

    SwitchboardIsOpen = False    ' Default in case of error

    Dim strSwitchboardName As String

    strSwitchboardName = "frm_Switchboard"

    'check for switchboard in all open forms ( AllForms.IsLoaded() )
    If CurrentProject.AllForms(strSwitchboardName).IsLoaded = True Then
        If CurrentProject.AllForms(strSwitchboardName).CurrentView = 1 Then
            SwitchboardIsOpen = True
        End If
    End If

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SwitchboardIsOpen[mod_UI])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     FormIsLoaded
' Description:  Returns whether the specified form is loaded in Form or Datasheet view
' Parameters:   strFormName - string for the name of the form to check
' Returns:      True if the specified form is open in Form view or Datasheet view
' Throws:       none
' References:   none
' Source/date:  From Northwind sample database, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_UI
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function FormIsLoaded(ByVal strFormName As String) As Integer
    On Error GoTo Err_Handler
 
    ' These variables are used to test the return values of the SysCmd function
    '  and the CurrentView property of the requested form.
    Const cObjStateClosed = 0
    Const cDesignView = 0

    ' Use the SysCmd function to check the current state of the requested form.
    '  Possible states: not open or nonexistent, open, new, or changed but not saved
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> cObjStateClosed Then
        ' Checks for the current view of the requested form, assuming the previous statement
        '   found it to be open ... return True if open and not in design view
        If Forms(strFormName).CurrentView <> cDesignView Then
            FormIsLoaded = True
        End If
    End If
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsLoaded[mod_UI])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          AddControl
' Description:  initialize application values
' Assumptions:  -
' Parameters:   frm - form (object)
'               ctrl - control (object)
'               ctrlName - name of control (string)
'               xPos - horizontal position (twips)
'               yPos - vertical position (twips)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' meloncolly, October 27, 2006
' http://forums.aspfree.com/microsoft-access-help-18/add-controls-form-dynamically-139627.html
' https://msdn.microsoft.com/en-us/library/bb237827(office.12).aspx
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
' ---------------------------------
Public Sub AddControl(frm As Form, ctrl As Control, ctrlName As String, _
                        xPos As Integer, yPos As Integer)
On Error GoTo Err_Handler

    ' Create ctrl
    Set ctrl = CreateControl(frm.Name, ctrl.ControlType, , "", "", xPos, yPos)
    
    ' Restore form
    DoCmd.Restore

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddControl[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ContinuousUpDown
' Description:  Respond to Up/Down in a continuous form by moving to next record
' Assumptions:  Active control's EnterKeyBehavior is OFF
' Usage:        Call ContinuousUpDown(Me, KeyCode)
' Parameters:   frm - form for key behavior
'               KeyCode - code for key being pressed (integer)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Allen Browne via Jeanette Cunningham, Apr 13, 2010
' http://www.pcreview.co.uk/threads/need-to-get-the-up-down-arrow-keys-working.3995845/
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015  - initial version
' ---------------------------------
Public Sub ContinuousUpDown(frm As Form, KeyCode As Integer)
On Error GoTo Err_Handler

    Dim strForm As String
    
    strForm = frm.Name
    
    'determine key being used
    Select Case KeyCode
        Case vbKeyUp
            If ContinuousUpDownOk Then
                
                'Save any edits
                If frm.Dirty Then
                    RunCommand acCmdSaveRecord
                End If
                
                'Go previous: error if already there.
                    RunCommand acCmdRecordsGoToPrevious
                KeyCode = 0 'Destroy the keystroke
            End If
    
    Case vbKeyDown
        If ContinuousUpDownOk Then
            
            'Save any edits
            If frm.Dirty Then
                frm.Dirty = False
            End If
            
            'Go to the next record, unless at a new record.
            If Not frm.NewRecord Then
                RunCommand acCmdRecordsGoToNext
            End If
            KeyCode = 0 'Destroy the keystroke
        End If
    End Select

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 2046, 2101, 2113, 3022, 2465 'Already at first record, or save
            'failed, or The value you entered isn't valid for this field.
            KeyCode = 0
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - ContinuousUpDown[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ContinuousUpDownOk
' Description:  Suppress moving up/down a record in a continuous form if:
'                - control is not in the Detail section
'                - multi-line text box (vertical scrollbar or EnterKeyBehavior true)
' Assumptions:  Active control's EnterKeyBehavior is OFF
' Usage:        Called by ContinuousUpDown SUB
' Parameters:   N/A
' Returns:      boolean - true if moving up/down a record in continuous form is ok, false if not
' Throws:       none
' References:   none
' Source/date:
' Allen Browne via Jeanette Cunningham, Apr 13, 2010
' http://www.pcreview.co.uk/threads/need-to-get-the-up-down-arrow-keys-working.3995845/
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015  - initial version
' ---------------------------------
Private Function ContinuousUpDownOk() As Boolean
On Error GoTo Err_Handler
    Dim blnDontDoIt As Boolean
    Dim ctl As Control
    
    Set ctl = Screen.ActiveControl
    If ctl.Section = acDetail Then
        If TypeOf ctl Is TextBox Then
            blnDontDoIt = ((ctl.EnterKeyBehavior) Or (ctl.ScrollBars > 1))
        End If
    Else
        blnDontDoIt = True
    End If

Exit_Function:
    ContinuousUpDownOk = Not blnDontDoIt
    Set ctl = Nothing

Exit Function

Err_Handler:
    Select Case Err.Number
        Case 2474 'No active control
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - ContinuousUpDownOk[mod_Forms])"
    End Select
    Resume Exit_Function
End Function