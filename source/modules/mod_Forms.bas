Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Forms
' Level:        Framework module
' Version:      1.07
' Description:  generic form functions & procedures
'
' Source/date:  Bonnie Campbell, 2/19/2015
' Revisions:    BLC - 2/19/2015 - 1.00 - initial version
'               BLC - 5/18/2015 - 1.01 - fixed ClearFields documentation
'               BLC - 6/9/2015  - 1.02 - added CloseFormsReports()
'               BLC - 6/30/2015 - 1.03 - shifted to mod_UI: ChangeBackColor
'                                        shifted from mod_UI: FormIsOpen, FormIsLoaded, SwitchboardIsOpen
'                                        shifted to mod_App_UI: ClearFields
'               BLC - 6/1/2016  - 1.04 - added SetFormOpacity(), CaptureEscapeKey(), constants & functions
'                                        from Uplands mod_App_UI
'               BLC - 6/24/2016 - 1.05 - added ToggleForm(), replaced Exit_Function > Exit_Handler
'               BLC - 7/1/2016  - 1.06 - added font weight constants
'               BLC - 7/28/2016 - 1.07 - added clearing lblMsg caption for ClearForm()
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

' ---------------------------------
'  Access FontWeight Property
' ---------------------------------
'   Thin    100         Extra Light         200
'   Light   300         (Default) Normal    400
'   Medium  500         Semi-Bold           600
'   Bold    700         Extra Bold          800
'   Heavy   900
' ---------------------------------

'=================================================================
'  Constants
'=================================================================

' -- font weight constants --
Public Const wtThin = 100
Public Const wtExtraLight = 200
Public Const wtLight = 300
Public Const wtNormal = 400
Public Const wtMedium = 500
Public Const wtSemiBold = 600
Public Const wtBold = 700
Public Const wtExtraBold = 800
Public Const wtHeavy = 900

'-- text align constants --
Public Const aGeneral = 0
Public Const aLeft = 1
Public Const aCenter = 2
Public Const aRight = 3
Public Const aDistribute = 4

'=================================================================
'  Declarations
'=================================================================
Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As _
     Integer
Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As _
     Integer

' -- Constants --
Private Const LWA_ALPHA     As Long = &H2
Private Const GWL_EXSTYLE   As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000

Public Const CTRL_DEFAULT_BACKCOLOR  As Long = 65535  'RGB(255, 255, 0) highlight yellow

' -- Values --
Public NoData As Scripting.Dictionary

' -- Functions --
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal crKey As Long, _
   ByVal bAlpha As Byte, _
   ByVal dwFlags As Long) As Long

Public RefSub As String 'referring subroutine

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

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CloseFormsReports[mod_Forms])"
    End Select
    Resume Exit_Handler
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

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsOpen[mod_UI])"
    End Select
    Resume Exit_Handler
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
'               BLC, 6/12/2016 - revised to use AppSettings SWITCHBOARD value
' =================================
Public Function SwitchboardIsOpen() As Boolean
    On Error GoTo Err_Handler

    SwitchboardIsOpen = False    ' Default in case of error

    'check for switchboard in all open forms ( AllForms.IsLoaded() )
    If CurrentProject.AllForms(SWITCHBOARD).IsLoaded = True Then
        If CurrentProject.AllForms(SWITCHBOARD).CurrentView = 1 Then
            SwitchboardIsOpen = True
        End If
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SwitchboardIsOpen[mod_UI])"
    End Select
    Resume Exit_Handler
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
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsLoaded[mod_UI])"
    End Select
    Resume Exit_Handler
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

Exit_Handler:
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
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          SetFormOpacity
' Description:  Sets form opacity
' Assumptions:  place in forms module mod_Form for protocols which utilize that module
' Parameters:   frm - form to prepare
'               sngOpacity - opacity of the form (single)
'               TColor - color for the form display (long)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Thenman, September 24, 2009
' http://www.access-programmers.co.uk/forums/showthread.php?t=154907
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 6/1/2016  - moved to mod_Forms from mod_App_UI (uplands)
' ---------------------------------
Public Sub SetFormOpacity(frm As Form, sngOpacity As Single, TColor As Long)
On Error GoTo Err_Handler

    Dim lngStyle As Long
    
    ' get the current window style, then set transparency
    lngStyle = GetWindowLong(frm.hWnd, GWL_EXSTYLE)
    SetWindowLong frm.hWnd, GWL_EXSTYLE, lngStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hWnd, TColor, (sngOpacity * 255), LWA_ALPHA
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetFormOpacity[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' SUB:          CaptureEscapeKey
' Description:  Handles ESCAPE key actions for certain forms
' Assumptions:
' Note:         Handles ESC for the following modal forms:
'               fsub_Soil_Stability, fsub_Fuels_LD, frm_Locations, frm_Unknown_Species
' Parameters:   KeyCode - keycode detected (key down)
' Returns:      -
' Throws:       none
' References:
'  John Spencer, 3/11/2010
'  http://msgroups.net/microsoft.public.access/how-best-to-disable-esc-key-on-form/21881
' Source/date:  Bonnie Campbell, August 21, 2015 - for NCPN tools
' Revisions:    BLC, 8/21/2015 - initial version
'               BLC, 6/1/2016  - added to mod_Forms from mod_App_UI (uplands)
' =================================
Public Sub CaptureEscapeKey(KeyCode As Integer)
On Error GoTo Err_Handler

    If KeyCode = vbKeyEscape Then
        If MsgBox("Undo changes?" & vbCrLf & vbCrLf & _
            "If yes, this may undo all recent changes (not just for a single field)." & vbCrLf & vbCrLf & _
            "Note:" & vbCrLf & _
            "If your cursor was in a..." & vbCrLf & _
            "+ text field, dropdown listbox, or checkbox field >> ALL changes will be undone." & vbCrLf & _
            "+ text field changed immediately before you clicked ESCAPE >> only the text field changes will be undone." & vbCrLf & vbCrLf & _
            "Previously saved data will remain unchanged.", vbYesNo, "ESCAPE Pressed!") = vbNo Then
            KeyCode = 0
        End If
        'KeyCode = 0
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CaptureEscapeKey[mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ToggleForm
' Description:  Minimizes, maximizes, or restores form display
' Assumptions:
' Note:         -
' Parameters:   strForm - form to change (string)
'               Sizing - how to change display (integer) -1 = minimize, 0 = normal/restore, 1 = maximize
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 24, 2016  - for NCPN tools
' Revisions:    BLC, 6/24/2016 - initial version
' ---------------------------------
Public Sub ToggleForm(strForm As String, Sizing As Integer)
On Error GoTo Err_Handler
    
    'ensure form is open, if not -> exit
    If Not FormIsOpen(strForm) Then GoTo Exit_Handler
    
    Forms(strForm).SetFocus
    
    Select Case Sizing
        Case -1 'minimize
            DoCmd.Minimize
        Case 0 'restore
            DoCmd.Restore
        Case 1 'maximize
            DoCmd.Maximize
    End Select
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[AppReleases form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ClearForm
' Description:  Clear form fields
' Assumptions:  Form setup is similar to big rivers contact form w/ data entry
'               above and list below
' Parameters:   frm - form
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 23, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/23/2016 - initial version
'   BLC - 6/27/2016 - shifted to mod_Forms from big rivers forms
'   BLC - 7/28/2016 - added clearing lblMsg caption
'   BLC - 8/30/2016 - added RefSub to identify form subs called by ClearForm
' ---------------------------------
Public Sub ClearForm(ByRef frm As Form)
On Error GoTo Err_Handler
    
    'set global
    RefSub = "ClearForm"
    
    With frm
    
        'clear recordsource
        .RecordSource = ""
        
        'clear values so they no longer look for original control sources
        Dim ctrl As Control
        
        'clear the control sources to clear the textboxes
        For Each ctrl In frm.Controls
            Select Case ctrl.ControlType
                Case acTextBox
                    ctrl.ControlSource = ""
                    ctrl.value = ""
                Case acComboBox
                    'ctrl.Value = "" '<< error: 2448 can't assign value to object
                    'ctrl.Value = Null '<< error: 2448 can't assign value to object
                    'ctrl.ItemData (0)
                    ' Johanness, October 12, 2012
                    ' http://stackoverflow.com/questions/12697427/vba-clear-selections-of-a-combobox
            End Select
        Next
        
        .Controls("tbxIcon") = StringFromCodepoint(uBullet)
        .Controls("tbxIcon").ForeColor = lngRed
        .Controls("tbxID") = 0
        .Controls("lblMsgIcon").Caption = ""
        .Controls("lblMsg").Caption = ""
        .Controls("lblMsgIcon").ForeColor = lngRobinEgg
        
        .Controls("btnSave").Enabled = False
        
        .list.Requery
        
        .Requery
    
    End With
    
Exit_Handler:
    RefSub = ""
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearForm[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          LimitKeyPress
' Description:  Limit form fields to a set number of characters
' Assumptions:  Control passed in is a text or combo box
' Parameters:   ctrl - textbox/combobox (control)
'               iMaxLen - # of allowed characters (integer)
'               KeyAscii - character passed in (integer)
' Returns:      -
' Throws:       none
' Usage:        Call LimitKeyPress(Me.MyTextBox, 12, KeyAscii) in control's KeyPress event
' References:   LimitChange() required in control's Change event also
' Source/date:
'   Allen Browne, unknown
'   http://allenbrowne.com/ser-34.html
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Public Sub LimitKeyPress(ctrl As Control, iMaxLen As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler
    
    With ctrl
        If Len(.Text) - .SelLength >= iMaxLen Then
            If KeyAscii <> vbKeyBack Then
                KeyAscii = 0
                Beep
            End If
        End If
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LimitKeyPress[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          LimitChange
' Description:  Limit form fields to a set number of characters
' Assumptions:  Control passed in is a textbox
' Parameters:   ctrl - textbox cotnrol
' Returns:      -
' Throws:       none
' Usage:        Call LimitChange(Me.MyTextBox, 12) in control's Change event
' References:   LimitKeyPress() required in controls KeyPress event also
' Source/date:
'   Allen Browne, unknown
'   http://allenbrowne.com/ser-34.html
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Public Sub LimitChange(ctrl As Control, iMaxLen As Integer)
On Error GoTo Err_Handler

    Dim msg As String
    
    With ctrl
        If Len(.Text) > iMaxLen Then
            msg = "Oops! " & .Name & " field too long. Truncated to " & iMaxLen & " characters."
        
            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                "msg" & PARAM_SEPARATOR & msg & _
                "|Type" & PARAM_SEPARATOR & "caution"
            
            .Text = Left(.Text, iMaxLen)
            .SelStart = iMaxLen
        End If
    End With
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LimitChange[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub