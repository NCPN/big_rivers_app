Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_UI
' Level:        Framework module
' Version:      1.08
' Description:  User interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/10/2015 - 1.01 - added GetRibbonXML()
'               BLC, 5/27/2015 - 1.02 - added functions
'               BLC, 6/30/2015 - 1.03 - moved to mod_Forms: FormIsOpen, FormIsLoaded, SwitchboardIsOpen
'                                       moved from mod_Forms: ChangeBackColor
'               BLC, 5/13/2016 - 1.04 - adapted for Big Rivers
'               BLC, 6/27/2016 - 1.05 - added acNormalSolid, acTransparent constants
'               BLC, 6/24/2016 - 1.06 - replaced Exit_Function > Exit_Handler
'               BLC, 7/6/2016  - 1.07 - added functions to hide VBE (shift off screen)
'                                       while the enum module is being updated
'               BLC, 9/1/2016  - 1.08 - updated ControlExists()
' =================================

' ---------------------------------
' Declarations
' ---------------------------------
Public Const acNormalSolid As Integer = 1
Public Const acTransparent As Integer = 0

' ---------------------------------
'  Properties
' ---------------------------------


' ---------------------------------
'  VBE
' ---------------------------------

' ---------------------------------
'   NOTES:
'
'
'       Application.Echo False
'       SetWindowPos FindWindow("wndclass_desked_gsk", _
'           Application.VBE.MainWindow.Caption), 0&, 0&, 2000&, 1, 1, &H80 Or &H1
'       <run your code that calls the VBE>
'       Application.VBE.MainWindow.visible = False
'       Application.Echo True
'
' ---------------------------------

Declare Function SetWindowPos Lib "user32.dll" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal ClassName As String, _
    ByVal WindowName As String) As Long

Const HWND_NOTOPMOST = -2
Const SWP_HIDEWINDOW = &H80
Const SWP_NOSIZE = &H1

' ---------------------------------
' Function:     OpenAndHideVBE
' Description:  Opens then hides VBE
' Notes:        Call OpenAndHideVBE before writing to the project
'               and ShowAndCloseVBE when done.
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Peter Thornton, March 23, 2013
'   https://social.msdn.microsoft.com/Forums/en-US/197a9f1d-96cb-49d6-b08c-0dcae1eafc08/vbe-flashes-while-programming-in-the-vbe?forum=isvvba
'   AOB, September 5, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=252942
' Source/date:  Bonnie Campbell, July 6, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/6/2016 - initial version
' ---------------------------------
Public Sub OpenAndHideVBE()
On Error GoTo Err_Handler
    
    Dim hWndVBE As Long
    Dim objVBE As VBE
    
    Set objVBE = Application.VBE
    
    hWndVBE = FindWindow("wndclass_desked_gsk", _
                            Application.VBE.MainWindow.Caption)
    
    Call SetWindowPos(hWndVBE, 0&, 0&, 2000&, 1, 1, _
                        SWP_HIDEWINDOW Or SWP_NOSIZE)
    
    Application.VBE.MainWindow.Visible = True
    'Application.Caption errors for Access w/ Method or data member not found
    'use "already open form caption", false instead
    'AppActivate Application.Caption
    AppActivate SWITCHBOARD, False
    DoCmd.OpenForm SWITCHBOARD, acNormal, , , , acDialog

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenAndHideVBE[mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     ShowAndCloseVBE
' Description:  Displays VBE and closes it
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Peter Thornton, March 23, 2013
'   https://social.msdn.microsoft.com/Forums/en-US/197a9f1d-96cb-49d6-b08c-0dcae1eafc08/vbe-flashes-while-programming-in-the-vbe?forum=isvvba
'   AOB, September 5, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=252942
' Source/date:  Bonnie Campbell, July 6, 2016 for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/6/2016 - initial version
' ---------------------------------
Public Sub ShowAndCloseVBE()
On Error GoTo Err_Handler
    
    Dim hWndVBE As Long
    Dim cbt As CommandBarButton
    Dim objVBE As VBIDE.VBE
    Dim objWin As VBIDE.Window

    Set objVBE = Application.VBE
    ' optionally close all module windows,
    ' or just the newly opened module Window
    
    For Each objWin In objVBE.Windows
        If objWin.Type = vbext_wt_CodeWindow Then
                objWin.Close
        ElseIf objWin.Type = vbext_wt_Designer Then
                objWin.Close
        End If
    Next
    
    objVBE.MainWindow.WindowState = vbext_ws_Minimize
    objVBE.MainWindow.Visible = False
    
    hWndVBE = FindWindow("wndclass_desked_gsk", _
                            Application.VBE.MainWindow.Caption)
    
    Call SetWindowPos(hWndVBE, HWND_NOTOPMOST, 0, 0, 400, 300, 0)
    
    Set cbt = Application.VBE.CommandBars.FindControl(ID:=752)
    
    'Application.Caption errors for Access w/ Method or data member not found
    'use "already open form caption", false instead
    'AppActivate Application.Caption
    AppActivate SWITCHBOARD, False
    DoCmd.OpenForm SWITCHBOARD, acNormal, , , , acDialog
    
    cbt.Execute

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ShowAndCloseVBE[mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Ribbon
' ---------------------------------
' =================================
' FUNCTION:     GetRibbonXML
' Description:  gets ribbon UI XML specified, if found
' Assumes:      USysRibbon table exists
' Parameters:   ribbon - name of the ribbon to retrieve, RibbonName in USysRibbon (string)
' Returns:      XML of the specified ribbon
' Throws:       none
' References:   none
' Source/date:  -
' Revisions:    BLC, 5/10/2015 - initial version
' =================================
Public Function GetRibbonXML(strRibbon As String) As String
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Dim strSQL As String, strXML As String
    
    strSQL = "SELECT RibbonXML FROM USysRibbons WHERE RibbonName = '" & strRibbon & "';"
    strXML = ""
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If Not (rs.BOF And rs.EOF) Then
        strXML = rs!RibbonXML
    End If
    
    GetRibbonXML = strXML

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRibbonXML[mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' SUB:          RibbonOnLoad
' Description:  Callback function for ribbon customization
' Parameters:   ribbon - office ribbon control (IRibbonUI object)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from http://www.experts-exchange.com/Database/MS_Access/Q_28470268.html
'               by Christian, 7/7/2014.
' Revisions:    BLC, 5/17/2015 - initial version
' =================================
'Public objRibbon As IRibbonUI
Public Sub RibbonOnLoad(ribbon As Office.IRibbonUI)
On Error GoTo Err_Handler
Dim prv_Ribbon As IRibbonUI

    Set prv_Ribbon = ribbon

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RibbonOnLoad[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' =================================
' SUB:          GetRibbonVisibility
' Description:  Callback function to indicate if ribbon control should be displayed or not
' Parameters:   ctrl - office ribbon control (IRibbonControl object)
'               visible - true (boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from http://www.access-programmers.co.uk/forums/showthread.php?t=246015
'               by Mark K., 4/26/2013.
' Revisions:    BLC, 5/10/2015 - initial version
' =================================
Public Sub GetRibbonVisibility(ctrl As Office.IRibbonControl, ByRef Visible)
On Error GoTo Err_Handler

    Select Case ctrl.ID
        Case "tabExportOptions"
            Visible = True
            TempVars.Add "ribbon", True
        Case Else
            Visible = False
            TempVars.Add "ribbon", False
    End Select
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRibbonVisibility[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Forms
' ---------------------------------

' =================================
' SUB:          SetWindowSize
' Description:  sets form size (width & height)
' Assumptions:  -
' Note:         dimensions are in twips (1 inch = 1440 twips)
' Parameters:   ctrl - office ribbon control (IRibbonControl object)
'               visible - true (boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Hasup, February 26,2014
'   http://stackoverflow.com/questions/22021802/resize-form-in-ms-access-by-changing-detail-height
' Adapted:      Bonnie Campbell, May 27, 2015 - for NCPN tools
' Revisions:    BLC, 5/27/2015 - initial version
' =================================
Public Sub SetWindowSize(ByRef frm As Form, ByRef lngHeight As Long, ByRef lngWidth As Long)
On Error GoTo Err_Handler

'    If Me.WindowHeight = 4044 Then
'        lngHeight = 8000
'    Else
'        lngHeight = 4044
'    End If
    frm.Move frm.WindowLeft, Height:=lngHeight, Width:=lngWidth
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetWindowHeight[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' =================================
' SUB:          PopulateSubformControl
' Description:  Set the form for a subform control
' Parameters:   ctrl - subform control to populate
'               strSubFormName - name of the subform to use in the control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, 5/1/2015 for NCPN tools
' Revisions:    BLC, 5/1/2015 - initial version
' =================================
Public Sub PopulateSubformControl(ctrl As SubForm, strSubFormName As String)
    On Error GoTo Err_Handler

    ctrl.SourceObject = strSubFormName 'Forms(strSubFormName)

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateSubformControl[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          RepaintParentForm
' Description:  Repaints the control's parent(or grandparent or great grandparent...) form
' Parameters:   ctl - control whose parent form you're looking to repaint
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell August, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 8/20/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub RepaintParentForm(ctl As Control)
On Error GoTo Err_Handler:
Dim parentControl As Object
        
    Set parentControl = ctl.Parent
    
    Do Until parentControl Is Nothing
      
        If TypeName(parentControl.Name) = "String" Then
            'form? -> refresh the display
            If getAccessObjectType(parentControl.Name) = -32768 Then
                parentControl.Repaint
                Exit Do
            End If
            Set parentControl = parentControl.Parent
        Else
            'form? -> refresh the display
            If CurrentProject.AllForms(parentControl.Name).IsLoaded Then
                parentControl.Repaint
                Exit Do
            End If
        End If
    Loop
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RepaintParentForm[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' FUNCTION:     ChangeBackColor
' Description:  change background color of control
' Assumptions:  -
' Parameters:   ctrl- control to change color
'               lngColor = color (long)
' Returns:      N/A
' Throws:       none
' References:   none
' Note:         MUST be a function vs. sub to be called w/in form event ( =ChangeBackColor(Me,lngYelLime) )
' Source/date:  Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015  - initial version
' ---------------------------------
Public Function ChangeBackColor(ctrl As Control, lngColor As Long)
On Error GoTo Err_Handler

    ctrl.BackColor = lngColor
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeBackColor[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ResetHeaders
' Description:  reset header fields to their
' Assumptions:  if only a subset of form controls are to be reset, these controls should have the same Tag property value
' Parameters:   frm - form to reset headers on
'               allCtrls - if all form controls should be reset (boolean) (true = reset all controls,
'                           false = reset one control [requires oCtrl to be populated])
'               ctrlTag - control's tag string if resetting only a subset of forms controls (string)
'               fontBold - whether text should be bold (boolean) (true = make font bold, false not bold),  (optional)
'               backstyle - if back control back color is normal or transparent (integer) (1-normal 0-transparent) (optional)
'               forecolor - text color (long) (optional)
'               backcolor - backgound color of control (long) (optional)
'               oCtrl - control to change, if only one control is to be changed (optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Fionnuala January 20, 2013
' http://stackoverflow.com/questions/3344649/how-to-loop-through-all-controls-in-a-form-including-controls-in-a-subform-ac
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
' ---------------------------------
Public Sub ResetHeaders(frm As Form, _
                        allCtrls As Boolean, _
                        ctrlTag As String, _
                        Optional fontBold As Boolean = True, _
                        Optional backstyle As Integer = 1, _
                        Optional ForeColor As Long, _
                        Optional BackColor As Long, _
                        Optional oCtrl As Control)
On Error GoTo Err_Handler

Dim ctrl As Control

    If allCtrls = True Then
    
        'iterate through all form controls
        For Each ctrl In frm
            
            'check control type
             If ctrl.ControlType = acTextBox Or _
                ctrl.ControlType = acComboBox Or _
                ctrl.ControlType = acListBox Or _
                ctrl.ControlType = acLabel _
             Then
             
                'check tag
                If ctrl.Tag = ctrlTag Then
                    If varType(fontBold) = vbBoolean Then ctrl.fontBold = fontBold
                    If varType(backstyle) = vbInteger Then ctrl.backstyle = backstyle
                    If varType(BackColor) = vbLong Then ctrl.BackColor = BackColor
                    If varType(ForeColor) = vbLong Then ctrl.ForeColor = ForeColor
                End If
                
          End If
          
        Next
    Else
        'reset only oCtrl

        'check tag
        If oCtrl.Tag = ctrlTag Then
        
            'check control type
            If oCtrl.ControlType = acTextBox Or _
                oCtrl.ControlType = acComboBox Or _
                oCtrl.ControlType = acListBox Or _
                oCtrl.ControlType = acLabel _
            Then
          
                If varType(fontBold) = vbBoolean Then oCtrl.fontBold = fontBold
                If varType(backstyle) = vbInteger Then oCtrl.backstyle = backstyle
                If varType(BackColor) = vbLong Then oCtrl.BackColor = BackColor
                If varType(ForeColor) = vbLong Then oCtrl.ForeColor = ForeColor
             
            End If
            
        End If

    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ResetHeaders[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ShowControls
' Description:  toggle control visibility
' Assumptions:  if only a subset of form controls are to be reset, these controls should have the same Tag property value
' Parameters:   frm - form to reset headers on
'               allCtrls - if all form controls should be reset (boolean) (true = reset all controls,
'                           false = reset one control [requires oCtrl to be populated])
'               ctrlTag - control's tag string if resetting only a subset of forms controls (string)
'               visibility - whether control should be visible or not (boolean) (true = make font bold, false not bold),  (optional)
'               oCtrl - control to change, if only one control is to be changed (optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Fionnuala January 20, 2013
' http://stackoverflow.com/questions/3344649/how-to-loop-through-all-controls-in-a-form-including-controls-in-a-subform-ac
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015 - initial version
'   BLC - 6/30/2015 - update documentation
' ---------------------------------
Public Sub ShowControls(frm As Form, _
                        allCtrls As Boolean, _
                        ctrlTag As String, _
                        visibility As Boolean, _
                        Optional oCtrl As Control)
On Error GoTo Err_Handler

Dim ctrl As Control

    If allCtrls = True Then
    
        'iterate through all form controls
        For Each ctrl In frm

            'check tag
            If ctrl.Tag = ctrlTag Then
                ctrl.Visible = visibility
            End If

        Next
    Else
        'reset only oCtrl

        'check tag
        If oCtrl.Tag = ctrlTag Then
                oCtrl.Visible = visibility
        End If

    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ShowControls[form_frmSpeciesSearch])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
'  Reports
' ---------------------------------

' =================================
' FUNCTION:     ReportIsLoaded
' Description:  Returns whether the specified report is loaded
' Parameters:   strReportName - string for the name of the report to check
' Returns:      True if the specified report is open, False if not
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell - 5/17/2015 - for NCPN tools
' Revisions:    BLC, 5/17/2015 - initial version
' =================================
Public Function ReportIsLoaded(ByVal strReportName As String) As Boolean
On Error GoTo Err_Handler
 
    ' Possible states returned by SysCmd & CurrentView
    Const cObjStateClosed = 0
    Const cDesignView = 0
    Const cPrintView = 5
    Const cReportView = 6
    Const cLayoutView = 7

    ' check current state - not open or nonexistent, design, print, layout, or report view
    If SysCmd(acSysCmdGetObjectState, acReport, strReportName) <> cObjStateClosed Then
        ' check current view, return True if open and not in design view
        If Reports(strReportName).CurrentView <> cDesignView Then
            ReportIsLoaded = True
        End If
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReportIsLoaded[mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Tabs
' ---------------------------------

' =================================
' SUB:          tabPageUnhide
' Description:  sets desired tab visible, all others hidden
' Parameters:   strTabName - tab page name to make visible
'               ctrl - tab control
'               blnHideOnly - true to hide tabs only (Boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Adapted from Tom's post comment, 9/12/2009
'               http://www.vbdotnetforums.com/gui/36561-loop-through-tab-pages-remove.html
'               Created 06/11/2014 blc; Last modified 06/11/2014 blc.
' Adapted:      Bonnie Campbell, June 11, 2014 - initial version
' Revisions:    BLC, June 11, 2014 - initial version
'               BLC, June 9, 2015  - adjust for hiding tabs only with blnHideOnly
' =================================
Public Sub tabPageUnhide(ctrl As TabControl, strTabName As String, Optional blnHideOnly As Boolean)
On Error GoTo Err_Handler

    Dim pg As Page
    
    For Each pg In ctrl.Pages
        If pg.Name = strTabName Then
            If Not blnHideOnly = True Then
                ctrl.Pages(pg.Name).Visible = True
            End If
        Else
            ctrl.Pages(pg.Name).Visible = False
        End If
    Next pg
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tabPageUnhide[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Controls
' ---------------------------------

' =================================
' FUNCTION:     HideObject
' Description:  Changes the hidden property of an object to hide / show in the database window
' Parameters:   strObjectName - name of the object (string)
'               blnHide - True to hide, False to show (default True)
'               varType - object type (default acTable)
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/25/2009
' Revisions:    JRB, 6/25/2009 - initial version
'               BLC, 4/30/2015 - move from mod_Utilities to mod_UI
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function HideObject(strObjectName As String, _
    Optional blnHide As Boolean = True, Optional varType As Variant = acTable)

    On Error GoTo Err_Handler

    SetHiddenAttribute varType, strObjectName, blnHide

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HideObject[mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     ControlExists
' Description:  determines if a control exists in a form
' Parameters:   ctlName - control to check for (string)
'               frm - form to check on (form)
' Returns:      boolean - true if control exists, false if not
' Throws:       none
' References:   none
' Source/date:
'   VBslammer, March 22, 2005
'   http://www.tek-tips.com/viewthread.cfm?qid=1029435
'   Mike Lyons September 21, 2007
'   http://www.utteraccess.com/forum/Control-Exist-Form-t1505884.html
' Adapted:      Bonnie Campbell, May 15, 2015 - for NCPN tools
' Revisions:    BLC, 5/12/2015 - initial version
'               BLC, 9/1/2016  - added false path, updated documentation
' =================================
Function ControlExists(ByRef ctlName As String, ByRef frm As Form) As Boolean
On Error GoTo Err_Handler
  Dim ctl As Control
  
  For Each ctl In frm.Controls
    If ctl.Name = ctlName Then
      ControlExists = True
      GoTo Exit_Handler
    End If
  Next ctl
  
  'doesn't exist
  ControlExists = False
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ControlExists[mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          buttonHighlight
' Description:  Toggle button color to strColor or transparent if already colored
' Parameters:   btn      - name of the button to change
'                          accommodates command and label as control buttons
'               strColor - color as a string (hex)
'               solo - display only this control & leave others transparent (Boolean)
'               toggle - change the display for a control (Boolean)
'               intEffect - control display effect (integer)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub buttonHighlight(btn As Control, Optional solo As Boolean, Optional Toggle As Boolean, Optional intEffect As Integer, Optional strColor As String)
' Special Effects:  0 - flat, 1 - raised, 2 - sunken, 3 - etched, 4 - shadowed, 5 - chiseled
' Colors:
'   lime                   #9EFF00
'   chartreuse 1           #7FFF00 127 255 00  65407
'   dark olive green 1     #CAFF70 202 255 112 7405514
'   mint                   #BDFCC9 189 252 201 13237437
'   light lime (like)      #E6FABF 230 250 191
'   darker lt lime         #CFF583 207 245 131
On Error GoTo Err_Handler:

    'toggle button
    If Toggle Then
        buttonUnHighlight btn, Toggle
    End If
    
    'change all others to transparent if solo
    If solo Then
        buttonUnHighlight btn
    End If
    
    With btn
        If .backstyle = 1 Then
            GoTo Transparent
        End If
        
        If (Len(strColor) <> 6) Then
            strColor = "CFF583"
        End If
    
        If intEffect > -1 Or intEffect > 6 Then
            intEffect = 0 'flat
        End If
           
        'change button background to given color
        .backstyle = 1 'Normal - required to change color
        .BackColor = HTMLConvert("#" & strColor)
        .SpecialEffect = intEffect
    End With
    
Exit_Procedure:
    Exit Sub

Transparent:
    btn.backstyle = 0 'Transparent
    GoTo Exit_Procedure

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - buttonHighlight[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          buttonUnHighlight
' Description:  Toggles all other buttons to transparent if already colored
' Parameters:   btn - name of the button control to change
'                     accommodates command and label as control buttons
'               blnToggle - toggle only the identified button (Boolean)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub buttonUnHighlight(btn As Control, Optional blnToggle As Boolean)
On Error GoTo Err_Handler:
Dim ctl As Control

    With btn
        'unhighlight only btn
        If blnToggle Then
            .backstyle = 0 'transparent
            .SpecialEffect = 0 'flat
            GoTo Exit_Procedure
        End If
        
        'unhighlight all other buttons
        For Each ctl In .Parent.Controls

            If ctl.Name <> btn.Name And _
                ctl.ControlType = acLabel Then
                With ctl
                    .backstyle = 0 'transparent
                End With
            End If

        Next
    
    End With
    
Exit_Procedure:
    'update display
    RepaintParentForm btn
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - buttonUnHighlight[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          DisableControl
' Description:  Set color scheme for labels so they appear disabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - moved from mod_List to mod_UI
' ---------------------------------
Public Sub DisableControl(ctrl As Control)

On Error GoTo Err_Handler
    
    ctrl.BackColor = lngLtGray
    ctrl.ForeColor = lngGray
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = lngGray
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControl[mod_UI])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          EnableControl
' Description:  Set color scheme for labels so they appear enabled
' Assumptions:  Assumes control has BackColor and ForeColor properties
' Parameters:   ctrl - control to set color scheme for
'               backColor - long value for desired back color
'               foreColor - long value for desired fore (text) color
'               optionally for command buttons:
'               borderColor - long value for desired border color
'               hoverColor - long value for desired hover color
'               pressedColor - long value for desired pressed button color
'               hoverForeColor - long value for desired hover fore (text) color
'               pressedForeColor - long value for desired pressed button fore (text) color
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/22/2015 - moved from mod_List to mod_UI
' ---------------------------------
Public Function EnableControl(ctrl As Control, BackColor As Long, ForeColor As Long, _
                                Optional borderColor As Long, _
                                Optional HoverColor As Long, _
                                Optional pressColor As Long, _
                                Optional hoverForeColor As Long, _
                                Optional pressedForeColor As Long)
On Error GoTo Err_Handler
    
    ctrl.BackColor = BackColor
    ctrl.ForeColor = ForeColor
    
    If ctrl.ControlType = acCommandButton Then
        ctrl.borderColor = borderColor
        ctrl.HoverColor = HoverColor
        ctrl.PressedColor = pressColor
        ctrl.hoverForeColor = hoverForeColor
        ctrl.pressedForeColor = pressedForeColor
    End If

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableControl[mod_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ToggleControl
' Description:  Toggles control font (fore) color & enables/disables
' Parameters:   frmName - name of parent form (string)
'               btnName - name of the button control to change
'                     accommodates command and label as control buttons (string)
'               color - optional color value (long)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell May 12, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 5/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub ToggleControl(frmName As String, btnName As String, Optional color As Variant = Null)
On Error GoTo Err_Handler:
    
    Dim ctrl As Control
    Set ctrl = Forms(frmName).Controls(btnName)
    
    'invert enabled value (change true -> false, false -> true) & change color
    With ctrl
    
        'enable/disable control (includes acCommandButton, acComboBox, acListBox, acTextBox, acToggleButton)
        If Not ctrl.ControlType = acLabel Then
            .Enabled = Not .Enabled
        End If
        
        If Not IsNull(color) Then
            ' change font color for appropriate controls with text
            Select Case ctrl.ControlType
                Case acCommandButton, acComboBox, acLabel, acListBox, acTextBox, acToggleButton
                    .ForeColor = color
                Case Else
            End Select
        End If
    End With
    
Exit_Procedure:
    'update display
    RepaintParentForm Forms(frmName).Controls(btnName)
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleControl[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Text
' ---------------------------------

' =================================
' FUNCTION:     CrumbsToArray
' Description:  Prepares breadcrumb elements from Me.OpenArgs values
' Parameters:   strCrumbs - Me.OpenArgs values from form open subs
'               delimiter - delimiter used for separating string values, default = | (pipe)
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    BLC, 6/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function CrumbsToArray(strCrumbs As String, Optional delimiter = "|")

On Error GoTo Err_Handler

    Dim strCrumbTrail As String

    If Len(strCrumbs) > 0 Then
        Dim aryCrumbs As Variant
        
        aryCrumbs = Split(strCrumbs, delimiter)
        
    End If

    CrumbsToArray = aryCrumbs
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CrumbsToArray[mod_UI])"
    End Select
    Resume Exit_Procedure
End Function

' =================================
' SUB:     PrepareCrumbs
' Description:  Sets breadcrumb label control captions & click events based on crumb element array
' Assumptions:  Breadcrumbs are displayed using label controls (lblCrumb01...)
'               & labels already exist on the targeted form
' Parameters:   frm - form holding crumb labels
'               aryCrumbs - breadcrumb array
'               separator - non-clickable value between crumbs, default = >
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    BLC, 6/12/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' =================================
Public Sub PrepareCrumbs(frm As SubForm, aryCrumbs As Variant, Optional separator = ">")
 On Error GoTo Err_Handler
 
    Dim ctrl As Control
    Dim i As Integer, intLastCtrlWidth As Integer, intLastCtrlPosition As Integer
    Dim strNum As String, strCtrlName As String, strCtrlSeparator As String
    
    'initialize
    intLastCtrlPosition = 10
    
    'avoid flicker
    'Painting = False
    
    For i = 1 To UBound(aryCrumbs)
        ' set lbl caption
        If (i < 10) Then
            strNum = 0 & i
        Else
            strNum = i
        End If
        
        strCtrlName = "lblCrumb" & strNum
        
        With frm.Controls(strCtrlName)
       
            If .ControlType = acLabel Then
                'label control
                .Caption = aryCrumbs(i)
            Else
                'hyperlink control (displaytext vs caption)
                .Value = aryCrumbs(i)
            End If
            
            'set control position
            If intLastCtrlPosition > frm.Controls(strCtrlName).Parent.Width Then
                .Left = frm.Controls(strCtrlName).Parent.Width - .Width
            Else
                .Left = intLastCtrlPosition
            End If
            
            'set control width
'            setControlWidth frm.Controls(strCtrlName), , frm.Controls(strCtrlName).Parent.Width
            
            'save new ctrl width for setting separator position
            intLastCtrlWidth = .Width
        
        End With
        
        'display the separator
        If (i < UBound(aryCrumbs)) Then
          strCtrlSeparator = "lblSep" & strNum
          With frm.Controls(strCtrlSeparator)
            .Left = intLastCtrlPosition + intLastCtrlWidth + 10
            .Caption = separator
            .Visible = True
            
            'determine position of next control
            intLastCtrlPosition = .Left + .Width + 10
          End With
        End If
        
    Next i
    
    'ready for viewing
    'Painting = True
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareCrumbs[mod_UI])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Drawing
' ---------------------------------

' ---------------------------------
' SUB:          CircleControl
' Description:  Draws a circle around the control
' Assumptions:  -
' Parameters:   ctrl - control to circle (control)
'               ellipse - whether it should be an ellipse vs. circle (boolean)
' Returns:      -
' Throws:       none
' References:
'   Duane Hookom, October 6, 2008
'   http://www.pcreview.co.uk/threads/circle-a-word-in-access-report.3639434/
'
'   https://msdn.microsoft.com/en-us/library/office/aa195881(v=office.11).aspx
' Source/date:  Bonnie Campbell, May 10, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/10/2016 - initial version
' ---------------------------------
Public Sub CircleControl(ctrl As Control, Optional ellipse As Boolean = False)
On Error GoTo Err_Handler
    
    Dim iWidth As Integer, iHeight As Integer
    Dim iCenterX As Integer, iCenterY As Integer
    Dim iRadius As Integer
    Dim dblAspect As Double
    Dim sngStart As Single, sngEnd As Single

    iCenterX = ctrl.Left + ctrl.Width / 2
    iCenterY = ctrl.Top + ctrl.Height / 2
    iRadius = ctrl.Width '/ 3 '/ 2 + 100
    dblAspect = 1 'ctrl.Height / ctrl.Width
    
    sngStart = -0.00000001                    ' Start of pie slice.

    sngEnd = -2 * PI / 3                         ' End of pie slice.
    ctrl.Parent.fillColor = RGB(51, 51, 51)            ' Color pie slice red.
    ctrl.Parent.FillStyle = 0                          ' Fill pie slice.
    
    'add the circle to the parent
    ' X,Y center | radius | [ color, start, end, aspect ]
    ctrl.Parent.Circle (iCenterX, iCenterY), iRadius, lngLime, sngStart, sngEnd, dblAspect

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CircleControl[mod_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddControl
' Description:  Adds a control to a form
' Assumptions:  -
' Parameters:   frm - form to add controls to (form)
'               ctlType - type of control to add (control)
'               ctlName - name of control to add (string)
'               ctlData - data for control (optional, variant)
' Returns:      -
' Throws:       none
' References:
'   Chip Pearson, unknown
'   http://www.ozgrid.com/Excel/free-training/ExcelVBA2/excelvba2lesson21.htm
' Source/date:  Bonnie Campbell, October 11, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/11/2016 - initial version
' ---------------------------------
Public Sub AddFormControl(frmName As String, ctlType As Long, ctlName As String, Optional ctlData As Variant, _
                        Optional w As Integer, Optional h As Integer, _
                        Optional xPos As Integer, Optional yPos As Integer)
On Error GoTo Err_Handler
    
    'Dim progID As String
    Dim c As Control
    
    'progID = "Forms." & ctlType & ".1"
    
'    Set c = frm.Controls.Add(progID, ctlName)
    Set c = CreateControl(frmName, ctlType, acDetail)

    c.Name = ctlName
    
    If Not ctlData Is Nothing Then
        Set c.Recordset = ctlData
    End If
    
    'set dimensions & location
    If IsNumeric(w) Then c.Width = w
    If IsNumeric(h) Then c.Height = h
    If IsNumeric(xPos) Then c.Left = xPos
    If IsNumeric(yPos) Then c.Top = yPos

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddControl[mod_UI])"
    End Select
    Resume Exit_Handler
End Sub