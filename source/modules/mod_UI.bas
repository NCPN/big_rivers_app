Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_UI
' Level:        Framework module
' Version:      1.01
' Description:  Generic UI functions & procedures
'
' Source/date:  Bonnie Campbell, 5/13/2016
' Revisions:    BLC, 5/13/2016 - 1.00 - initial version
'               BLC, 6/27/2016 - 1.01 - added acNormal, acTransparent constants
' =================================

' ---------------------------------
' Declarations
' ---------------------------------
Public Const acNormal As Integer = 1
Public Const acTransparent As Integer = 0

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Subroutines & Functions
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
    iCenterY = ctrl.top + ctrl.Height / 2
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
' SUB:          ConvertLongToRGB
' Description:  Convert long color value to RGB component values
' Assumptions:  User will call specific color values via dict("R"), dict("G"), dict("B") as needed
' Parameters:   lngValue - long color value
' Returns:      -
' Throws:       none
' References:   none
' Requires:     -
' Source/date:
' Adapted:      Bonnie Campbell, May 13, 2016 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015  - initial version
' ---------------------------------
Public Function HTMLConvert(strHTML As String) As Long
On Error GoTo Err_Handler
    


Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HTMLConvert[mod_UI])"
    End Select
    Resume Exit_Function
End Function