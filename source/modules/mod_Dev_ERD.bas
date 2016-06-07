Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_ERD
' Level:        Development module
' Version:      1.00
'
' Description:  ERD related functions & procedures
'
' Source/date:  Bonnie Campbell, April 27, 2016
' Revisions:    BLC - 2/27/2016 - 1.00 - initial version
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
Private Type RECT
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias _
"FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function apiGetWindow Lib "user32" _
Alias "GetWindow" (ByVal hwnd As Long, _
ByVal wCmd As Long) As Long

Private Declare Function apiGetClassName Lib "user32" _
Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long

' ---------------------------------
'  Constants
' ---------------------------------
Private Const SWP_NOSIZE = &H1
Private Const WM_CLOSE = &H10
' GetWindow() Constants
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5

' ---------------------------------
' Sub:          FixERD
' Description:  ERD fixing actions for positioning tables so they
'               are visible in the diagram (fixes negative positions)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Stephen Lebans, August 28, 2006
'   https://bytes.com/topic/access/answers/528324-releationship-diagram-goes-haywire
' Source/date:  Bonnie Campbell, April 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub FixERD()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FixERD[mod_Dev_ERD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cmdCloseWindow_Click
' Description:  Close the Debug Window via code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Stephen Lebans, August 28, 2006
'   https://bytes.com/topic/access/answers/528324-releationship-diagram-goes-haywire
' Source/date:  Bonnie Campbell, April 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cmdCloseWindow_Click()
On Error GoTo Err_Handler
    Dim lngRet As Long
    Dim hWndMDI As Long
    Dim hWndDebug As Long
    
    ' If this instance of Access has set the
    ' Debug Window to "Always on Top" via the menu:
    ' Tools->Options->Module
    ' then the Debug WIndow is a top level window.
    hWndDebug = FindWindow(vbNullString, "Debug Window")
    If hWndDebug < 0 Then
        lngRet = PostMessage(hWndDebug, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If

    ' The Debug Window is a child of the MDI window
    ' find MDIClient first
    hWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)

    ' Find the Debug Window
    hWndDebug = FindWindowEx(hWndMDI, 0&, "OImmediate", "Debug Window")
    If hWndDebug < 0 Then
        lngRet = PostMessage(hWndDebug, WM_CLOSE, 0&, 0&)
    Else
        MsgBox "The Debug Window is not open."
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdCloseWindow_Click[mod_Dev_ERD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cmdFix_Click
' Description:  Fix any windows that are off the Left edge of the Relationships window
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Stephen Lebans, August 28, 2006
'   https://bytes.com/topic/access/answers/528324-releationship-diagram-goes-haywire
' Source/date:  Bonnie Campbell, April 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cmdFix_Click()
On Error GoTo Err_Handler

    Dim lngRet As Long
    Dim hWndMDI As Long
    Dim hWndRel As Long
    Dim hWndODsk As Long
    Dim hWndTemp
    Dim rc As RECT
    
    ' Force the Relationships window to open
    DoCmd.RunCommand acCmdRelationships
    
    ' Window must be maximized
    DoCmd.Maximize
    DoEvents
    
    ' Relationships Window is a child of the MDI Client window
    ' find MDIClient first.
    hWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)
    
    ' Find the Relationships Window
    hWndRel = FindWindowEx(hWndMDI, 0&, "OSysRel", "Relationships")
    
    If hWndRel = 0 Then
        MsgBox "The Relationships Window is not open.", vbCritical, "Critical Error"
        Exit Sub
    End If
    
    ' The first child window is of class ODsk
    hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)
    
    ' Loop through all of this level's Windows.
    ' We are looking for any windows with a negative
    ' Left value in it's Window Rectangle
    ' Let's get first Child Window of the ODsk window
    hWndTemp = apiGetWindow(hWndODsk, GW_CHILD)
    If hWndTemp = 0 Then
        MsgBox "Their are no Relationships!", vbCritical, "Critical Error"
        Exit Sub
    Else
        lngRet = GetWindowRect(hWndTemp, rc)
        If rc.Left < 1 Then
            rc.Left = 1
            lngRet = SetWindowPos(hWndTemp, 0&, rc.Left, rc.top, 0&, 0&, SWP_NOSIZE)
        End If
    End If
    
    ' Let's walk through every sibling window
    Do
    
        ' Let's get the NEXT SIBLING Window
        hWndTemp = apiGetWindow(hWndTemp, GW_HWNDNEXT)
        
        If hWndTemp < 0 Then
            lngRet = GetWindowRect(hWndTemp, rc)
            If rc.Left < 1 Then
                rc.Left = 1
                lngRet = SetWindowPos(hWndTemp, 0&, rc.Left, rc.top, 0&, _
                                        0&, SWP_NOSIZE)
            End If
        End If
        
        ' Let's Start the process from the Top again.
        ' End this loop if no more Windows.
    Loop While hWndTemp < 0
    ' All done

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdFix_Click[mod_Dev_ERD])"
    End Select
    Resume Exit_Handler
End Sub