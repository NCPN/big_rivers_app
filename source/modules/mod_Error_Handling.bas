Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Error_Handling
' Level:        Framework module
' Version:      1.00
' Description:  Framework-wide related mathematical values, functions & subroutines
'
' Source/date:  Bonnie Campbell, February 23, 2017 for NCPN tools
' Revisions:    BLC, 2/23/2017 - initial version
' =================================

' ---------------------------------
' CONSTANTS:    global constant values
' Description:  values setting application level contants
' References:
'   FMS Inc., Unknown
'   http://www.fmsinc.com/tpapers/vbacode/debug.asp
' Source/date:  Bonnie Campbell, February 23, 2017 for NCPN tools
' Adapted:      -
' Revisions:    BLC, 2/23/2017 - initial version
' ---------------------------------

'-----------------------------------------------------------------------
' Constants
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' Declarations
'-----------------------------------------------------------------------
' Current pointer to the array element of the call stack
Private mintStackPointer As Integer

' Array of procedure names in the call stack
Private masterCallStack() As String

' The number of elements to increase the array
Private Const IncrementStackSize As Integer = 10

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
' SUB:          PushCallStack
' Description:  Add current procedure name to Call Stack
' Assumptions:
'   PushCallStack is called at the beginning of procedures/functions
' Parameters:   strName - name of procedure/function (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 23, 2017 for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 2/23/2017 - initial version
' ---------------------------------
Public Sub PushCallStack(strName As String)
On Error GoTo Err_Handler

    If Not IsArray(masterCallStack) Then
        Debug.Print "not array"
        GoTo Exit_Handler
    End If

    'handle first pass to enable first push to stack
    If IsArrayEmpty(masterCallStack) Then
        ReDim Preserve masterCallStack(IncrementStackSize)
        Debug.Print "empty array"
    End If
    
        ' Verify stack array can handle the current array element
        If mintStackPointer > UBound(masterCallStack) Then
            ' If array has not been defined, initialize the error handler
            If Err.Number = 9 Then
                'ErrorHandlerInit
            Else
                ' Increase the size of the array to not go out of bounds
                ReDim Preserve masterCallStack(UBound(masterCallStack) + IncrementStackSize)
            End If
        End If
    
    On Error GoTo 0
    
    masterCallStack(mintStackPointer) = strName
    
    ' Increment pointer to next element
    mintStackPointer = mintStackPointer + 1

Exit_Handler:
    'cleanup
    Dim proc As Variant
    Dim strproc As String
    For Each proc In masterCallStack
        strproc = strproc & proc
    Next
    Debug.Print strproc
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PushCallStack[mod_Error_Handling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ErrorHandlerInit
' Description:  Add current procedure name to Call Stack
' Assumptions:
'   PushCallStack is called at the beginning of procedures/functions
' Parameters:   strName - name of procedure/function (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 23, 2017 for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 2/23/2017 - initial version
' ---------------------------------
Public Sub ErrorHandlerInit()
On Error GoTo Err_Handler

    mintStackPointer = 1
    ReDim masterCallStack(1 To IncrementStackSize)

Exit_Handler:
    'cleanup
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ErrorHandlerInit[mod_Error_Handling])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          PopCallStack
' Description:  Remove current procedure name from Call Stack
' Assumptions:
'   PopCallStack is called at the end of procedures/functions
' Parameters:   strName - name of procedure/function (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February 23, 2017 for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 2/23/2017 - initial version
' ---------------------------------
Public Sub PopCallStack(strName As String)
On Error GoTo Err_Handler

    If mintStackPointer <= UBound(masterCallStack) Then
        masterCallStack(mintStackPointer) = ""
    End If
    
    ' Reset pointer to previous element
    mintStackPointer = mintStackPointer - 1

Exit_Handler:
    'cleanup
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopCallStack[mod_Error_Handling])"
    End Select
    Resume Exit_Handler
End Sub