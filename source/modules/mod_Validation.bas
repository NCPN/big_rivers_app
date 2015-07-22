Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Validation
' Level:        Framework module
' Version:      1.00
' Description:  data validation functions & procedures
'
' Source/date:  Bonnie Campbell, 2/10/2015
' Revisions:    BLC - 2/10/2015 - 1.00 - initial version
' =================================

' ---------------------------------
' FUNCTION:     IsBlank
' Description:  Determines if an item is blank
' Assumptions:  -
' Parameters:   arg - item to check
' Returns:      boolean - True if argument is Nothing, Null, Empty, Missing or an empty string
' Throws:       none
' References:   none
' Source/date:
' Renaud Bompuis, September 9, 2009
' http://blog.nkadesign.com/2009/access-checking-blank-variables/
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function IsBlank(arg As Variant) As Boolean
On Error GoTo Err_Handler

    Select Case varType(arg)
        Case vbEmpty
            IsBlank = True
        Case vbNull
            IsBlank = True
        Case vbString
            IsBlank = (LenB(arg) = 0)
        Case vbObject
            IsBlank = (arg Is Nothing)
        Case Else
            IsBlank = IsMissing(arg)
    End Select
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsBlank[mod_Validation])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     ValidateString
' Description:  Checks if string is proper type
' Assumptions:  -
' Parameters:   strInspect - string to check
'               strType - string type (alpha, alphanum, numeric, etc.)
' Returns:      boolean - True (string is valid), False (string is invalid)
' Throws:       none
' References:   none
' Source/date:  Unknown
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function ValidateString(ByVal strInspect As String, strType As String) As Boolean
On Error GoTo Err_Handler

    Dim blnIsValid As Boolean

    'default
    blnIsValid = False

    Select Case strType
        Case "alpha"
            blnIsValid = IsAlpha(Trim(strInspect))
        Case "alphanum"
            blnIsValid = IsAlphaNum(Trim(strInspect))
        Case "numeric"
            blnIsValid = IsNumeric(Trim(strInspect))
        Case "alphanumdash"
            blnIsValid = IsAlphaNumDash(Trim(strInspect))
        Case "alphaspace"
            blnIsValid = IsAlphaNumDash(Replace(strInspect, " ", ""))
    End Select

    ValidateString = blnIsValid

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountInString[mod_Validation])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     IsAlpha
' Description:  Checks if string is alphabetic
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is alpha), False (string contains non-alpha characters)
' Throws:       none
' References:   none
' Source/date:
' si_the_geek, March 30, 2007
' http://www.vbforums.com/showthread.php?460464-RESOLVED-is-there-a-method-like-quot-isAlphabetic-quot
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Function IsAlpha(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsAlpha = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case Else
          IsAlpha = False
          Exit For
      End Select
    
    Next i
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlpha[mod_Validation])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     IsAlphaNum
' Description:  Checks if string is alphanumeric
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is alphanum), False (string contains non-alphanumeric characters)
' Throws:       none
' References:   none
' Source/date:
' si_the_geek, March 30, 2007
' http://www.vbforums.com/showthread.php?460464-RESOLVED-is-there-a-method-like-quot-isAlphabetic-quot
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Function IsAlphaNum(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsAlphaNum = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case "0" To "9"
        Case Else
          IsAlphaNum = False
          Exit For
      End Select
    
    Next i
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaNum[mod_Validation])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     IsAlphaNumDash
' Description:  Checks if string is alphanumeric w/ or w/o dash
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is alphanum), False (string contains non-alphanumeric characters)
' Throws:       none
' References:   none
' Source/date:
' si_the_geek, March 30, 2007
' http://www.vbforums.com/showthread.php?460464-RESOLVED-is-there-a-method-like-quot-isAlphabetic-quot
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Function IsAlphaNumDash(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsAlphaNumDash = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case "0" To "9"
        Case "-"
        Case Else
          IsAlphaNumDash = False
          Exit For
      End Select
    
    Next i
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaNumDash[mod_Validation])"
    End Select
    Resume Exit_Function
End Function