Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Validation
' Level:        Framework module
' Version:      1.06
' Description:  string functions & procedures
'
' Source/date:  Bonnie Campbell, 2/10/2015
' Revisions:    BLC - 2/10/2015 - 1.00 - initial version
'               BLC - 11/12/2015 - 1.01 - added IsAlphaNumDashSlash(), IsAlphaNumDashUnder(),
'                                         IsWord(), IsParagraph(),
'                                         & verifications via VerifyString()
'               BLC - 4/4/2016 - 1.02 - added IsInArray(), updated ValidString(),
'                                       replaced Exit_Function w/ Exit_Handler
'               BLC - 5/20/2016 - 1.03 - added IsTypeMatch()
'               BLC - 6/7/2016  - 1.04 - added IsPhone()
'               BLC - 6/13/2016 - 1.05 - added IsNothing()
'               BLC - 6/22/2016 - 1.06 - removed IsNothing(), added IsZLS(), IsZero()
' =================================

' ---------------------------------
' FUNCTION:     IsBlank
' Description:  Counts the number of instances of character(s) in a string
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
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsBlank[mod_Validation])"
    End Select
    Resume Exit_Handler
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
' Source/date:
'
'
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
'   BLC - 11/12/2015 - added alphanumdashslashspace, alphanumdashspace, alphanumspace, alphanumdashunder,
'                      alphanumdashunderspace
'   BLC - 4/4/2016   - added name & email cases
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
        Case "alphadashunderscore"
            blnIsValid = IsAlphaDashUnderscore(Trim(strInspect))
        Case "alphanumspace"
            blnIsValid = IsAlphaNum(Trim(Replace(strInspect, " ", "")))
        Case "numeric"
            blnIsValid = IsNumeric(Trim(strInspect))
        Case "alphanumdash"
            blnIsValid = IsAlphaNumDash(Trim(strInspect))
        Case "alphaspace"
            blnIsValid = IsAlphaNumDash(Replace(strInspect, " ", ""))
        Case "alphanumdashspace"
            blnIsValid = IsAlphaNumDash(Trim(Replace(strInspect, " ", "")))
        Case "alphanumdashunder"
            blnIsValid = IsAlphaNumDashUnder(Trim(strInspect))
        Case "alphanumdashunderspace"
            blnIsValid = IsAlphaNumDashUnder(Trim(Replace(strInspect, " ", "")))
        Case "alphanumdashslash"
            blnIsValid = IsAlphaNumDashSlash(Trim(strInspect))
        Case "alphanumdashslashspace"
            blnIsValid = IsAlphaNumDashSlash(Trim(Replace(strInspect, " ", "")))
        Case "name"
            blnIsValid = IsName(Trim(strInspect))
        Case "email"
            blnIsValid = IsEmail(Trim(strInspect))
        Case "word"
            blnIsValid = IsWord(Trim(strInspect))
        Case "paragraph"
            blnIsValid = IsParagraph(Trim(Replace(strInspect, " ", "")))
    End Select

    ValidateString = blnIsValid

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountInString[mod_Validation])"
    End Select
    Resume Exit_Handler
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
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlpha[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsAlphaDashUnderscore
' Description:  Checks if string includes only letter, underscore, dash
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is a name), False (string contains non-word characters)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Function IsAlphaDashUnderscore(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim strPattern As String
    strPattern = "^[a-zA-Z_-]+$"

    IsAlphaDashUnderscore = IsRegExpMatch(strInspect, strPattern)
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaDashUnderscore[mod_Validation])"
    End Select
    Resume Exit_Handler
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
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaNum[mod_Validation])"
    End Select
    Resume Exit_Handler
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
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaNumDash[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsAlphaNumDashUnder
' Description:  Checks if string is alphanumeric w/ or w/o dash or underscore
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is alphanum), False (string contains non-alphanumeric/dash/
'                         underscore characters)
' Throws:       none
' References:   none
' Source/date:
' si_the_geek, March 30, 2007
' http://www.vbforums.com/showthread.php?460464-RESOLVED-is-there-a-method-like-quot-isAlphabetic-quot
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/12/2015 - initial version
' ---------------------------------
Function IsAlphaNumDashUnder(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsAlphaNumDashUnder = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case "0" To "9"
        Case "-"
        Case "/"
        Case "_"
        Case Else
          IsAlphaNumDashUnder = False
          Exit For
      End Select
    
    Next i
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaNumDashUnder[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsAlphaNumDashSlash
' Description:  Checks if string is alphanumeric w/ or w/o dash or slash
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is alphanum), False (string contains non-alphanumeric/dash/slash characters)
' Throws:       none
' References:   none
' Source/date:
' si_the_geek, March 30, 2007
' http://www.vbforums.com/showthread.php?460464-RESOLVED-is-there-a-method-like-quot-isAlphabetic-quot
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/12/2015 - initial version
' ---------------------------------
Function IsAlphaNumDashSlash(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsAlphaNumDashSlash = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case "0" To "9"
        Case "-"
        Case "/"
        Case Else
          IsAlphaNumDashSlash = False
          Exit For
      End Select
    
    Next i
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsAlphaNumDashSlash[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsName
' Description:  Checks if string is a name (can contain: letter, period, dash, space)
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is a name), False (string contains non-word characters)
' Throws:       none
' References:   none
' Source/date:
'   Matthew Scharley, November 8, 2008
'   http://stackoverflow.com/questions/275160/regex-for-names
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Function IsName(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim strPattern As String
    strPattern = "^[A-Z]'?[- a-zA-Z]([a-zA-Z])*$"

    IsName = IsRegExpMatch(strInspect, strPattern)
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsName[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsEmail
' Description:  Checks if string is an email address
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is a valid email address),
'                         False (string isn't a valid email address)
' Throws:       none
' References:   none
' Notes:
'               Pattern => ^[a-zA-Z]+(?:\s+[a-zA-Z]+)*$
'               Explanation:
'                ^          Start of string
'                [a-zA-Z]   Any character in the class a to z or A to Z
'                +          One or more repititions
'                (?:   )    Match expresion but don't capture
'                \s+        Whitespace, One or more repititions
'                *          Zero or more repititions
'                $          End of string
'
'   In the end, the simplest pattern is best to avoid rejecting valid
'   email addresses -- e.g. w/ tags email+tag@example.com
'               Pattern => ^.+@.+\..+$    (originally /.+@.+\..+/i)
'               Explanation:
'                   ^  start
'                   .+ any character(s) - one or more times
'                   @  followed by @
'                   .+ any character(s) - one or more times
'                   \. followed by a period (.)
'                   .+ any character(s) - one or more times
'                   $  end
' Source/date:
'   David Celis, September 6, 2012
'   https://davidcel.is/posts/stop-validating-email-addresses-with-regex/
'   Chris Nielson July 17, 2012
'   http://stackoverflow.com/questions/11501860/regular-expression-pattern-to-validate-name-field
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Function IsEmail(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim strPattern As String
    strPattern = "^.+@.+\..+$"

    IsEmail = IsRegExpMatch(strInspect, strPattern)
    
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsEmail[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsWord
' Description:  Checks if string is alphabetic
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is a word), False (string contains non-word characters)
' Throws:       none
' References:   none
' Source/date:
'
'
' Adapted:      Bonnie Campbell, November 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/12/2015 - initial version
' ---------------------------------
Function IsWord(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsWord = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case "-"
        Case Else
          IsWord = False
          Exit For
      End Select
    
    Next i
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsWord[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsParagraph
' Description:  Checks if string is alphabetic
' Assumptions:  -
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is alpha), False (string contains non-alpha characters)
' Throws:       none
' References:   none
' Source/date:
'
'
' Adapted:      Bonnie Campbell, November 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 11/12/2015 - initial version
' ---------------------------------
Function IsParagraph(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim i As Integer
    
    'default
    IsParagraph = True
    
    For i = 1 To Len(Trim(strInspect))
      
      Select Case Mid$(Trim(strInspect), i, 1)
        Case "A" To "Z", "a" To "z"
        Case "0" To "9"
        Case "-", "/", "_", ".", "?", "!"
        Case Else
          IsParagraph = False
          Exit For
      End Select
    
    Next i
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsParagraph[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsPhone
' Description:  Checks if string is a valid phone #
' Assumptions:  Assumes phone #s are U.S. numbers (7 or 10 digits w/ area code)
' Parameters:   strInspect - string to check
' Returns:      boolean - True (string is valid phone #), False (string contains non-phone characters)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 7, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/7/2016 - initial version
' ---------------------------------
Function IsPhone(strInspect As String) As Boolean
On Error GoTo Err_Handler:

    Dim strPhone As String
    
    'remove spaces, remove () & -, and
    strPhone = Replace(Replace(Replace(InternalTrim(strInspect), "(", ""), ")", ""), "-", "")
    
    'check length is 7 or 10
    If Len(strPhone) = 7 Or Len(strPhone) = 10 Then
        'check if remainder is numeric
        IsPhone = IsNumeric(strPhone)
    Else
        IsPhone = False
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsPhone[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsRegExpMatch
' Description:  Checks if string is a match for the regular expression pattern
' Assumptions:  -
' Parameters:   strInspect - string to check
'               strPattern - pattern to check against (string)
' Returns:      boolean - True (string matches), False (string does not match)
' Throws:       none
' References:   Microsoft VBScript Regular Expressions 5.5 (added reference)
' Source/date:
'   RICHA, March 31, 2014
'   https://blog.udemy.com/vba-regex/
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Function IsRegExpMatch(strInspect As String, strPattern As String) As Boolean
On Error GoTo Err_Handler:

    Dim oRegExp As VBScript_RegExp_55.RegExp
    
    Set oRegExp = CreateObject("vbscript.regexp")
    
    With oRegExp
        .Global = True
        .IgnoreCase = True
        .pattern = strPattern
        
        IsRegExpMatch = .test(strInspect)

    End With
                
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsRegExpMatch[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsInArray
' Description:  Checks if string is found in the supplied array
' Assumptions:  -
' Parameters:   strFind - string to check
'               aryLookIn - array to look in
' Returns:      boolean - True (string is found), False (string isn't found)
' Throws:       none
' References:   none
' Source/date:
'   Jimmy Pena, June 20, 2012
'   http://stackoverflow.com/questions/11109832/how-to-find-if-an-array-contains-a-string
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Function IsInArray(strFind As String, aryLookIn As Variant) As Boolean
On Error GoTo Err_Handler

  IsInArray = (UBound(filter(aryLookIn, strFind)) > -1)
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsInArray[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsBetween
' Description:  Checks if value is between supplied bounding values/limits
' Assumptions:  -
' Parameters:   iValue - value to check (variant)
'               lowBound - lower limit (double)
'               highBound - upper limit (double)
'               inclusive - whether the lower & upper limits should be included (boolean)
' Returns:      boolean - True (value is between limits), False (value is outside limits)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Function IsBetween(iValue As Variant, lowBound As Double, highBound As Double, inclusive As Boolean) As Boolean
On Error GoTo Err_Handler

    Dim isOK As Boolean
    'default
    isOK = False
    
    'ensure numeric
    If Not IsNumeric(iValue) Then GoTo Exit_Handler

    If inclusive Then
        Select Case iValue
            'rejects --> all result in isOK = false (no change)
            Case Is < lowBound
            Case Is > highBound
            
            'valid cases
            Case Is = lowBound
                isOK = True
            Case Is = highBound
                isOK = True
            Case Is > lowBound And (iValue < highBound)
                isOK = True
        End Select
    Else
        Select Case iValue
            'rejects --> all result in isOK = false (no change)
            Case Is < lowBound
            Case Is > highBound
            Case Is = lowBound
            Case Is = highBound
            
            'valid cases
            Case Is > lowBound And (iValue < highBound)
                isOK = True
        End Select
    End If
    
    IsBetween = isOK
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsBetween[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsTypeMatch
' Description:  Checks if value is or can be converted to the data type noted
'               Relies on attempting to convert, if it fails via type mismatch or otherwise false is returned
' Assumptions:  -
' Note:
'               Value     Variant type          Value     Variant type
'               0     Empty (unitialized)       10     Error Value
'               1     Null (no valid data)      11     Boolean
'               2     Integer                   12     Variant (only used with arrays of variants)
'               3     Long Integer              13     Data access object
'               4     Single                    14     Decimal value
'               5     Double                    17     Byte
'               6     Currency                  36     User Defined Type
'               7     Date                      8192           Array
'               8     String
'               9     Object
'
' Parameters:   iValue - value to check (variant)
'               dataType - data type name (string)
' Returns:      boolean - True (value is or can be converted), False (value isn't/can't be converted to the data type passed in)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, May 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/20/2016 - initial version
' ---------------------------------
Function IsTypeMatch(iValue As Variant, DataType As String) As Boolean
On Error GoTo Err_Handler

    Dim isOK As Boolean
    Dim result As Variant
    
    'default
    isOK = True
    
    'check type
    Select Case DataType
        Case "boolean"  '0 or 1, yes/no values are mismatches
            result = CBool(iValue)
        Case "byte"     '0 or 1, yes/no values are mismatches
            result = CByte(iValue)
        Case "number"
            If Not IsNumeric(iValue) Then isOK = False
        Case "integer"
            result = CInt(iValue)
        Case "long"
            result = CLng(iValue)
        Case "double"
            result = CDbl(iValue)
        Case "single"
            result = CSng(iValue)
        Case "decimal"
            result = CDec(iValue)
        Case "string"
            result = CStr(iValue)
        Case "date"
            result = CDate(iValue)
        Case "currency"
            result = CDate(iValue)
        Case Else
            isOK = False
    End Select
    
Exit_Handler:
    IsTypeMatch = isOK
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case "13" 'RunTime Error 13: Type Mismatch
        isOK = False
        Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsTypeMatch[mod_Validation])"
    End Select
    'fail on error
    isOK = False
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsZero
' Description:  Checks if value is zero
' Assumptions:  -
' Parameters:   chkValue - value to check
' Returns:      boolean - True (value is zero), False (value isn't zero)
' Throws:       none
' References:   none
' Source/date:  -
' Adapted:      Bonnie Campbell, June 22, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Function IsZero(ByVal chkValue) As Boolean
On Error GoTo Err_Handler

  Dim blnZero As Boolean
  
  'default
  blnZero = False
  
  If chkValue = 0 Then blnZero = True

  IsZero = blnZero
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsZero[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsZLS
' Description:  Checks if string is a zero length string (ZLS)
' Assumptions:  -
' Parameters:   strCheck - string to check
' Returns:      boolean - True (string is ZLS), False (string isn't ZLS)
' Throws:       none
' References:   none
' Source/date:  -
' Adapted:      Bonnie Campbell, June 22, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/22/2016 - initial version
' ---------------------------------
Function IsZLS(strCheck As String) As Boolean
On Error GoTo Err_Handler

  Dim blnZLS As Boolean
  
  'default
  blnZLS = False
  
  If Len(strCheck) = 0 Then blnZLS = True

  IsZLS = blnZLS
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsZLS[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function