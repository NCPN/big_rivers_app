Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Strings
' Level:        Framework module
' Version:      1.01
' Description:  String related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/12/2016 - 1.01 - added Unicode strings, GetUnicode()
' =================================

'---------------------
' Declarations
'---------------------
' Hex Unicode constants --> use w/ ChrW() or StringFromCodepoint() if supplementary unicode characters (codepoints) outside ChrW() range
'                           Ranges: Chr (0-255) & ChrW (-32768 - 65535), StringFromCodepoint(all)
'                           Out of Range? --> Argument exception error # 5 occurs
'                           long values are listed below
'---------------------
Public Const uSpiral = &HAA5C               '-21924 (Cham Punctuation Spiral)
Public Const uAmpersand = &H26              '38     doesn't work :(
Public Const uMu = &H3BC                    '956    microns
Public Const uRArrow = &H2192               '8594   right arrow c.f. https://en.wikipedia.org/wiki/Arrow_(symbol)
Public Const uDArrow = &H2193               '8495   down arrow
Public Const uLessThanOrEqual = &H2264      '8804
Public Const uGreaterThanOrEqual = &H2265   '8805
Public Const uCircle1 = &H2460              '9312
Public Const uCircle2 = &H2461              '9313
Public Const uCircle3 = &H2462              '9314
Public Const uBullet = &H25CF               '9679
Public Const uCheckboxEmpty = &H2610        '9744
Public Const uCheckboxCheck = &H2611        '9745
Public Const uCheckboxX = &H2612            '9746
Public Const uCheck = &H2714                '10004
Public Const uCircleFilled1 = &H278A        '10122
Public Const uCircleFilled2 = &H278B        '10123
Public Const uCircleFilled3 = &H278C        '10124
Public Const uLTriangle = &H2BC7            '11207  left-pointing triangle
Public Const uRTriangle = &H2BC8            '11208  right-pointing triangle
Public Const uMtn = &H30D8                  '12504  mountain (Katakana Letter He)
Public Const uMtnSun = &H30DA               '12506  mountain & sun (Katakana Letter Pe)
'--- use StringFromCodepoint() from here ---
Public Const uMtnSunrise = &H1F304          '127748 mountain sunrise
Public Const uLightningCloud = &H1F329      '127785
Public Const uGrass = &H1F33E               '127806 grass(ear of rice)
Public Const uHerb = &H1F33F                '127807
Public Const uCamping = &H1F3D5             '127957
Public Const uNatlPark = &H1F3DE            '127966 path & tree
Public Const uDesert = &H1F3DC              '127964 cactus & sun
Public Const uBlkPennant = &H1F3F1          '127985 right facing black pennant
Public Const uWhtPennant = &H1F3F2          '127986 right facing white pennant
Public Const uTag = &H1F3F7                 '127991 marking tag
Public Const uSpeech = &H1F4AC              '128172
Public Const uCow = &H1F404                 '128004
Public Const uSnail = &H1F40C               '128012
Public Const uPawPrints = &H1F43E           '128062
Public Const uCalendarTearOff = &H1F4C6     '128198
Public Const uClipboard = &H1F4CB           '128203
Public Const uPushpin = &H1F4CC             '128204
Public Const uPushpinRnd = &H1F4CD          '128205 round-head pushpin
Public Const uPaperclip = &H1F4CE           '128206
Public Const uRuler = &H1F4CF               '128207 straight ruler
Public Const uRulerTriangle = &H1F4D0       '128208 roofing triangle
Public Const uMemo = &H1F4DD                '128221
Public Const uCamera = &H1F4F7              '128247 camera icon
Public Const uFlashCamera = &H1F4F8         '128248 camera w/flash icon
Public Const uLocked = &H1F512              '128274
Public Const uUnlocked = &H1F513            '128275
Public Const uHammer = &H1F528              '128296
Public Const uPencil = &H1F589              '128393
Public Const uFolder = &H1F5C0              '128448
Public Const uFolderOpen = &H1F5C1          '128449
Public Const uNotepad = &H1F5CA             '128458
Public Const uDocument = &H1F5CE            '128462
Public Const uCalendarSpiral = &H1F5D3      '128467
Public Const uRefresh = &H1F5D8             '128472 clockwise left & right arrows
Public Const uCancel = &H1F5D9              '128473 X
Public Const uComment = &H1F5E9             '128489 speech bubble
Public Const uDelete = &H1F5F4              '128500 script ballot X
Public Const uCheckItem = &H1F5F9           '128505 checked ballot box
Public Const uPedestrian = &H1F6B6          '128694
Public Const uCancel2 = &H1F5D9             '128473
Public Const uHammerWrench = &H1F6E0        '128736 crossed hammer and wrench
Public Const uRHArrow = &H1F846             '129094 heavy right arrow
Public Const uLHArrow = &H1F844             '129092 heavy left arrow
Public Const uLizard = &H1F98E              '129422

' =================================
' FUNCTION:     ReplaceString
' Description:  Replaces a substring in a string with another
' Parameters:   strTextIn - string to work on
'               strFind - string to find
'               strReplace - string to replace with
'               fCaseSensitive - True for case sensitive search (default=False)
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, date unknown
' Revisions:    John R. Boetsch, 5/17/2006 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function ReplaceString(strTextIn As String, strFind As String, _
    strReplace As String, Optional fCaseSensitive As Boolean = False) As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intPos As Integer
    Dim intCaseSensitive As Integer

    ' Convert the case-sensitive boolean to the comparison constant (1=binary, 2=textual)
    intCaseSensitive = fCaseSensitive + 1

    strTemp = strTextIn
    intPos = InStr(1, strTemp, strFind, intCaseSensitive)

    Do While intPos > 0
        strTemp = Left$(strTemp, intPos - 1) & strReplace & Mid$(strTemp, intPos + Len(strFind))
        intPos = InStr(intPos + Len(strReplace), strTemp, strFind, intCaseSensitive)
    Loop

    ReplaceString = strTemp

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReplaceString[mod_Strings])"
    End Select
    Resume Exit_Function
End Function

' =================================
' FUNCTION:     ChangeDelimiter
' Description:  Replaces delimiters in an input string; default is to change double-quotes
'               to single quotes
' Parameters:   strInputText - string to work on
'               strCurrDelimiter - current delimiter in the string (default: double-quote)
'               strNewDelimiter - desired replacement delimiter (default: single-quote)
' Returns:      modified string
' Throws:       none
' References:   ReplaceString
' Source/date:  John R. Boetsch, 5/17/2006
' Revisions:    JRB, 5/17/2006
'               BLC, 4/30/2015 - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function ChangeDelimiter(strInputText As String, _
    Optional strCurrDelimiter As String = """", _
    Optional strNewDelimiter As String = "'") As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    
    ' Call the replace string function, specifying the delimiter and no case-sensitive search
    strTemp = ReplaceString(strInputText, strCurrDelimiter, strNewDelimiter)
    ChangeDelimiter = strTemp

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeDelimiter[mod_Strings])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     InsertSpace
' Description:  Inserts a space between capitalized letters
' Parameters:   str - string to inspect
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  theDBguy, May 20, 2010
'               http://www.utteraccess.com/forum/Split-string-capital-le-t1945127.html
' Adapted:      Bonnie Campbell, June 17, 2014
' Revisions:    BLC, 6/17/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_String, added error handling
' ---------------------------------
Public Function InsertSpace(str As String) As String
     
    On Error GoTo Err_Handler
     
     Dim strTemp As String
     Dim strChar As String
     Dim intLen As Integer
     
     If str > "" Then
          For intLen = 1 To Len(str)
               strChar = Mid(str, intLen, 1)
               If Asc(strChar) >= 65 And Asc(strChar) <= 90 Then
                    strTemp = strTemp & " " & strChar
               Else
                    strTemp = strTemp & strChar
               End If
          Next
     End If
        
     InsertSpace = strTemp
     
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - InsertSpace[mod_Strings])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     CountInString
' Description:  Counts the number of instances of character(s) in a string
' Assumptions:  -
' Parameters:   strInspect - string to check
'               strFind - string to count
' Returns:      count - number o instances strFind is found in strInspect
' Throws:       none
' References:   none
' Source/date:
'
' http://stackoverflow.com/questions/5193893/count-specific-character-occurrences-in-string
' Scott Huish, June 20, 2011
' http://www.mrexcel.com/forum/excel-questions/558568-count-occurrence-string-within-string-using-visual-basic-applications.html
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/7/2015 - initial version
'   BLC, 5/1/2015 - integrated into Invasives Reporting tool
' ---------------------------------
Public Function CountInString(ByVal strInspect As String, ByVal strFind As String) As Integer
On Error GoTo Err_Handler:
     Dim Count As Integer

    'default
    Count = 0
    
    If Len(strInspect) > 0 Then
        Count = UBound(Split(strInspect, strFind))
    End If
    
    CountInString = Count

Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CountInString[mod_Strings])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     StringFromCodepoint
' Description:  Handles unicode character values beyond the ranges available to
'               Chr (0-255) & ChrW (-32768 - 65535)
'               Uses surrogate characters technique
'               Uses 2 16-bit characters to code up to &H110000 characters
'               outside the main plane of characters (the basic multilingual plane or BMP).
'               The two surrogate characters are two bunches
'               of 1024 characters coded between &HD800 and &HDBFF
'               and between &HC00 and &HDFFF.
'               Therefore only CodePoints less than &H110000 are allowed,
'               and the 2 characters to code a valid CodePoint are computed
' Assumptions:  -
' Parameters:   CodePoint - string to check
' Returns:      count - number o instances strFind is found in strInspect
' Throws:       none
' References:   none
' Source/date:
'   Ben - June 17, 2014 & user1771398 - June 18, 2014
'   http://stackoverflow.com/questions/24272148/vba-word-insertsymbol-failure-with-high-value-unicode-for-maths-symbols
'   Microsoft
'   https://msdn.microsoft.com/en-us/library/windows/desktop/dd374069(v=vs.85).aspx
' Adapted:      Bonnie Campbell, May 31, 2016 - for NCPN tools
' Revisions:
'   BLC, 5/31/2016 - initial version
' ---------------------------------
Function StringFromCodepoint(ByVal CodePoint As Long) As String
On Error GoTo Err_Handler
    If CodePoint <= &HFFFF& Then
        StringFromCodepoint = ChrW(CodePoint)
        Exit Function
    ElseIf CodePoint > &H10FFFF Or CodePoint <= 0 Then
        Err.Raise 5, "Invalid Codepoint: " & str(CodePoint)
        Exit Function
    Else
        CodePoint = CodePoint - &H10000
        Dim SurrogateLow As Long
        Dim SurrogateHigh As Long
        SurrogateLow = CodePoint And &H3FF&
        SurrogateHigh = (CodePoint - SurrogateLow) / &H400&
        StringFromCodepoint = ChrW(SurrogateHigh Or &HD800&) + ChrW(SurrogateLow Or &HDC00&)
        Exit Function
    End If
    
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - StringFromCodePoint[mod_Strings])"
    End Select
    Resume Exit_Function
End Function