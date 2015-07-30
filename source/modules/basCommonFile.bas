' =================================
' MODULE:       basCommonFile
' Description:  Standard module of common file manipulation functions from Access97
'               Developer's Handbook
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
' Revisions:    JRB, May 2006 - documentation

Option Compare Database
Option Explicit

Type tagOPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    strFilter As String
    strCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    strFile As String
    nMaxFile As Long
    strFileTitle As String
    nMaxFileTitle As Long
    strInitialDir As String
    strTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    strDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Declare Function adh_apiGetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (ofn As tagOPENFILENAME) As Boolean
Declare Function adh_apiGetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (ofn As tagOPENFILENAME) As Boolean
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

Public Const adhOFN_READONLY = &H1
Public Const adhOFN_OVERWRITEPROMPT = &H2
Public Const adhOFN_HIDEREADONLY = &H4
Public Const adhOFN_NOCHANGEDIR = &H8
Public Const adhOFN_SHOWHELP = &H10
Public Const adhOFN_NOVALIDATE = &H100
Public Const adhOFN_ALLOWMULTISELECT = &H200
Public Const adhOFN_EXTENSIONDIFFERENT = &H400
Public Const adhOFN_PATHMUSTEXIST = &H800
Public Const adhOFN_FILEMUSTEXIST = &H1000
Public Const adhOFN_CREATEPROMPT = &H2000
Public Const adhOFN_SHAREAWARE = &H4000
Public Const adhOFN_NOREADONLYRETURN = &H8000
Public Const adhOFN_NOTESTFILECREATE = &H10000
Public Const adhOFN_NONETWORKBUTTON = &H20000
Public Const adhOFN_NOLONGNAMES = &H40000
Public Const adhOFN_EXPLORER = &H80000
Public Const adhOFN_NODEREFERENCELINKS = &H100000
Public Const adhOFN_LONGNAMES = &H200000

' =================================
' FUNCTION:     adhCommonFileOpenSave
' Description:  Calls the file open/save dialog
' Parameters:   multiple, all optional (see below)
'           Flags - one or more of the adhOFN_* constants, OR'd together
'           InitialDir - the directory in which to first look
'           Filter - a set of file filters, set up by calling adhAddFilterItem
'           FilterIndex - integer indicating which filter set to use (1 if unspecified)
'           DefaultExt - extension to use if the user doesn't enter one
'                       (only useful on file saves)
'           FileName - default value for the file name text box.
'           DialogTitle - title for the dialog.
'           OpenFile - boolean(True=Open File / False=Save As)
' Returns:      the selected filename or Null if user cancels
' Throws:       none
' References:   adh_apiGetOpenFileName, adh_apiGetSaveFileName, adhTrimNull
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
' Revisions:    John R. Boetsch, May 16, 2006 - fixed strInitialDir under With block,
'               added error-traps
' =================================

Function adhCommonFileOpenSave( _
    Optional ByRef flags As Variant, _
    Optional ByVal InitialDir As Variant, _
    Optional ByVal Filter As Variant, _
    Optional ByVal FilterIndex As Variant, _
    Optional ByVal DefaultExt As Variant, _
    Optional ByVal fileName As Variant, _
    Optional ByVal DialogTitle As Variant, _
    Optional ByVal OpenFile As Variant) As Variant

    On Error GoTo Err_Handler

    Dim ofn As tagOPENFILENAME
    Dim strFilename As String
    Dim strFileTitle As String
    Dim fResult As Boolean

    If IsMissing(InitialDir) Then InitialDir = ""
    If IsMissing(Filter) Then Filter = ""
    If IsMissing(FilterIndex) Then FilterIndex = 1
    If IsMissing(flags) Then flags = 0&
    If IsMissing(DefaultExt) Then DefaultExt = ""
    If IsMissing(fileName) Then fileName = ""
    If IsMissing(DialogTitle) Then DialogTitle = ""
    If IsMissing(OpenFile) Then OpenFile = True

    ' Allocate string space for the returned string.
    strFilename = Left(fileName & String(256, 0), 256)
    strFileTitle = String(256, 0)

    ' Set up the data structure before you call the function
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = Application.hWndAccessApp
        .strFilter = Filter
        .nFilterIndex = FilterIndex
        .strFile = strFilename
        .nMaxFile = Len(strFilename)
        .strFileTitle = strFileTitle
        .nMaxFileTitle = Len(strFileTitle)
        .strTitle = DialogTitle
        .flags = flags
        .strDefExt = DefaultExt
        .strInitialDir = InitialDir

        ' Didn't think that most people would want to deal with these options
        .hInstance = 0
        .strCustomFilter = String(255, 0)
        .nMaxCustFilter = 255
        .lpfnHook = 0
    End With

    ' This will pass the desired data structure to the
    '   Windows API, which will in turn uses it to display
    '   the Open/Save As dialog
    If OpenFile Then
        fResult = adh_apiGetOpenFileName(ofn)
    Else
        fResult = adh_apiGetSaveFileName(ofn)
    End If

    If fResult Then
        ' You might care to check the Flags member of the
        '   structure to get information about the chosen file.
        '   In this example, if you bothered to pass in a
        '   value for Flags, we'll fill it in with the outgoing
        '   Flags value
        If Not IsMissing(flags) Then flags = ofn.flags
        adhCommonFileOpenSave = adhTrimNull(ofn.strFile)
    Else
        adhCommonFileOpenSave = Null
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (adhCommonFileOpenSave)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     adhAddFilterItem
' Description:  Modifies the file filter value by appending the description (like
'               "Databases"), a null character, the skeleton (like "*.mdb; *.mda")
'               and a final null character.
' Parameters:   strFilter - existing file filter
'               strDescription - new filter description
'               varItem - new filter
' Returns:      new file filter
' Throws:       none
' References:   none
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
' Revisions:    John R. Boetsch, May 17, 2006 - documentation and error-trapping
' =================================

Function adhAddFilterItem(strFilter As String, _
    strDescription As String, Optional varItem As Variant) As String

    On Error GoTo Err_Handler
    
    If IsMissing(varItem) Then varItem = "*.*"
    adhAddFilterItem = strFilter & strDescription & vbNullChar & _
        varItem & vbNullChar
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (adhAddFilterItem)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     adhTrimNull
' Description:  Trims the Null from a string returned by an API call
' Parameters:   strItem - string that contains null terminator
' Returns:      same string without null terminator
' Throws:       none
' References:   none
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
' Revisions:    John R. Boetsch, May 17, 2006 - documentation and error-trapping
' =================================

Function adhTrimNull(ByVal strItem As String) As String
    On Error GoTo Err_Handler

    Dim intPos As Integer

    intPos = InStr(strItem, vbNullChar)
    If intPos > 0 Then
        ' If the Null character is present, trim the string
        adhTrimNull = Left(strItem, intPos - 1)
    Else
        adhTrimNull = strItem
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (adhTrimNull)"
            Resume Exit_Procedure
    End Select

End Function