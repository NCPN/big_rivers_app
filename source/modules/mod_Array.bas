Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Array
' Level:        Framework module
' Version:      1.00
' Description:  array functions & procedures
'
' Source/date:  Bonnie Campbell, 9/19/2016
' Revisions:    BLC, 9/19/2016 - 1.00 - initial version
' =================================

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Subroutines & Functions
' ---------------------------------


' ---------------------------------
' SUB:          StringTo2DArray
' Description:  array conversion to recordset actions
' Assumptions:  Array to convert is two dimensional
' Parameters:   str - string to change into array (string)
'               delimiter1 - first split delimiter (string)
'               delimiter2 - second split delimiter (string)
' Returns:      2D array
' Throws:       none
' References:
'   vgarcia, May 16, 2002
'   http://www.sitepoint.com/forums/showthread.php?60433-How-to-Convert-a-String-into-a-Multidimensional-Array
' Source/date:  Bonnie Campbell, September 19, 2016 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 9/20/2016 - initial version
' ---------------------------------
Public Function StringTo2DArray(str As String, delimiter1 As String, _
                                delimiter2 As String) As Variant
On Error GoTo Err_Handler

    If Len(str) = 0 Then GoTo Err_Handler
    
    Dim aryDimOne() As String
    Dim aryDimTwo() As Variant
    Dim tempArray() As Variant
    Dim i As Integer, j As Integer
    
    'first dimension split
    aryDimOne = Split(str, delimiter1)

    For i = 0 To UBound(aryDimOne) - 1
    
        'ReDim aryDimTwo(UBound(aryDimOne) - 1, UBound(Split(aryDimOne(i), delimiter2)))
        
        For j = 0 To UBound(aryDimOne)
        
            'temp array
            tempArray = Split(aryDimOne(i), delimiter2)
            
            'second dimension split
'            aryDimTwo(i) = tempArray
'
'            strto2darray = aryDimTwo
    
        Next
        
    Next
 
 
 
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateMultiArray[mod_Array])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ArrayToRecordset
' Description:  array conversion to recordset actions
' Assumptions:  Array to convert is two dimensional
' Parameters:   aryFields - recordset field names (string array)
'               aryData - recordset data (variant array)
' Returns:      ADODB recordset containing array data
' Throws:       none
' References:   -
'   Vishesh, 17 September, 2010
'   http://excelexperts.com/Array-To-ADO-Recordset
' Source/date:  Bonnie Campbell, September 19, 2016 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 9/19/2015 - initial version
' ---------------------------------
'aryFields() As String, aryData() As Variant) As DAO.Recordset 'ADODB.Recordset
Public Function ArrayToRecordset(aryFields() As Variant, aryData() As Variant, _
                                delimiter As String) As DAO.Recordset 'ADODB.Recordset
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rsData As DAO.Recordset 'ADODB.Recordset
    Dim recordString As String
    Dim aryRecord As Variant
    Dim aryCols As String
    Dim i As Integer
    Dim j As Integer
 
'    ReDim aryRecord(1 To 1, 1 To UBound(aryData, 2))
 
    Set db = CurrentDb
    
    Dim LB As Integer, UB As Integer
    
    CreateTempTable "usys_temp_rs", aryFields 'aryData
    
    Set rsData = db.OpenRecordset("usys_temp_rs")

    For i = 0 To UBound(aryFields)
    
 '       aryCols(i) = Split(aryFields(i), delimiter)(0)
    
    Next
    
'    AddRecords rsData, aryCols, aryData, "|"


'    lb = LBound(aryData, 1)
'    ub = UBound(aryData, 1)
'    For i = lb To ub
'        rsData.AddNew
'
'        aryRecord = Split(aryData(i), "|")
'
'        For j = 0 To UBound(aryRecord) - 1
'
'            rsData(j) = aryRecord(j) 'aryData(j) ',i)
'
'        Next
'
'        rsData.update
'
'    Next
'      rs!ID = matrix(0, i)
'      rs!value1 = matrix(1, X)
'      rs!value2 = matrix(2, X)
'      rs!value3 = matrix(3, X)
'      rs!value4 = matrix(4, X)
'      rs!value5 = matrix(5, X)
'      rs.update
'    Next X
    
'    Set rsData = db.OpenRecordset("SELECT '';")   'ADODB.Recordset
'
'    'add fields
'    Dim aryField As Variant
'    For i = 1 To UBound(aryField, 2)
' 'adVarChar
'
'        rsData.fields.Append  '.CreateField(aryField(1, i), dbText, 50)
'
'        'aryField(1, i), dbText, 500  'adVarChar, 500
'
'    Next i
'
'    'open rs for data
'    rsData.OpenRecordset '.Open
'
'    'add data
'    For i = 1 To UBound(arrData, 1)
'
'        For j = 1 To UBound(arrData, 2)
'
'            arrRecord(1, j) = arrData(i, j)
'
'        Next j
'
'        rsData.AddNew 'arrField, aryRecord
'
'        For f = 0 To UBound(aryFields) - 1
'            rsData.fields(f).Value = aryRecord(1, f)
'        Next f
'
'        rsData.update
'
'    Next i
 
    Set ArrayToRecordset = rsData
 
    'cleanup
'    Erase aryRecord
    Set rsData = Nothing
 
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ArrayToRecordset[mod_Array])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ArrayReplace
' Description:  iterate through array elements replacing portions of the element string
' Assumptions:  Array to convert is a one dimensional string
'               regex to remove
'                   #s:     "[0-9]+-"
'                   text:   "-([a-z]|\s|_|&|/|/.|[a-z],\s[a-z]+)+" (used for isolating SOP #s)
' Parameters:   ary - array data (variant, string or array)
'               UseRegEx - whether to use regex or not (boolean)
'               Remove - item to replace (string)
'               ReplaceWith - item to replace with (optional, string)
'               iStart - array position to start modification (optional, integer)
'               iEnd - array position to stop modification (optional, integer)
' Returns:      array of items after modification (variant array)
' Throws:       none
' References:
'   osknows, April 18, 2013
'   http://stackoverflow.com/questions/16084909/vba-multiple-matches-within-one-string-using-regular-expressions-execute-method
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 1/19/2017 - initial version
' ---------------------------------
Public Function ArrayReplace(ary As Variant, _
            Remove As String, _
            UseRegEx As Boolean, _
            Optional ReplaceWith As String = "", _
            Optional iStart As Integer = 0, _
            Optional iEnd As Integer = 0 _
            ) As Variant
On Error GoTo Err_Handler

    If UseRegEx Then
        'Dim rgx As New Regex(pattern)
      'Dim result As String = rgx.Replace(input, replacement)
        Dim RegEx As New RegExp
        
        RegEx.Pattern = Remove
        RegEx.Global = True     'replace globally!
        RegEx.IgnoreCase = True
        
    End If
    
    If Not IsArray(ary) Then
        Dim strNew As String
        
        If UseRegEx Then
            strNew = RegEx.Replace(ary, ReplaceWith)
        Else
            strNew = Replace(ary, Remove, ReplaceWith)
        End If
        
        ArrayReplace = strNew
        
    Else
    
        Dim aryNew As Variant '() As String
        Dim i As Integer
        
        'set default end
        If iEnd = 0 Then iEnd = UBound(ary)
        
        'iterate through array
        For i = 0 To UBound(ary)
            Select Case i
                Case Is = iStart, _
                     Is < iStart, _
                     Is < iEnd, _
                     Is = iEnd
                     
                        If UseRegEx Then
                            ary(i) = RegEx.Replace(ary(i), ReplaceWith)
                        Else
                            ary(i) = Replace(ary(i), Remove, ReplaceWith)
                        End If
                
                Case Else
                    'do nothing
            End Select
        Next
        
        'aryNew = ary
        
        ArrayReplace = ary
        
    End If

    'ArrayReplace = aryNew 'strNew

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ArrayReplace[mod_Array])"
    End Select
    Resume Exit_Handler
End Function



' ---------------------------------
' FUNCTION:     IsArrayEmpty
' Description:  tests whether the array is allocated
' Assumptions:  -
' Parameters:   ary - array data (variant, string or array)
' Returns:      True or False whether array is empty (boolean)
' Throws:       none
' References:
'   Chip Pearson, Unknown
'   http://www.cpearson.com/excel/vbaarrays.htm
' Source/date:  Bonnie Campbell, February 23, 2017 - for NCPN tools
' Adapted:  -
' Revisions:
'   BLC - 2/23/2017 - initial version
' ---------------------------------
Public Function IsArrayEmpty(ary As Variant) As Boolean
On Error GoTo Err_Handler

    Dim LB As Long
    Dim UB As Long
    
    Err.Clear
    On Error Resume Next
    If IsArray(ary) = False Then
        'not an array --> return true
        IsArrayEmpty = True
    End If
    
    ' UBound check
    UB = UBound(ary, 1)
    
    If (Err.Number <> 0) Then
        'UBound leads to an error when array is unallocated --> return true
        IsArrayEmpty = True
    Else
    
        ''''''''''''''''''''''''''''''''''''''''''
        ' Chip Pearson:
        ' On rare occassion, under circumstances I
        ' cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBound is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(ary)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsEmptyArray[mod_Array])"
    End Select
    Resume Exit_Handler
End Function