Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =================================
' CLASS:        ExtArray
' Level:        Framework class
' Version:      1.00
'
' Description:  Extended array object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 8/1/2016
' References:   -
' Revisions:    BLC - 8/1/2016 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
'Dim arr() As Variant

Private m_ID As Long
Private m_Name As String
Private m_Length As Integer

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Length(Value As Integer)
    m_Length = Value
End Property

Public Property Get Length() As Integer
    Length = m_Length
End Property

Public Property Let Count(Value As Integer)
    m_Count = Value
End Property

Public Property Get Count() As Integer
    Count = m_Count
End Property

'======== Standard Methods ===========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

    Dim arr() As Variant

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Person])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    'Set m_ID = 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Person])"
    End Select
    Resume Exit_Handler
End Sub


'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Resize
' Description:  Resize an array
' Assumptions:  -
' Parameters:   length - add item to the array
' Returns:      -
' Throws:       none
' References:
'   Eric F., May 10, 2015
'   https://stackoverflow.com/questions/18097756/fastest-way-to-add-an-item-to-an-array
' Source/date:  Bonnie Campbell, August 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/1/2016 - initial version
' ---------------------------------
Public Sub Resize(Length As Integer)
On Error GoTo Err_Handler
        
    If Not arr Is Nothing Then
'        ReDim Preserve arr(length)
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Resize[ExtArray class])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          AddItem
' Description:  Add new item to the object
' Assumptions:  -
' Parameters:   item - add item to the array
' Returns:      -
' Throws:       none
' References:
'   Eric F., May 10, 2015
'   https://stackoverflow.com/questions/18097756/fastest-way-to-add-an-item-to-an-array
' Source/date:  Bonnie Campbell, August 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/1/2016 - initial version
' ---------------------------------
Public Sub AddItem(item As Variant)
On Error GoTo Err_Handler

'    Public Sub Add(Of T)(ByRef arr As T(), item As T)
        
'        If Not IsNothing(arr) Then
'            Array.Resize(arr, arr.Length + 1)
'            arr(arr.length - 1) = item
'        Else
'            ReDim arr(0)
'            arr(0) = item
'        End If
'
'    End Sub
'End Module
'
'    If Not Me Is Nothing Then
'        Me.Resize Me.Length + 1
'        arr(arr.Length - 1) = item
'    Else
'        ReDim arr(0)
'        arr(0) = item
'    End If


'If Not Len(TestArrayValue(UBound(TestArray))) > 0 Then
'            TestArrayValue(UBound(TestArray)) = Chr(intCharCode)
'        Else
'            ReDim Preserve mstrTestArray(UBound(mstrTestArray) + 1) As String
'            TestArrayValue(UBound(TestArray)) = Chr(intCharCode)
'        End If
'End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddItem[ExtArray class])"
    End Select
    Resume Exit_Handler
End Sub