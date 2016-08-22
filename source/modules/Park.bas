Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Park
' Level:        Framework class
' Version:      1.01
'
' Description:  Record Park object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 8/8/2016  - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_Code As String
Private m_Name As String
Private m_State As String
Private m_IsActiveForProtocol As Boolean

'---------------------
' Events
'---------------------
Public Event InvalidParkID(Value As Long)
Public Event InvalidParkCode(Value As String)
Public Event InvalidPark(Value As String)
Public Event InvalidParkState(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let Code(Value As String)
    If Len(Trim(Value)) = 4 Then
        m_Code = Value
    Else
        RaiseEvent InvalidParkCode(Value)
    End If
End Property

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Name(Value As String)
    'max length = 25
    If Len(Trim(Value)) < 26 Then
        m_Name = Value
    Else
        RaiseEvent InvalidPark(Value)
    End If
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let State(Value As String)
    'max length = 2
    If Len(Trim(Value)) < 3 Then
        m_State = Value
    Else
        RaiseEvent InvalidParkState(Value)
    End If
End Property

Public Property Get State() As String
    State = m_State
End Property

Public Property Let IsActiveForProtocol(Value As Boolean)
    m_IsActiveForProtocol = Value
End Property

Public Property Get IsActiveForProtocol() As Boolean
    IsActiveForProtocol = m_IsActiveForProtocol
End Property

'---------------------
' Methods
'---------------------

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

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Park])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Park])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'
'    Set db = CurrentDb
'
'    'record Parks must have:
'    strSQL = "INSERT INTO Park(ParkCode, ParkName, ParkState, ActiveForProtocol) VALUES " _
'                & "('" & Me.Code & "','" & Me.Name & "','" _
'                & Me.State & "'," & Me.IsActiveForProtocol & ");"
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim template As String
    
    template = "i_park"
    
    Dim params(0 To 6) As Variant
    
    With Me
        params(0) = "Park"
        params(1) = .Code
        params(2) = .Name
        params(3) = .State
        params(4) = .IsActiveForProtocol
        
        If IsUpdate Then
            template = "u_park"
            params(5) = .ID
        End If
        
        .ID = SetRecord(template, params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Park])"
    End Select
    Resume Exit_Handler
End Sub