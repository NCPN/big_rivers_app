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
Public Event InvalidParkID(value As Long)
Public Event InvalidParkCode(value As String)
Public Event InvalidPark(value As String)
Public Event InvalidParkState(value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let Code(value As String)
    If Len(Trim(value)) = 4 Then
        m_Code = value
    Else
        RaiseEvent InvalidParkCode(value)
    End If
End Property

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Name(value As String)
    'max length = 25
    If Len(Trim(value)) < 26 Then
        m_Name = value
    Else
        RaiseEvent InvalidPark(value)
    End If
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let State(value As String)
    'max length = 2
    If Len(Trim(value)) < 3 Then
        m_State = value
    Else
        RaiseEvent InvalidParkState(value)
    End If
End Property

Public Property Get State() As String
    State = m_State
End Property

Public Property Let IsActiveForProtocol(value As Boolean)
    m_IsActiveForProtocol = value
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

    Dim Template As String
    
    Template = "i_park"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "Park"
        Params(1) = .Code
        Params(2) = .Name
        Params(3) = .State
        Params(4) = .IsActiveForProtocol
        
        If IsUpdate Then
            Template = "u_park"
            Params(5) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
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