Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Waterway
' Level:        Framework class
' Version:      1.01
'
' Description:   Waterway (River) object related properties, events, functions & procedures
' Note:          The term Waterway is used instead of River to avoid collision with the
'                River Enum object
'
' Source/date:  Bonnie Campbell, 4/6/2015
' References:   -
' Revisions:    BLC - 4/6/2015 - 1.00 - initial version
'               BLC - 8/8/2016 - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_ParkID As Long
Private m_Name As String
Private m_Segment As String

'---------------------
' Events
'---------------------
Public Event InvalidParkID(Value As Long)
Public Event InvalidName(Value As String)
Public Event InvalidSegment(Value As String)

Public Event InvalidContactID(Value As Long)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let ParkID(Value As Long)
    m_ParkID = Value
End Property

Public Property Get ParkID() As Long
    ParkID = m_ParkID
End Property

Public Property Let Name(Value As String)
    If Len(Name) > 25 Then
        m_Name = Value
    Else
        RaiseEvent InvalidName(Value)
    End If
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let segment(Value As String)
    If Len(Name) > 5 Then
        m_Segment = Value
    Else
        RaiseEvent InvalidSegment(Value)
    End If
End Property

Public Property Get segment() As String
    segment = m_Segment
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Waterway])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Waterway])"
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
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted--using-excel-vba
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
'    ' Waterways must have:
'    strSQL = "INSERT INTO River(Park_ID, River, Segment) VALUES " _
'                & "(" & Me.ParkID & ",'" & Me.Name & "','" _
'                & Me.segment & "');"
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim template As String
    
    template = "i_waterway"
    
    Dim params(0 To 5) As Variant

    With Me
        params(0) = "River"
        params(1) = .ParkID
        params(2) = .Name
        params(3) = .segment
        
        If IsUpdate Then
            template = "u_waterway"
            params(4) = .ID
        End If
        
        .ID = SetRecord(template, params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Waterway])"
    End Select
    Resume Exit_Handler
End Sub