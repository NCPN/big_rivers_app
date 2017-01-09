Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Transducer
' Level:        Framework class
' Version:      1.01
'
' Description:  Transducer object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
'   Jeff Smith, Oct 31, 2007
'   http://weblogs.sqlteam.com/jeffs/archive/2007/10/31/sql-server-2005-date-time-only-data-types.aspx
'   Jeff Smith, August 29, 2007
'   http://weblogs.sqlteam.com/jeffs/archive/2007/08/29/SQL-Dates-and-Times.aspx
'   Michael user3480989, January 14, 2016
'   http://stackoverflow.com/questions/34783997/inserting-date-from-access-db-into-sql-server-2008r2
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 8/8/2016  - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer

Private m_EventID As Long

Private m_TransducerType As String '1
Private m_TransducerNumber As String '10
Private m_SerialNumber As String '50
Private m_IsSurveyed As Boolean
Private m_Timing As String '2
Private m_ActionDate As Date 'date
Private m_ActionTime As Date 'time

'transducer distances

'recorder/observer/downloader

Private m_ContactID As Long

'---------------------
' Events
'---------------------
Public Event InvalidTransducerType(value As String)
Public Event InvalidTransducerNumber(value As String)
Public Event InvalidSerialNumber(value As String)
Public Event InvalidTransducerTiming(value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let EventID(value As Long)
    m_EventID = value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let TransducerType(value As String)
    Dim aryTypes() As String
    aryTypes = Split(TRANSDUCER_TYPES, ",")
    
    If IsInArray(m_TransducerType, aryTypes) Then
        m_TransducerType = value
    Else
        RaiseEvent InvalidTransducerType(value)
    End If
End Property

Public Property Get TransducerType() As String
    TransducerType = m_TransducerType
End Property

Public Property Let TransducerNumber(value As String)
    If Len(Trim(value)) < 11 Then
        m_TransducerNumber = value
    Else
        RaiseEvent InvalidTransducerNumber(value)
    End If
End Property

Public Property Get TransducerNumber() As String
    TransducerNumber = m_TransducerNumber
End Property

Public Property Let SerialNumber(value As String)
    m_SerialNumber = value
End Property

Public Property Get SerialNumber() As String
    SerialNumber = m_SerialNumber
End Property

Public Property Let IsSurveyed(value As Boolean)
    m_IsSurveyed = value
End Property

Public Property Get IsSurveyed() As Boolean
    IsSurveyed = m_IsSurveyed
End Property

Public Property Let Timing(value As String)
    Dim aryTiming() As String
    aryTiming = Split(TRANSDUCER_TIMING, ",")
    If IsInArray(value, aryTiming) Then
        m_Timing = value
    Else
        RaiseEvent InvalidTransducerTiming(value)
    End If
End Property

Public Property Get Timing() As String
    Timing = m_Timing
End Property

Public Property Let ActionDate(value As Date)
    m_ActionDate = Format(value, "mm/dd/yyyy")
End Property

Public Property Get ActionDate() As Date
    ActionDate = m_ActionDate
End Property

Public Property Let ActionTime(value As Date)
    m_ActionTime = Format(value, "hh:mm:ss")
End Property

Public Property Get ActionTime() As Date
    ActionTime = m_ActionTime
End Property

Public Property Let ContactID(value As Long)
    m_ContactID = value
End Property

Public Property Get ContactID() As Long
    ContactID = m_ContactID
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Transducer])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Transducer])"
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
    
'    Dim strSQL As String, params As String
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'
'    Set db = CurrentDb
'
'    'record Transducers must have:
''    strSQL = "INSERT INTO Transducer(Event_ID, TransducerType, TransducerNumber, " _
''                & "SerialNumber, IsSurveyed, Timing, ActionDate, ActionTime) VALUES " _
''                & "(" & Me.EventID & ",'" & Me.TransducerType & "','" _
''                & Me.TransducerNumber & "','" & Me.SerialNumber & "'," _
''                & Me.IsSurveyed & ",'" & Me.Timing & "',#" _
''                & CDate(Me.ActionDate) & "#,#" & Format(Me.ActionTime, "hh:mm:ss") & "#);"
'
'    params = "EventID" & PARAM_SEPARATOR & Me.EventID & _
'            "|TransducerType" & PARAM_SEPARATOR & Me.TransducerType & _
'            "|TransducerNumber" & PARAM_SEPARATOR & Me.TransducerNumber & _
'            "|SerialNumber" & PARAM_SEPARATOR & Me.SerialNumber & _
'            "|IsSurveyed" & PARAM_SEPARATOR & Me.IsSurveyed & _
'            "|Timing" & PARAM_SEPARATOR & Me.Timing & _
'            "|ActionDate" & PARAM_SEPARATOR & CDate(Me.ActionDate) & _
'            "|ActionTime" & PARAM_SEPARATOR & Format(Me.ActionTime, "hh::mm::ss")
'
'    strSQL = GetTemplate("i_transducer", params)
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)


    Dim Template As String
    
    Template = "i_transducer"
    
    Dim Params(0 To 10) As Variant

    With Me
        Params(0) = "Transducer"
        Params(1) = .EventID
        Params(2) = .TransducerType
        Params(3) = .TransducerNumber
        Params(4) = .SerialNumber
        Params(5) = .IsSurveyed
        Params(6) = .Timing
        Params(7) = .ActionDate
        Params(8) = .ActionTime
    
        If IsUpdate Then
            Template = "u_transducer"
            Params(9) = .ID
        End If

        .ID = SetRecord(Template, Params)
    End With
    
'    'add a record for created by
'    Dim act As New RecordAction
'
'    With act
'        .RefAction = "R"
'        .ContactID = TempVars("UserID")
'        .RefID = Me.ID
'        .RefTable = "VegTransect"
'        .SaveToDb
'    End With


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Transducer])"
    End Select
    Resume Exit_Handler
End Sub