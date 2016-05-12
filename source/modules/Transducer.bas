Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Transducer
' Level:        Framework class
' Version:      1.00
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
Public Event InvalidTransducerType(Value As String)
Public Event InvalidTransducerNumber(Value As String)
Public Event InvalidSerialNumber(Value As String)
Public Event InvalidTransducerTiming(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let EventID(Value As Long)
    m_EventID = Value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let TransducerType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(TRANSDUCER_TYPES, ",")
    
    If IsInArray(m_TransducerType, aryTypes) Then
        m_TransducerType = Value
    Else
        RaiseEvent InvalidTransducerType(Value)
    End If
End Property

Public Property Get TransducerType() As String
    TransducerType = m_TransducerType
End Property

Public Property Let TransducerNumber(Value As String)
    If Len(Trim(Value)) < 11 Then
        m_TransducerNumber = Value
    Else
        RaiseEvent InvalidTransducerNumber(Value)
    End If
End Property

Public Property Get TransducerNumber() As String
    TransducerNumber = m_TransducerNumber
End Property

Public Property Let SerialNumber(Value As String)
    m_SerialNumber = Value
End Property

Public Property Get SerialNumber() As String
    SerialNumber = m_SerialNumber
End Property

Public Property Let IsSurveyed(Value As Boolean)
    m_IsSurveyed = Value
End Property

Public Property Get IsSurveyed() As Boolean
    IsSurveyed = m_IsSurveyed
End Property

Public Property Let Timing(Value As String)
    Dim aryTiming() As String
    aryTiming = Split(TRANSDUCER_TIMING, ",")
    If IsInArray(Value, aryTiming) Then
        m_Timing = Value
    Else
        RaiseEvent InvalidTransducerTiming(Value)
    End If
End Property

Public Property Get Timing() As String
    Timing = m_Timing
End Property

Public Property Let ActionDate(Value As Date)
    m_ActionDate = Format(Value, "mm/dd/yyyy")
End Property

Public Property Get ActionDate() As Date
    ActionDate = m_ActionDate
End Property

Public Property Let ActionTime(Value As Date)
    m_ActionTime = Format(Value, "hh:mm:ss")
End Property

Public Property Get ActionTime() As Date
    ActionTime = m_ActionTime
End Property

Public Property Let ContactID(Value As Long)
    m_ContactID = Value
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb()
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As dao.Database
    Dim rs As dao.Recordset
    
    Set db = CurrentDb
    
    'record Transducers must have:
    strSQL = "INSERT INTO Transducer(Event_ID, TransducerType, TransducerNumber, " _
                & "SerialNumber, IsSurveyed, Timing, ActionDate, ActionTime) VALUES " _
                & "(" & Me.EventID & ",'" & Me.TransducerType & "','" _
                & Me.TransducerNumber & "','" & Me.SerialNumber & "'," _
                & Me.IsSurveyed & ",'" & Me.Timing & "',#" _
                & CDate(Me.ActionDate) & "#,#" & Format(Me.ActionTime, "hh:mm:ss") & "#);"

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

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