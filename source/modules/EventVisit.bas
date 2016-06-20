Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        EventVisit
' Level:        Framework class
' Version:      1.00
'
' Description:  Event object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 4/4/2016   - 1.01 - renamed to "EventVisit" to avoid collision w/ "Event" vba term
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_StartDate As Date
Private m_SiteID As Integer
Private m_LocationID As Integer
Private m_ProtocolID As Integer

'---------------------
' Events
'---------------------
Public Event InvalidEventID()
Public Event InvalidSiteID()
Public Event Modified()
Public Event SavedToDb()
Public Event Removed()

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let SiteID(value As Integer)
    m_SiteID = value
End Property

Public Property Get SiteID() As Integer
    SiteID = m_SiteID
End Property

Public Property Let LocationID(value As Integer)
    m_LocationID = value
End Property

Public Property Get LocationID() As Integer
    LocationID = m_LocationID
End Property

Public Property Let ProtocolID(value As Integer)
    m_ProtocolID = value
End Property

Public Property Get ProtocolID() As Integer
    ProtocolID = m_ProtocolID
End Property

Public Property Let StartDate(value As Date)
    m_StartDate = value
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Event])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Event])"
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
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    'events must have: start date, site ID, location ID, protocol ID
'    strSQL = "INSERT INTO Event(Protocol_ID, Site_ID, Location_ID, StartDate) VALUES " _
'                & "(" & Me.ProtocolID & "," & Me.SiteID & "," _
'                & Me.LocationID & "," & Me.StartDate & ");"

    strSQL = GetTemplate("i_event_record", _
                "ProtocolID" & PARAM_SEPARATOR & Me.ProtocolID & "|" _
                & "SiteID" & PARAM_SEPARATOR & Me.SiteID & "|" _
                & "LocationID" & PARAM_SEPARATOR & Me.LocationID & "|" _
                & "StartDate" & PARAM_SEPARATOR & Format(Me.StartDate, "YYYY-mm-dd"))
    
    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    'add a record for created by
    Dim act As New action
    With act
        act.action = "R"
        'act.ContactID =
        act.RefID = Me.ID
        act.RefTable = "Event"
        act.SaveToDb
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Event])"
    End Select
    Resume Exit_Handler
End Sub