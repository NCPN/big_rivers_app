Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        EventVisit
' Level:        Framework class
' Version:      1.03
'
' Description:  Event object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 4/4/2016   - 1.01 - renamed to "EventVisit" to avoid collision w/ "Event" vba term
'               BLC - 8/8/2016   - 1.02 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               BLC - 9/1/2016   - 1.03 - SaveToDb() code cleanup
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
'Public Event Modified()
'Public Event SavedToDb()
'Public Event Removed()

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let SiteID(Value As Integer)
    m_SiteID = Value
End Property

Public Property Get SiteID() As Integer
    SiteID = m_SiteID
End Property

Public Property Let LocationID(Value As Integer)
    m_LocationID = Value
End Property

Public Property Get LocationID() As Integer
    LocationID = m_LocationID
End Property

Public Property Let ProtocolID(Value As Integer)
    m_ProtocolID = Value
End Property

Public Property Get ProtocolID() As Integer
    ProtocolID = m_ProtocolID
End Property

Public Property Let StartDate(Value As Date)
    m_StartDate = Value
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
' Parameters:   IsUpdate - indicates if data is an update vs. an insert (boolean, optional)
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
'   BLC, 7/27/2016 - added update parameter to identify if this is an update vs. an insert
'   BLC, 9/1/2016  - commented code cleanup
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_event"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "Event"
        Params(1) = .SiteID
        Params(2) = .LocationID
        Params(3) = .ProtocolID
        Params(4) = CDate(Format(.StartDate, "YYYY-mm-dd"))
        
        If IsUpdate Then
            Template = "u_event"
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
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Event])"
    End Select
    Resume Exit_Handler
End Sub