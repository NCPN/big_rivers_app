Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegTransect
' Level:        Framework class
' Version:      1.01
'
' Description:  VegTransect object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/8/2016   - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_LocationID As Long
Private m_EventID As Long
Private m_TransectNumber As Integer
Private m_SampleDate As Date

Private m_Park As String
Private m_ObserverID As Integer
Private m_RecorderID As Integer
Private m_Observer As String
Private m_Recorder As String

'---------------------
' Events
'---------------------
Public Event InvalidTransectNumber(Value As Integer)


'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let LocationID(Value As Long)
    m_LocationID = Value
    'set the appropriate park value
'    Me.Park = GetParkCode(Value)
End Property

Public Property Get LocationID() As Long
    LocationID = m_LocationID
End Property

Public Property Let EventID(Value As Long)
    m_EventID = Value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let TransectNumber(Value As Integer)
    If IsNull(Me.Park) Then
        MsgBox "Park must be set before setting transect number.", vbCritical, "Missing Park"
        
    End If
    'validate park (BLCA & CANY only)
    Select Case Me.Park
        Case "BLCA", "CANY"
            'check value
            'validate transect #
            Dim aryTransectNum() As String
            aryTransectNum = Split(TRANSECT_NUMBERS, ",")
            If IsInArray(CStr(Value), aryTransectNum) Then
                m_TransectNumber = Value
            Else
                RaiseEvent InvalidTransectNumber(Value)
            End If
        Case "DINO"
            'invalid
            RaiseEvent InvalidTransectNumber(Value)
        Case Else
            'invalid
            RaiseEvent InvalidTransectNumber(Value)
    End Select
End Property

Public Property Get TransectNumber() As Integer
    TransectNumber = m_TransectNumber
End Property

Public Property Let SampleDate(Value As Date)
    m_SampleDate = Value
End Property

Public Property Get SampleDate() As Date
    SampleDate = m_SampleDate
End Property


Public Property Let Park(Value As String)
    m_Park = Value
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let ObserverID(Value As Integer)
    m_ObserverID = Value
End Property

Public Property Get ObserverID() As Integer
    ObserverID = m_ObserverID
End Property

Public Property Let Observer(Value As String)
    m_Observer = Value
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let RecorderID(Value As Integer)
    m_RecorderID = Value
End Property

Public Property Get RecorderID() As Integer
    RecorderID = m_RecorderID
End Property

Public Property Let Recorder(Value As String)
    m_Recorder = Value
End Property

Public Property Get Recorder() As String
    Recorder = m_Recorder
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_VegPlot])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_VegPlot])"
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
'   BLC, 9/8/2016 - code cleanup
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_vegtransect"
    
    Dim Params(0 To 6) As Variant

    With Me
        Params(0) = "VegTransect"
        Params(1) = .LocationID
        Params(2) = .EventID
        Params(3) = .TransectNumber
        Params(4) = .SampleDate
        
        If IsUpdate Then
            Template = "u_vegtransect"
            Params(5) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
    SetObserverRecorder Me, "VegTransect"

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub