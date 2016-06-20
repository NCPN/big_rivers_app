Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Site
' Level:        Framework class
' Version:      1.00
'
' Description:  Site object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer 'siteID
Private m_ParkID As Integer
Private m_RiverID As Integer
Private m_Code As String
Private m_Name As String
Private m_Description As String
Private m_Directions As String
Private m_IsActiveForProtocol As Boolean
Private m_Park As String
Private m_LocationID As Integer
Private m_ObserverID As Integer
Private m_RecorderID As Integer
Private m_Observer As String
Private m_Recorder As String
Private m_CommentID As Integer
Private m_Comment As String

'---------------------
' Events
'---------------------
Public Event InvalidPark(value)
Public Event InvalidSiteName(value)
Public Event InvalidSiteCode(value)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let ParkID(value As Integer)
    m_ParkID = value
End Property

Public Property Get ParkID() As Integer
    ParkID = m_ParkID
End Property

Public Property Let RiverID(value As Integer)
    m_RiverID = value
End Property

Public Property Get RiverID() As Integer
    RiverID = m_RiverID
End Property

Public Property Let Code(value As String)
    If Len(Trim(value)) = 2 Then
        m_Code = value
    Else
        RaiseEvent InvalidSiteCode(value)
    End If
End Property

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Name(value As String)
    m_Name = value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Description(value As String)
    m_Description = value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Directions(value As String)
    m_Directions = value
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let IsActiveForProtocol(value As Integer)
    m_IsActiveForProtocol = value
End Property

Public Property Get IsActiveForProtocol() As Integer
    IsActiveForProtocol = m_IsActiveForProtocol
End Property

Public Property Let Park(value As String)
    Dim aryParks() As String
    aryParks = Split(PARKS, ",")
    If IsInArray(value, aryParks) Then
        m_Park = value
    Else
        RaiseEvent InvalidPark(value)
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let LocationID(value As Integer)
    m_LocationID = value
End Property

Public Property Get LocationID() As Integer
    LocationID = m_LocationID
End Property

Public Property Let ObserverID(value As Integer)
    m_ObserverID = value
End Property

Public Property Get ObserverID() As Integer
    ObserverID = m_ObserverID
End Property

Public Property Let Observer(value As String)
    m_Observer = value
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let RecorderID(value As Integer)
    m_RecorderID = value
End Property

Public Property Get RecorderID() As Integer
    RecorderID = m_RecorderID
End Property

Public Property Let Recorder(value As String)
    m_Recorder = value
End Property

Public Property Get Recorder() As String
    Recorder = m_Recorder
End Property

'---------------------
'change to comment object instead??
'---------------------
Public Property Let CommentID(value As Integer)
    m_CommentID = value
End Property

Public Property Get CommentID() As Integer
    CommentID = m_CommentID
End Property

Public Property Let Comment(value As String)
    m_Comment = value
End Property

Public Property Get Comment() As String
    Comment = m_Comment
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Site])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Site])"
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
    strSQL = "INSERT INTO Site(Park_ID, River_ID, SiteCode, SiteName, " _
                & "SiteDirections, SiteDescription, " _
                & "IsActiveForProtocol) VALUES " _
                & "(" & Me.ParkID & "," & Me.RiverID & ",'" _
                & Me.Code & "','" & Me.Name & "','" _
                & Me.Directions & "','" & Me.Description & "'," _
                & Me.IsActiveForProtocol & ");"

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    'handle record actions
    Dim act As New action
    With act
    
    'Recorder
        .action = "R"
        .ContactID = Me.RecorderID
        .RefID = Me.ID
        .RefTable = "Site"
        .SaveToDb
        
    'Observer
        .action = "O"
        .ContactID = Me.ObserverID
        .RefID = Me.ID
        .RefTable = "Site"
        .SaveToDb
        
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Site])"
    End Select
    Resume Exit_Handler
End Sub