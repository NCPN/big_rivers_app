Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegWalk
' Level:        Framework class
' Version:      1.01
'
' Description:  Veg walk object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 4/19/2016
' References:   -
' Revisions:    BLC - 4/19/2016 - 1.00 - initial version
'               BLC - 8/8/2016  - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_EventID As Long
Private m_CollectionPlaceID As Long
Private m_CollectionType As String
Private m_StartDate As Date
Private m_CreateDate As Date
Private m_CreatedByID As Long
Private m_LastModified As Date
Private m_LastModifiedByID As Long

'---------------------
' Events
'---------------------
Public Event InvalidEventID(value As Long)
Public Event InvalidCollectionPlaceID(value As Long)
Public Event InvalidCollectionType(value As String)
Public Event InvalidStartDate(value As Date)
Public Event InvalidDate(value As Date)
Public Event InvalidContactID(value As Long)

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

Public Property Let CollectionPlaceID(value As Long)
    m_CollectionPlaceID = value
End Property

Public Property Get CollectionPlaceID() As Long
    CollectionPlaceID = m_CollectionPlaceID
End Property

Public Property Let CollectionType(value As String)
    Dim aryTypes() As String
    aryTypes = Split(COLLECTION_TYPES, ",")
    'check for valid collection type
    If IsInArray(value, aryTypes) Then
        m_CollectionType = value
    Else
        RaiseEvent InvalidCollectionType(value)
    End If
End Property

Public Property Get CollectionType() As String
    CollectionType = m_CollectionType
End Property

Public Property Let StartDate(value As Date)
    m_StartDate = value
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let CreateDate(value As Date)
    m_CreateDate = value
End Property

Public Property Get CreateDate() As Date
    CreateDate = m_CreateDate
End Property

Public Property Let CreatedByID(value As Long)
    m_CreatedByID = value
End Property

Public Property Get CreatedByID() As Long
    CreatedByID = m_CreatedByID
End Property

Public Property Let LastModified(value As Date)
    m_LastModified = value
End Property

Public Property Get LastModified() As Date
    LastModified = m_LastModified
End Property

Public Property Let LastModifiedByID(value As Long)
    m_LastModifiedByID = value
End Property

Public Property Get LastModifiedByID() As Long
    LastModifiedByID = m_LastModifiedByID
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
'   BLC - 4/19/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_VegWalk])"
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
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_VegWalk])"
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
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
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
'    'record actions must have:
'    strSQL = "INSERT INTO VegWalk(Event_ID, CollectionPlace_ID, " _
'                & "CollectionType, WalkStartDate, " _
'                & "CreateDate, CreatedBy_ID, LastModified, LastModifiedBy_ID) VALUES " _
'                & "(" & Me.EventID & "," & Me.CollectionPlaceID & ",'" _
'                & Me.CollectionType & "',#" & Me.StartDate & "#,#" _
'                & Now() & "#," & Me.CreatedByID & ",#" _
'                & Now() & "#," & Me.LastModifiedByID & ");"
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim Template As String
    
    Template = "i_vegwalk"
    
    Dim Params(0 To 10) As Variant
    
    With Me
        Params(0) = "VegWalk"
        Params(1) = .EventID
        Params(2) = .CollectionPlaceID
        Params(3) = .CollectionType
        Params(4) = .StartDate
'        params(5) = .CreateDate
'        params(6) = .CreatedByID
'        params(7) = .LastModified
'        params(8) = .LastModifiedByID
        
        If IsUpdate Then
            Template = "u_vegwalk"
            Params(9) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_VegWalk])"
    End Select
    Resume Exit_Handler
End Sub