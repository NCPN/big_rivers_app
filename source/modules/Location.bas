Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Location
' Level:        Framework class
' Version:      1.00
'
' Description:  Location object related properties, Locations, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long 'Integer
Private m_LocationType As String
Private m_CollectionSourceName As String
Private m_LocationName As String
Private m_HeadtoOrientDistance As Integer
Private m_HeadtoOrientBearing As Integer
Private m_LocationNotes As String
Private m_LastModified As Date
Private m_LastModifiedByID As Long 'Integer
Private m_CreateDate As Date
Private m_CreatedByID As Long 'Integer

'---------------------
' Events
'---------------------
Public Event InvalidLocationType(Value)
Public Event InvalidLocationName(Value)
Public Event InvalidBearing(Value)
Public Event InvalidSourceName(Value)
'Public Event Modified()
'Public Event SavedToDb()
'Public Event Deleted()

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let LocationType(Value As String)
    'get valid location types (mod_App_Settings)
    Dim aryLocTypes() As String
    aryLocTypes = Split(LOCATION_TYPES, ",")

    If Len(Trim(Value)) = 1 And IsInArray(Value, aryLocTypes) Then
        m_LocationType = Value
    Else
        RaiseEvent InvalidLocationType(Value)
    End If

End Property

Public Property Get LocationType() As String
    LocationType = m_LocationType
End Property

Public Property Let CollectionSourceName(Value As String)
    'Collection feature ID (A, B, C, ...) or Transect number (1-8)
    'limit = 25
    If Len(Trim(Value)) < 26 Then
        m_CollectionSourceName = Value
    Else
        RaiseEvent InvalidSourceName(Value)
    End If
End Property

Public Property Get CollectionSourceName() As String
    CollectionSourceName = m_CollectionSourceName
End Property

Public Property Let LocationName(Value As String)
    'limit = 100
    If Len(Trim(Value)) < 101 Then
        m_LocationName = Value
    Else
        RaiseEvent InvalidLocationName(Value)
    End If
End Property

Public Property Get LocationName() As String
    LocationName = m_LocationName
End Property

Public Property Let HeadtoOrientDistance(Value As Integer)
    m_HeadtoOrientDistance = Value
End Property

Public Property Get HeadtoOrientDistance() As Integer
    HeadtoOrientDistance = m_HeadtoOrientDistance
End Property

Public Property Let HeadtoOrientBearing(Value As Integer)
    If IsBetween(Value, 0, 360, True) Then
        m_HeadtoOrientBearing = Value
    End If
End Property

Public Property Get HeadtoOrientBearing() As Integer
    HeadtoOrientBearing = m_HeadtoOrientBearing
End Property

Public Property Let CreatedByID(Value As Integer)
    m_CreatedByID = Value
End Property

Public Property Get CreatedByID() As Integer
    CreatedByID = m_CreatedByID
End Property

Public Property Let CreateDate(Value As Date)
    m_CreateDate = Value
End Property

Public Property Get CreateDate() As Date
    CreateDate = m_CreateDate
End Property

Public Property Let LastModifiedByID(Value As Integer)
    m_LastModifiedByID = Value
End Property

Public Property Get LastModifiedByID() As Integer
    LastModifiedByID = m_LastModifiedByID
End Property

Public Property Let LastModified(Value As Date)
    m_LastModified = Value
End Property

Public Property Get LastModified() As Date
    LastModified = m_LastModified
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Location])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Location])"
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
'    'events must have: collection source
'    strSQL = "INSERT INTO Location(CollectionSourceName, LocationType, LocationName, " _
'                & "HeadtoOrientDistance_m, HeadtoOrientBearing, CreateDate, " _
'                & "CreatedBy_ID, LastModified, LastModifiedBy_ID) VALUES " _
'                & "('" & Me.CollectionSourceName & "','" & Me.LocationType & "','" _
'                & Me.LocationName & "'," & Me.HeadtoOrientDistance _
'                & "," & Me.HeadtoOrientBearing & ", Now()," _
'                & Me.CreatedByID & ", Now()," _
'                & Me.CreatedByID & ");"
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim template As String
    
    template = "i_location"
    
    Dim params(0 To 10) As Variant

    With Me
        params(0) = .CollectionSourceName
        params(1) = .LocationType
        params(2) = .LocationName
        params(3) = .HeadtoOrientDistance
        params(4) = .HeadtoOrientBearing
        'params 5-8 are create, last modified
        
        If IsUpdate Then
            template = "u_location"
            params(9) = .ID
        End If
        
        .ID = SetRecord(template, params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Location])"
    End Select
    Resume Exit_Handler

End Sub