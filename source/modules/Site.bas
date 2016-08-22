Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Site
' Level:        Framework class
' Version:      1.02
'
' Description:  Site object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:
'   Olivier Jacot-Descombes, January 12, 2012
'   http://stackoverflow.com/questions/8827447/why-is-yes-a-value-of-1-in-ms-access-database
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 6/28/2016 -  1.01 - revised boolean to byte to avoid Access use of -1 for true
'                                         & force IsActiveForProtocol flag to be 1 or 0
'                                         see Olivier Jacot-Descombes notes on why Access uses -1
'                                         but preference is to use 1 & 0 to facilitate clarity
'                                         within SQL
'               BLC - 8/8/2016   - 1.02 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_ParkID As Long
Private m_RiverID As Long
Private m_Code As String
Private m_Name As String
Private m_Description As String
Private m_Directions As String
Private m_IsActiveForProtocol As Byte
Private m_Park As String
Private m_River As String
Private m_LocationID As Long
Private m_ObserverID As Long
Private m_RecorderID As Long
Private m_Observer As String
Private m_Recorder As String
Private m_CommentID As Long
Private m_Comment As String

'---------------------
' Events
'---------------------
Public Event InvalidPark(Value)
Public Event InvalidRiver(Value)
Public Event InvalidSiteName(Value)
Public Event InvalidSiteCode(Value)

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

Public Property Let RiverID(Value As Long)
    m_RiverID = Value
End Property

Public Property Get RiverID() As Long
    RiverID = m_RiverID
End Property

Public Property Let Code(Value As String)
    If Len(Trim(Value)) = 2 Then
        m_Code = Value
    Else
        RaiseEvent InvalidSiteCode(Value)
    End If
End Property

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Directions(Value As String)
    m_Directions = Value
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let IsActiveForProtocol(Value As Byte)
    m_IsActiveForProtocol = Value
End Property

Public Property Get IsActiveForProtocol() As Byte
    IsActiveForProtocol = m_IsActiveForProtocol
End Property

Public Property Let Park(Value As String)
    Dim aryParks() As String
    aryParks = Split(PARKS, ",")
    If IsInArray(Value, aryParks) Then
        m_Park = Value
        
        'set park id also
        ParkID = GetParkID(m_Park)
    Else
        RaiseEvent InvalidPark(Value)
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let River(Value As String)
    If Len(Value) > 2 Then
        m_River = Value
        
        'set River id also
        RiverID = GetRiverSegmentID(m_River)
    Else
        RaiseEvent InvalidRiver(Value)
    End If
End Property

Public Property Get River() As String
    River = m_River
End Property

Public Property Let LocationID(Value As Long)
    m_LocationID = Value
End Property

Public Property Get LocationID() As Long
    LocationID = m_LocationID
End Property

Public Property Let ObserverID(Value As Long)
    m_ObserverID = Value
End Property

Public Property Get ObserverID() As Long
    ObserverID = m_ObserverID
End Property

Public Property Let Observer(Value As String)
    m_Observer = Value
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let RecorderID(Value As Long)
    m_RecorderID = Value
End Property

Public Property Get RecorderID() As Long
    RecorderID = m_RecorderID
End Property

Public Property Let Recorder(Value As String)
    m_Recorder = Value
End Property

Public Property Get Recorder() As String
    Recorder = m_Recorder
End Property

'---------------------
'change to comment object instead??
'---------------------
Public Property Let CommentID(Value As Long)
    m_CommentID = Value
End Property

Public Property Get CommentID() As Long
    CommentID = m_CommentID
End Property

Public Property Let Comment(Value As String)
    m_Comment = Value
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
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'    Dim db As DAO.Database
'    Dim qdf As DAO.QueryDef
'    Dim rs As DAO.Recordset
'    Dim iCount As Integer
'
'    Set db = CurrentDb
'
'    'events must have: start date, site ID, location ID, protocol ID
''    strSQL = "INSERT INTO Site(Park_ID, River_ID, SiteCode, SiteName, " _
''                & "SiteDirections, SiteDescription, " _
''                & "IsActiveForProtocol) VALUES " _
''                & "(" & Me.ParkID & "," & Me.RiverID & ",'" _
''                & Me.Code & "','" & Me.Name & "','" _
''                & Me.Directions & "','" & Me.Description & "'," _
''                & Me.IsActiveForProtocol & ");"
'    With db
'        Set qdf = .QueryDefs("usys_temp_qdf")
'
'        With qdf
'            'check if record exists in site
'            .SQL = GetTemplate("s_count_tbl", _
'                    "field" & PARAM_SEPARATOR & "ID" & _
'                    "|tbl" & PARAM_SEPARATOR & "Site WHERE SiteCode = '" & Me.Code & _
'                    "' AND Park_ID = " & Me.ParkID & " AND River_ID = " & Me.RiverID)
'            Set rs = .OpenRecordset
'            If rs.Fields(0) > 0 Then iCount = rs.Fields(0)
'        End With
'
'        Set qdf = .QueryDefs("usys_temp_qdf")
'
'        With qdf
'            'update if site is in site, otherwise insert new record
'            If iCount > 0 Then
'                .SQL = GetTemplate("u_site")
'            Else
'                .SQL = GetTemplate("i_site_record")
'            End If
'
'            '-- required parameters --
'            .Parameters("parkid") = Me.ParkID
'            .Parameters("riverid") = Me.RiverID
'            .Parameters("code") = Me.Code
'            .Parameters("sitename") = Me.Name
'            .Parameters("flag") = Me.IsActiveForProtocol
'
'            '-- optional parameters --
'            If Not IsNull(Me.Directions) And Not Len(Me.Directions) = 0 Then _
'                .Parameters("dir") = Me.Directions
'            If Not IsNull(Me.Description) And Not Len(Me.Description) = 0 Then _
'                .Parameters("descr") = Me.Description
'
'            .Execute dbFailOnError
'
'            'cleanup
'            .Close
'        End With
'
'        'retrieve identity
'        Me.ID = .OpenRecordset("SELECT @@IDENTITY;")(0)
'
'    End With


    Dim template As String
    
    template = "i_site"
    
    Dim params(0 To 9) As Variant
    
    With Me
        params(0) = "Site"
        params(1) = .ParkID
        params(2) = .RiverID
        params(3) = .Code
        params(4) = .Name
        params(5) = .IsActiveForProtocol
        
        params(6) = .Directions
        params(7) = .Description
        
        If IsUpdate Then
            template = "u_site"
            params(8) = .ID
        End If
        
        .ID = SetRecord(template, params)
    End With


'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

'    'handle record actions
'    Dim act As New RecordAction
'    With act
'
'    'Recorder
'        .RefAction = "R"
'        .ContactID = Me.RecorderID
'        .RefID = Me.ID
'        .RefTable = "Site"
'        .SaveToDb
'
'    'Observer
'        .RefAction = "O"
'        .ContactID = Me.ObserverID
'        .RefID = Me.ID
'        .RefTable = "Site"
'        .SaveToDb
'
'    End With

    SetObserverRecorder Me, "Site"

Exit_Handler:
'    'cleanup
'    Set qdf = Nothing
'    Set rs = Nothing
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Site])"
    End Select
    Resume Exit_Handler
End Sub