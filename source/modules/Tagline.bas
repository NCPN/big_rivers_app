Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Tagline
' Level:        Framework class
' Version:      1.01
'
' Description:  Record Tagline object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 6/1/2016  - 1.01 - updated to use GetTemplate() in SaveToDb()
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_LineDistSource As String
Private m_LineDistSourceID As Long
Private m_LineDistType As String
Private m_LineDistance As Integer
Private m_HeightType As String
Private m_Height As Integer

'---------------------
' Events
'---------------------
Public Event InvalidLineDistSource(Value As String)
Public Event InvalidLineDistType(Value As String)
Public Event InvalidLineDistance(Value As Integer) 'in m
Public Event InvalidHeightType(Value As String)
Public Event InvalidHeight(Value As Integer)    'in cm

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let LineDistSource(Value As String)
    Dim arySources() As String
    arySources = Split(LINE_DIST_SOURCES, ",")
    If IsInArray(Value, arySources) Then
            m_LineDistSource = Value
    Else
        RaiseEvent InvalidLineDistSource(Value)
    End If
End Property

Public Property Get LineDistSource() As String
    LineDistSource = m_LineDistSource
End Property

Public Property Let LineDistSourceID(Value As Long)
    m_LineDistSourceID = Value
End Property

Public Property Get LineDistSourceID() As Long
    LineDistSourceID = m_LineDistSourceID
End Property

Public Property Let LineDistType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(LINE_DIST_TYPES, ",")
    If IsInArray(Value, aryTypes) Then
            m_LineDistType = Value
    Else
        RaiseEvent InvalidLineDistType(Value)
    End If
End Property

Public Property Get LineDistType() As String
    LineDistType = m_LineDistType
End Property

Public Property Let LineDistance(Value As Integer)
    m_LineDistance = Value
End Property

Public Property Get LineDistance() As Integer
    LineDistance = m_LineDistance
End Property

Public Property Let HeightType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(HEIGHT_TYPES, ",")
    If IsInArray(Value, aryTypes) Then
        m_HeightType = Value
    Else
        RaiseEvent InvalidHeightType(Value)
    End If
End Property

Public Property Get HeightType() As String
    HeightType = m_HeightType
End Property

Public Property Let Height(Value As Integer)
    m_Height = Value
End Property

Public Property Get Height() As Integer
    Height = m_Height
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Tagline])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Tagline])"
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
'   BLC, 6/1/2016 - updated to use GetTemplate()
'---------------------------------------------------------------------------------------
Public Sub SaveToDb()
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As dao.Database
    Dim rs As dao.Recordset
    
    Set db = CurrentDb
    
    If Me.ID > 0 Then
        'update tagline:
        strSQL = GetTemplate("u_tagline_record", _
                    "LineDistSource:" & Me.LineDistSource _
                    & "|LineDistSourceID:" & Me.LineDistSourceID _
                    & "|LineDistType:" & Me.LineDistType _
                    & "|LineDistance:" & Me.LineDistance _
                    & "|HeightType:" & Me.HeightType _
                    & "|Height:" & Me.Height _
                    & "|ID:" & Me.ID)
    Else
        'insert tagline
        strSQL = GetTemplate("i_tagline_record", _
                    "LineDistSource:" & Me.LineDistSource _
                    & "|LineDistSourceID:" & Me.LineDistSourceID _
                    & "|LineDistType:" & Me.LineDistType _
                    & "|LineDistance:" & Me.LineDistance _
                    & "|HeightType:" & Me.HeightType _
                    & "|Height:" & Me.Height)
    End If
    
    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Tagline])"
    End Select
    Resume Exit_Handler
End Sub