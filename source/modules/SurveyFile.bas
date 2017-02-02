Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        SurveyFile
' Level:        Framework class
' Version:      1.02
'
' Description:  SurveyFile object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, January 26, 2017
' References:   -
' Revisions:    BLC - 1/26/2017 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_Name As String
Private m_Path As String
Private m_FullPath As String
Private m_SurveyType As String
Private m_SurveySource As String
Private m_StartRecord As Long
Private m_EndRecord As Long
Private m_PointCount As Integer
Private m_TranslationPtID As Long
Private m_RotationPtID As Long
Private m_TranslationErrorID As Long
Private m_RotationErrorID As Long
Private m_BaseErrorID As Long
Private m_SurveyErrorID As Long

'---------------------
' Events
'---------------------
Public Event InvalidName(Value)
Public Event InvalidPath(Value)
Public Event InvalidSurveyType(Value)
Public Event InvalidSurveySource(Value)
Public Event InvalidPointID(Value)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Path(Value As String)
    m_Path = Value
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let FullPath(Value As String)
    m_FullPath = Value
End Property

Public Property Get FullPath() As String
    FullPath = m_FullPath
End Property

Public Property Let SurveyType(Value As String)
    m_SurveyType = Value
End Property

Public Property Get SurveyType() As String
    SurveyType = m_SurveyType
End Property

Public Property Let SurveySource(Value As String)
    m_SurveySource = Value
End Property

Public Property Get SurveySource() As String
    SurveySource = m_SurveySource
End Property

Public Property Let StartRecord(Value As Long)
    m_StartRecord = Value
End Property

Public Property Get StartRecord() As Long
    StartRecord = m_StartRecord
End Property

Public Property Let EndRecord(Value As Long)
    m_EndRecord = Value
End Property

Public Property Get EndRecord() As Long
    EndRecord = m_EndRecord
End Property

Public Property Let PointCount(Value As Integer)
    m_PointCount = Value
End Property

Public Property Get PointCount() As Integer
    PointCount = m_PointCount
End Property

'---------------------
' Points / Pt Errors
'---------------------
Public Property Let TranslationPtID(Value As Long)
    m_TranslationPtID = Value
End Property

Public Property Get TranslationPtID() As Long
    TranslationPtID = m_TranslationPtID
End Property

Public Property Let RotationPtID(Value As Long)
    m_RotationPtID = Value
End Property

Public Property Get RotationPtID() As Long
    RotationPtID = m_RotationPtID
End Property

Public Property Let TranslationErrorID(Value As Long)
    m_TranslationErrorID = Value
End Property

Public Property Get TranslationErrorID() As Long
    TranslationErrorID = m_TranslationErrorID
End Property

Public Property Let RotationErrorID(Value As Long)
    m_RotationErrorID = Value
End Property

Public Property Get RotationErrorID() As Long
    RotationErrorID = m_RotationErrorID
End Property

Public Property Let BaseErrorID(Value As Long)
    m_BaseErrorID = Value
End Property

Public Property Get BaseErrorID() As Long
    BaseErrorID = m_BaseErrorID
End Property

Public Property Let SurveyErrorID(Value As Long)
    m_SurveyErrorID = Value
End Property

Public Property Get SurveyErrorID() As Long
    SurveyErrorID = m_SurveyErrorID
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_SurveyFile])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_SurveyFile])"
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
'   BLC, 1/26/2017 - updated for survey file class
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler

    Dim Template As String
    
    Template = "i_surveyfile"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "SurveyFile"
        Params(1) = .Name
        Params(2) = .Path
        Params(3) = .SurveyType
        Params(4) = .SurveySource
        
        If IsUpdate Then
            Template = "u_surveyfile"
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
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_SurveyFile])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SetPointData
' Description:  Set survey point data for the survey file (points & errors)
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 1/26/2017 - for NCPN tools
' Revisions:
'   BLC, 1/26/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SetPointData(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler

    Dim Template As String
    
    Template = "u_surveyfile_ptdata"
    
    Dim Params(0 To 12) As Variant
    
    With Me
        Params(0) = "SurveyFile"
        Params(1) = .Name
        Params(2) = .Path
        Params(3) = .SurveyType
        Params(4) = .SurveySource
        Params(5) = .TranslationPtID
        Params(6) = .RotationPtID
        Params(7) = .TranslationErrorID
        Params(8) = .RotationErrorID
        Params(9) = .BaseErrorID
        Params(10) = .SurveyErrorID
        
        'always an update -> use ID
        Params(11) = .ID
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SetPointData[cls_SurveyFile])"
    End Select
    Resume Exit_Handler
End Sub