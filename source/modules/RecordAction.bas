Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        RecordAction
' Level:        Framework class
' Version:      1.01
'
' Description:  Record action object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 7/26/2016 - 1.01 - revised Action to RefAction to avoid conflict (Jet reserved word)
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_RefAction As String
Private m_RefTable As String
Private m_RefID As Long
Private m_ContactID As Long
Private m_ActionType As String
Private m_ActionDate As Date

'---------------------
' Events
'---------------------
Public Event InvalidAction(Value As String)
Public Event InvalidRefTable(Value As String)
Public Event InvalidRefID(Value As Long)
Public Event InvalidContactID(Value As Long)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let RefTable(Value As String)
    If ValidateString(Value, "alphadashunderscore") Then
        m_RefTable = Value
    Else
        RaiseEvent InvalidRefTable(Value)
    End If
End Property

Public Property Get RefTable() As String
    RefTable = m_RefTable
End Property

Public Property Let RefID(Value As Long)
    m_RefID = Value
End Property

Public Property Get RefID() As Long
    RefID = m_RefID
End Property

Public Property Let ContactID(Value As Long)
    m_ContactID = Value
End Property

Public Property Get ContactID() As Long
    ContactID = m_ContactID
End Property

'Action type is verbose for action
Public Property Let ActionType(Value As String)
    Select Case Value
        Case "Observe"
            Me.RefAction = "O"
        Case "Record"
            Me.RefAction = "R"
        Case "DataEntry"
            Me.RefAction = "DE"
        Case "Download"
            Me.RefAction = "D"
        Case "Upload"
            Me.RefAction = "U"
        Case "Change"
            Me.RefAction = "E"
        Case "Verify"
            Me.RefAction = "V"
        Case "Certify"
            Me.RefAction = "C"
    End Select

    m_ActionType = Value
End Property

Public Property Get ActionType() As String
    ActionType = m_ActionType
End Property

Public Property Let RefAction(Value As String)
    Dim aryActions() As String
    aryActions = Split(RECORD_ACTIONS, ",")
    
    If IsInArray(m_RefAction, aryActions) Then
        m_RefAction = Value
    Else
        RaiseEvent InvalidAction(Value)
    End If
End Property

Public Property Get RefAction() As String
    RefAction = m_RefAction
End Property

Public Property Let ActionDate(Value As Date)
    m_ActionDate = Value
End Property

Public Property Get ActionDate() As Date
    ActionDate = m_ActionDate
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Action])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Action])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save data to database
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
    
'    Dim strSQL As String
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'
'    Set db = CurrentDb
'
'    'record actions must have:
''    strSQL = "INSERT INTO RecordAction(ReferenceType, Reference_ID, Contact_ID, Action, ActionDate) VALUES " _
''                & "('" & Me.RefTable & "'," & Me.RefID & "," _
''                & Me.ID & ",'" & Me.action & "', Now() );"
'
'    strSQL = GetTemplate("i_action_record", _
'                         "RefType" & PARAM_SEPARATOR & Me.RefTable _
'                        & "|RefID" & PARAM_SEPARATOR & Me.RefID _
'                        & "|ID" & PARAM_SEPARATOR & Me.ID _
'                        & "|action" & PARAM_SEPARATOR & Me.action _
'                        & "|actiondate" & PARAM_SEPARATOR & "Now()")
'
''********************
''  FIX: Me.RefTable & actiondate values
''********************
'
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim params(0 To 4) As Variant

    params(0) = Me.RefTable
    params(1) = Me.RefID
    params(2) = Me.ID
    params(3) = Me.RefAction
    params(4) = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
    
    Me.ID = SetRecord("i_record_action", params)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Action])"
    End Select
    Resume Exit_Handler
End Sub