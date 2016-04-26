Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Action
' Level:        Framework class
' Version:      1.00
'
' Description:  Record action object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Action As String
Private m_RefTable As String
Private m_RefID As Long
Private m_ContactID As Long
Private m_ActionType As String
Private m_ActionDate As Date

'---------------------
' Events
'---------------------
Public Event InvalidAction(value As String)
Public Event InvalidRefTable(value As String)
Public Event InvalidRefID(value As Long)
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

Public Property Let RefTable(value As String)
    If ValidateString(value, "alphadashunderscore") Then
        m_RefTable = value
    Else
        RaiseEvent InvalidRefTable(value)
    End If
End Property

Public Property Get RefTable() As String
    RefTable = m_RefTable
End Property

Public Property Let RefID(value As Long)
    m_RefID = value
End Property

Public Property Get RefID() As Long
    RefID = m_RefID
End Property

Public Property Let ContactID(value As Long)
    m_ContactID = value
End Property

Public Property Get ContactID() As Long
    ContactID = m_ContactID
End Property

'Action type is verbose for action
Public Property Let ActionType(value As String)
    Select Case value
        Case "Observe"
            Me.action = "O"
        Case "Record"
            Me.action = "R"
        Case "DataEntry"
            Me.action = "DE"
        Case "Download"
            Me.action = "D"
        Case "Upload"
            Me.action = "U"
        Case "Change"
            Me.action = "E"
        Case "Verify"
            Me.action = "V"
        Case "Certify"
            Me.action = "C"
    End Select

    m_ActionType = value
End Property

Public Property Get ActionType() As String
    ActionType = m_ActionType
End Property

Public Property Let action(value As String)
    Dim aryActions() As String
    aryActions = Split(RECORD_ACTIONS, ",")
    
    If IsInArray(m_Action, aryActions) Then
        m_Action = value
    Else
        RaiseEvent InvalidAction(value)
    End If
End Property

Public Property Get action() As String
    action = m_Action
End Property

Public Property Let ActionDate(value As Date)
    m_ActionDate = value
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
    
    'record actions must have:
    strSQL = "INSERT INTO RecordAction(ReferenceType, Reference_ID, Contact_ID, Action, ActionDate) VALUES " _
                & "('" & Me.RefTable & "'," & Me.RefID & "," _
                & Me.ID & ",'" & Me.action & "', Now() );"

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

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