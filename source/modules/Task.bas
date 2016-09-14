Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Task
' Level:        Framework class
' Version:      1.00
'
' Description:  Task object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
' =================================

'    [ID] [smallint] IDENTITY(1,1) NOT NULL,
'    [TaskType] [nvarchar](1) NOT NULL,
'    [Label] [nvarchar](1) NOT NULL,
'    [Summary] [smallint] NOT NULL,
'    [Priority] [smallint] NOT NULL,
'    [Status] [smallint] NOT NULL,
'    [FollowupNotes] [nvarchar](max) NULL,
'    [RequestDate] [date] NULL,
'    [CompleteDate] [date] NULL,
'    [RequestedBy] [int] NOT NULL,
'    [FollowupBy] [int] NULL,
'    [CompletedBy] [int] NULL,

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Task As String
Private m_TaskType As String
Private m_Priority As Integer
Private m_Status As Integer
Private m_RequestedByID As Integer
Private m_FollowupByID As Integer
Private m_CompletedByID As Integer
Private m_Requestor As Person
Private m_FollowupBy As Person
Private m_CompletedBy As Person
Private m_RequestDate As Date
Private m_FollowupDate As Date
Private m_CompleteDate As Date

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let Task(Value As String)
    If ValidateString(Value, "paragraph") Then
        m_Task = Value
    End If
End Property

Public Property Get Task() As String
    Task = m_Task
End Property

Public Property Let TaskType(Value As String)
    If ValidateString(Value, "paragraph") Then
        m_TaskType = Value
    End If
End Property

Public Property Get TaskType() As String
    TaskType = m_TaskType
End Property

Public Property Let Priority(Value As String)
    m_Priority = Value
End Property

Public Property Get Priority() As String
    Priority = m_Priority
End Property

Public Property Let Status(Value As String)
    m_Status = Value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let RequestedByID(Value As Integer)
    m_RequestedByID = Value
End Property

Public Property Get RequestedByID() As Integer
    RequestedByID = m_RequestedByID
End Property

Public Property Let FollowupByID(Value As Integer)
    m_FollowupByID = Value
End Property

Public Property Get FollowupByID() As Integer
    FollowupByID = m_FollowupByID
End Property

Public Property Let CompletedByID(Value As Integer)
    m_CompletedByID = Value
End Property

Public Property Get CompletedByID() As Integer
    CompletedByID = m_CompletedByID
End Property

Public Property Let Requestor(Value As String)
    m_Requestor = Value
End Property

Public Property Get Requestor() As String
    Requestor = m_Requestor
End Property

Public Property Let FollowupBy(Value As String)
    m_FollowupBy = Value
End Property

Public Property Get FollowupBy() As String
    FollowupBy = m_FollowupBy
End Property

Public Property Let CompletedBy(Value As String)
    m_CompletedBy = Value
End Property

Public Property Get CompletedBy() As String
    CompletedBy = m_CompletedBy
End Property

Public Property Let RequestDate(Value As Date)
    m_RequestDate = Value
End Property

Public Property Get RequestDate() As Date
    RequestDate = m_RequestDate
End Property

Public Property Let FollowupDate(Value As Date)
    m_FollowupDate = Value
End Property

Public Property Get FollowupDate() As Date
    FollowupDate = m_FollowupDate
End Property

Public Property Let CompleteDate(Value As Date)
    m_CompleteDate = Value
End Property

Public Property Get CompleteDate() As Date
    CompleteDate = m_CompleteDate
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Task])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Task])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          AddTask
' Description:  Add new task item
' Assumptions:  -
' Parameters:   context - what the task is about/task type (string)
'               task
'               recordID - ID for the record the task references (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/3/2015 - initial version
' ---------------------------------
Public Sub AddTask()
On Error GoTo Err_Handler

''context As String, recordID As Integer, description As String, _
'                    status As Integer, priority As Integer, requestor As Integer, _
'                    Optional completor As Integer
'
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim strSQL As String
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset("Task")
'
'    With rs
'        .AddNew
'        !TaskType = Me.TaskType
'        !Task = Me.Task
'        !Status = Me.Status
'        !Priority = Me.Priority
'        !RequestedBy = Me.RequestedByID
'        !RequestDate = Me.RequestDate
'        !CompletedBy = Me.CompletedByID
'        !CompleteDate = Me.CompleteDate
'        !LastUpdateBy = 1
'        !LastUpdate = Now()
'
'        .update
'        If IsNumeric(!ID) Then
'            Me.ID = !ID
'        End If
'    End With
    
    Me.SaveToDb False

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddTask[Task class])"
    End Select
    Resume Exit_Handler
End Sub

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
    
    Dim Template As String
    
    Template = "i_task"
    
    Dim Params(0 To 13) As Variant
    
    With Me
        Params(0) = "Task"
        Params(1) = .Task
        Params(2) = .Status
        Params(3) = .Priority
        Params(4) = .RequestedByID
        Params(5) = CDate(Format(.RequestDate, "YYYY-mm-dd"))
        Params(6) = .CompletedByID
        Params(7) = CDate(Format(.CompleteDate, "YYYY-mm-dd"))
        
        'params 8-11 --> createdate, lastmodified
        
        If IsUpdate Then
            Template = "u_task"
            Params(12) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Task])"
    End Select
    Resume Exit_Handler
End Sub