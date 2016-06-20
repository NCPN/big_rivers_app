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
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let Task(value As String)
    If ValidateString(value, "paragraph") Then
        m_Task = value
    End If
End Property

Public Property Get Task() As String
    Task = m_Task
End Property

Public Property Let TaskType(value As String)
    If ValidateString(value, "paragraph") Then
        m_TaskType = value
    End If
End Property

Public Property Get TaskType() As String
    TaskType = m_TaskType
End Property

Public Property Let Priority(value As String)
    m_Priority = value
End Property

Public Property Get Priority() As String
    Priority = m_Priority
End Property

Public Property Let Status(value As String)
    m_Status = value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let RequestedByID(value As Integer)
    m_RequestedByID = value
End Property

Public Property Get RequestedByID() As Integer
    RequestedByID = m_RequestedByID
End Property

Public Property Let FollowupByID(value As Integer)
    m_FollowupByID = value
End Property

Public Property Get FollowupByID() As Integer
    FollowupByID = m_FollowupByID
End Property

Public Property Let CompletedByID(value As Integer)
    m_CompletedByID = value
End Property

Public Property Get CompletedByID() As Integer
    CompletedByID = m_CompletedByID
End Property

Public Property Let Requestor(value As String)
    m_Requestor = value
End Property

Public Property Get Requestor() As String
    Requestor = m_Requestor
End Property

Public Property Let FollowupBy(value As String)
    m_FollowupBy = value
End Property

Public Property Get FollowupBy() As String
    FollowupBy = m_FollowupBy
End Property

Public Property Let CompletedBy(value As String)
    m_CompletedBy = value
End Property

Public Property Get CompletedBy() As String
    CompletedBy = m_CompletedBy
End Property

Public Property Let RequestDate(value As Date)
    m_RequestDate = value
End Property

Public Property Get RequestDate() As Date
    RequestDate = m_RequestDate
End Property

Public Property Let FollowupDate(value As Date)
    m_FollowupDate = value
End Property

Public Property Get FollowupDate() As Date
    FollowupDate = m_FollowupDate
End Property

Public Property Let CompleteDate(value As Date)
    m_CompleteDate = value
End Property

Public Property Get CompleteDate() As Date
    CompleteDate = m_CompleteDate
End Property

'---------------------
' Methods
'---------------------
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

'context As String, recordID As Integer, description As String, _
                    status As Integer, priority As Integer, requestor As Integer, _
                    Optional completor As Integer

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Task")
    
    With rs
        .AddNew
        !TaskType = Me.TaskType
        !Task = Me.Task
        !Status = Me.Status
        !Priority = Me.Priority
        !RequestedBy = Me.RequestedByID
        !RequestDate = Me.RequestDate
        !CompletedBy = Me.CompletedByID
        !CompleteDate = Me.CompleteDate
        !LastUpdateBy = 1
        !LastUpdate = Now()
        
        .Update
        If IsNumeric(!ID) Then
            Me.ID = !ID
        End If
    End With
    

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