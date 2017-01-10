Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        AppComment
' Level:        Framework class
' Version:      1.00
'
' Description:  Comment object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/9/2016   - 1.01 - added SaveToDb() revised to AppComment (Comment reserved word)
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_CommentType As String
Private m_TypeID As Integer
Private m_Comment As String
Private m_CommentDate As Date
Private m_CommentorID As Integer    'Long??
Private m_MaxLength As Integer

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    If IsNumeric(Value) Then
        m_ID = Value
    End If
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let TypeID(Value As Integer)
    If IsNumeric(Value) Then
        m_TypeID = Value
    End If
End Property

Public Property Get TypeID() As Integer
    TypeID = m_TypeID
End Property

Public Property Let CommentType(Value As String)
    If ValidateString(Value, "alphanumdashslashspace") Then
        m_CommentType = Value
    End If
End Property

Public Property Get CommentType() As String
    CommentType = m_CommentType
End Property

Public Property Let Comment(Value As String)
    If ValidateString(Value, "paragraph") Then
        m_Comment = Value
    End If
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let CommentorID(Value As Integer)
    If IsNumeric(Value) Then
        m_CommentorID = Value
    End If
End Property

Public Property Get CommentorID() As Integer
    ID = m_CommentorID
End Property

Public Property Let MaxLength(Value As Integer)
    If IsNumeric(Value) Then
        m_MaxLength = Value
    End If
End Property

Public Property Get MaxLength() As Integer
    MaxLength = m_MaxLength
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_AppComment])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_AppComment])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          AddComment
' Description:  Add new Comment item
' Assumptions:  -
' Parameters:   context - what the Comment is about/Comment type (string)
'               Comment
'               recordID - ID for the record the Comment references (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 19, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/19/2015 - initial version
' ---------------------------------
Public Sub AddComment()
On Error GoTo Err_Handler

'context As String, recordID As Integer, description As String, _
                    status As Integer, priority As Integer, requestor As Integer, _
                    Optional completor As Integer

'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim strSQL As String
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset("Comment")
'
'    With rs
'        .AddNew
'        !CommentType = Me.CommentType
'        !TypeID = Me.TypeID
'        !Comment = Me.Comment
'        !CreatedBy = Me.CommentorID
'        !CreateDate = Now()
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
            "Error encountered (#" & Err.Number & " - AddComment[AppComment class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  -
' Parameters:   IsUpdate - indicates if data is an update vs. an insert (boolean, optional)
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 8/9/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_comment"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "Comment"
        Params(1) = .CommentType
        Params(2) = .TypeID
        Params(3) = .Comment
        Params(4) = .CommentorID
        
        If IsUpdate Then
            Template = "u_comment"
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
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_AppComment])"
    End Select
    Resume Exit_Handler
End Sub