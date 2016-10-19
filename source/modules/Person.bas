Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Person
' Level:        Framework class
' Version:      1.02
'
' Description:  Person object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/8/2016   - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               BLC - 9/1/2016   - 1.02 - SaveToDb() code cleanup
'               BLC - 10/15/2016 - 1.03 - adjust SaveToDb() to accommodate non-users
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_FirstName As String
Private m_LastName As String
Private m_MiddleInitial As Variant 'String
Private m_Name As String
Private m_Email As String
Private m_Organization As String
Private m_PosTitle As String
Private m_WorkPhone As Variant 'Integer
Private m_WorkExtension As Variant 'Integer
Private m_Role As String
Private m_AccessLevel As Integer
Private m_AccessRole As String
Private m_Username As String
Private m_IsActive As Byte 'using byte to avoid Access vs. SQL boolean issues

'---------------------
' Events
'---------------------
Public Event InvalidName(Value)
Public Event InvalidInitial(Value)
Public Event InvalidEmail(Value)
Public Event InvalidRole(Value)
Public Event InvalidPhone(Value)
Public Event InvalidAccessRole(Value)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let FirstName(Value As String)
    If IsName(Value) Then
        m_FirstName = Value
    Else
        RaiseEvent InvalidName(Value)
    End If
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let LastName(Value As String)
    If IsName(Value) Then
        m_LastName = Value
    Else
        RaiseEvent InvalidName(Value)
    End If
End Property

Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Let MiddleInitial(Value As Variant) 'String)
    If IsAlpha(CStr(Value)) And Len(Value) = 1 Then
        m_MiddleInitial = Value
    Else
        m_MiddleInitial = Null
'        RaiseEvent InvalidInitial(Value)
    End If
End Property

Public Property Get MiddleInitial() As Variant 'String
    MiddleInitial = m_MiddleInitial
End Property

Public Property Let Name(Value As String) 'Optional first As String, Value As String)
    m_Name = Value
End Property

Public Property Get Name() As String 'Optional first As String) As String
    Name = m_Name
End Property

Public Property Let Email(Value As String)
    If IsEmail(Value) Then
        m_Email = Value
    Else
        RaiseEvent InvalidEmail(Value)
    End If
End Property

Public Property Get Email() As String
    Email = m_Email
End Property

Public Property Let Organization(Value As String)
    m_Organization = Value
End Property

Public Property Get Organization() As String
    Organization = m_Organization
End Property

Public Property Let WorkPhone(Value As Variant) 'Integer)
    If Not IsNull(Value) Then
        If IsPhone(CStr(Value)) Then
            m_WorkPhone = Value
        Else
           RaiseEvent InvalidPhone(Value)
        End If
    Else
        m_WorkPhone = Null
        'RaiseEvent InvalidPhone(Value)
    End If
End Property

Public Property Get WorkPhone() As Variant 'Integer
    WorkPhone = m_WorkPhone
End Property

Public Property Let WorkExtension(Value As Variant) 'Long)
    m_WorkExtension = Value
End Property

Public Property Get WorkExtension() As Variant 'Long
    WorkExtension = m_WorkExtension
End Property

Public Property Let PosTitle(Value As String)
    m_PosTitle = Value
End Property

Public Property Get PosTitle() As String
    PosTitle = m_PosTitle
End Property

Public Property Let Username(Value As String)
    m_Username = Value
End Property

Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let IsActive(Value As Byte)
    m_IsActive = Value
End Property

Public Property Get IsActive() As Byte
    IsActive = m_IsActive
End Property

Public Property Let Role(Value As String)
    Dim aryRoles() As String
    aryRoles = Split(CONTACT_ROLES, ",")

    If IsInArray(Value, aryRoles) Then
        m_Role = Value
    Else
        RaiseEvent InvalidRole(Value)
    End If
End Property

Public Property Get Role() As String
    Role = m_Role
End Property

Public Property Let AccessRole(Value As String)
    Dim aryAccessRoles() As String
    aryAccessRoles = Split(ACCESS_ROLES, ",")

    If IsInArray(Value, aryAccessRoles) Then
        m_AccessRole = Value
        'set access level based on role name (admin, power user, data entry, read only)
        AccessLevel = AccessID(m_AccessRole)
    Else
        RaiseEvent InvalidAccessRole(Value)
    End If
End Property

Public Property Get AccessRole() As String
    AccessRole = m_AccessRole
End Property

Public Property Let AccessLevel(Value As Long)
    m_AccessLevel = Value
End Property

Public Property Get AccessLevel() As Long
    AccessLevel = m_AccessLevel
End Property

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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Person])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Person])"
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
'   MarkK, September 11, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=253284
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'   BLC, 6/22/2016 - revised to use parameterized qdf
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'   BLC, 9/1/2016 - commented code cleanup
'   BLC, 10/14/2016 - adjusted to accommodate photographers who are non-app users
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_contact" '"i_contact_new"
    
    Dim Params() As Variant
    
    'dimension for contact
    ReDim Params(0 To 12) As Variant

    With Me
        Params(0) = "Contact"
        Params(1) = .FirstName
        Params(2) = .LastName
        Params(3) = .Email
        
        Params(4) = IIf(Len(.Username) > 0, .Username, Null)
        Params(5) = IIf(Len(.Organization) > 0, .Organization, Null)
        Params(6) = IIf(Len(.MiddleInitial) > 0, .MiddleInitial, Null)
        Params(7) = IIf(Len(.PosTitle) > 0, .PosTitle, Null)
        Params(8) = IIf(.WorkPhone = 0, Null, .WorkPhone)
        Params(9) = IIf(.WorkExtension = 0, Null, .WorkExtension)
        Params(10) = IIf(.IsActive > 0, .IsActive, 0) 'Null)
        
        If IsUpdate Then
            Template = "u_contact"
            Params(11) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
    'skip access if not set (non-users)
    If Not (Me.AccessLevel > 0) Then GoTo Exit_Handler

    'set the person's role (update if contact template was update)
    Template = IIf(Left(Template, 1) = "u", "u_contact_access", "i_contact_access")
    
    'dimension for role
    ReDim Params(0 To 3) As Variant

    With Me
        Params(0) = "Contact_Access"
        Params(1) = .ID
        Params(2) = .AccessLevel
        
        'ID not generated here
        SetRecord Template, Params
    End With
    

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Person])"
    End Select
    Resume Exit_Handler
End Sub