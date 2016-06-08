Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Person
' Level:        Framework class
' Version:      1.00
'
' Description:  Person object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_FirstName As String
Private m_LastName As String
Private m_MiddleInitial As String
Private m_Name As String
Private m_Email As String
Private m_Organization As String
Private m_PosTitle As String
Private m_WorkPhone As Integer
Private m_WorkExtension As Integer
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

Public Property Let MiddleInitial(Value As String)
    If IsAlpha(Value) And Len(Value) = 1 Then
        m_MiddleInitial = Value
    Else
        RaiseEvent InvalidInitial(Value)
    End If
End Property

Public Property Get MiddleInitial() As String
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

Public Property Let WorkPhone(Value As String)
    If IsPhone(Value) Then
        m_WorkPhone = Value
    Else
        RaiseEvent InvalidPhone(Value)
    End If
End Property

Public Property Get WorkPhone() As String
    WorkPhone = m_WorkPhone
End Property

Public Property Let WorkExtension(Value As String)
    m_WorkExtension = Value
End Property

Public Property Get WorkExtension() As String
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
    
    'persons must have: first & last name, email, organization
    'optional: middleinitial, username, workphone, workextension, positiontitle
'    strSQL = "INSERT INTO Contact(FirstName, LastName, Email, Organization," _
'                & "MiddleInitial, Username, WorkPhone, WorkExtension, PositionTitle) VALUES " _
'                & "('" & Me.FirstName & "','" & Me.LastName & "','" _
'                & Me.Email & "','" & Me.Organization & "','" _
'                & Me.MiddleInitial & "','" & Me.Username & "','" _
'                & Me.WorkPhone & "','" & Me.WorkExtension & "','" _
'                & Me.PositionTitle & "');"
    strSQL = GetTemplate("i_contact", _
                "FirstName" & PARAM_SEPARATOR & Me.FirstName & "|" & _
                "LastName" & PARAM_SEPARATOR & Me.LastName & "|" & _
                "email" & PARAM_SEPARATOR & Me.Email & "|" & _
                "org" & PARAM_SEPARATOR & Me.Organization & "|" & _
                "MI" & PARAM_SEPARATOR & Me.MiddleInitial & "|" & _
                "username" & PARAM_SEPARATOR & Me.Username & "|" & _
                "WorkPhone" & PARAM_SEPARATOR & Me.WorkPhone & "|" & _
                "WorkExt" & PARAM_SEPARATOR & Me.WorkExtension & "|" & _
                "position" & PARAM_SEPARATOR & Me.PosTitle & "|" & _
                "IsActive" & PARAM_SEPARATOR & Me.IsActive)
    
    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    'set the person's role
    
'    strSQL = "INSERT INTO Contact_Access(Contact_ID, Access_ID)" _
'                & "VALUES (" & Me.ID & "," &  & ");"

    strSQL = GetTemplate("i_contact_access", _
                "contactID" & PARAM_SEPARATOR & Me.ID & "|" & _
                "accessID" & PARAM_SEPARATOR & Me.AccessLevel)
    db.Execute strSQL, dbFailOnError

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