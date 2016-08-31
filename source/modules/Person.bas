Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Person
' Level:        Framework class
' Version:      1.01
'
' Description:  Person object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/8/2016   - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim iCount As Integer
'
'    '---------------------
'    ' CurrentDb returns a different dao.database reference every time it is called
'    ' So set a reference to make sure CurrentDb isn't called multiple times
'    ' This ensures @@Identity will contain the proper ID #
'    '---------------------
'    Set db = CurrentDb
'
'    'default
'    iCount = 0
'
'    'persons must have: first & last name, email, organization
'    'optional: middleinitial, username, workphone, workextension, positiontitle
''    strSQL = "INSERT INTO Contact(FirstName, LastName, Email, Organization," _
''                & "MiddleInitial, Username, WorkPhone, WorkExtension, PositionTitle) VALUES " _
''                & "('" & Me.FirstName & "','" & Me.LastName & "','" _
''                & Me.Email & "','" & Me.Organization & "','" _
''                & Me.MiddleInitial & "','" & Me.Username & "','" _
''                & Me.WorkPhone & "','" & Me.WorkExtension & "','" _
''                & Me.PositionTitle & "');"
''    strSQL = GetTemplate("i_contact", _
''                "FirstName" & PARAM_SEPARATOR & Me.FirstName & "|" & _
''                "LastName" & PARAM_SEPARATOR & Me.LastName & "|" & _
''                "email" & PARAM_SEPARATOR & Me.Email & "|" & _
''                "org" & PARAM_SEPARATOR & Me.Organization & "|" & _
''                "MI" & PARAM_SEPARATOR & IIf(IsZLS(Me.MiddleInitial), DbNull, Me.MiddleInitial) & "|" & _
''                "username" & PARAM_SEPARATOR & Me.Username & "|" & _
''                "WorkPhone" & PARAM_SEPARATOR & IIf(IsZero(Me.WorkPhone), DbNull, Me.WorkPhone) & "|" & _
''                "WorkExt" & PARAM_SEPARATOR & IIf(IsZero(Me.WorkExtension), DbNull, Me.WorkExtension) & "|" & _
''                "position" & PARAM_SEPARATOR & IIf(IsZLS(Me.PosTitle), DbNull, Me.PosTitle) & "|" & _
''                "IsActive" & PARAM_SEPARATOR & Me.IsActive)
'
'    Dim qdf As DAO.QueryDef
'
'    With db
'        Set qdf = .QueryDefs("usys_temp_qdf")
'
'        With qdf
'            If Me.ID > 0 Then
'                .SQL = GetTemplate("u_contact")
'                .Parameters("ContactID") = Me.ID
'            Else
'                .SQL = GetTemplate("i_contact_new")
'            End If
'            '-- required parameters --
''            .Parameters("FirstName") = Me.FirstName
''            .Parameters("LastName") = Me.LastName
''            .Parameters("Email") = Me.Email
''            .Parameters("Username") = Me.Username
'            .Parameters("First") = Me.FirstName
'            .Parameters("Last") = Me.LastName
'            .Parameters("EmailAddress") = Me.Email
'            .Parameters("Login") = Me.Username
'            .Parameters("Org") = Me.Organization
'
'            '-- optional parameters --
' '           If Not IsZLS(Me.MiddleInitial) Then _
'                .Parameters("MiddleInitial") = Me.MiddleInitial
'                .Parameters("MI") = Me.MiddleInitial
'
''            If Not IsZLS(Me.PosTitle) Then '_
'                .Parameters("Position") = Me.PosTitle
'
''            If Not IsZero(Me.WorkPhone) Then '_
'                .Parameters("Phone") = Me.WorkPhone
'
''            If Not IsZero(Me.WorkExtension) Then '_
'                .Parameters("Ext") = Me.WorkExtension
'
''            .Parameters("IsActive") = Me.IsActive
'            .Parameters("IsActiveFlag") = Me.IsActive
'
'            .Execute dbFailOnError
'
'            'cleanup
'            .Close
'        End With
'
''    db.Execute strSQL, dbFailOnError
'        If Not Me.ID > 0 Then _
'            Me.ID = .OpenRecordset("SELECT @@IDENTITY")(0)
'
'    'set the person's role
'
''    strSQL = "INSERT INTO Contact_Access(Contact_ID, Access_ID)" _
''                & "VALUES (" & Me.ID & "," &  & ");"
'
''    strSQL = GetTemplate("i_contact_access", _
''                "contactID" & PARAM_SEPARATOR & Me.ID & "|" & _
''                "accessID" & PARAM_SEPARATOR & Me.AccessLevel)
''
''    db.Execute strSQL, dbFailOnError
'
'        Set qdf = .QueryDefs("usys_temp_qdf")
'
'        With qdf
'            'check if value exists in contact_access
'            .SQL = GetTemplate("s_count_tbl", _
'                    "field" & PARAM_SEPARATOR & "Contact_ID" & _
'                    "|tbl" & PARAM_SEPARATOR & "Contact_Access WHERE Contact_ID = " & Me.ID)
'            Set rs = .OpenRecordset
'            If rs.Fields(0) > 0 Then iCount = rs.Fields(0)
'        End With
'
'        Set qdf = .QueryDefs("usys_temp_qdf")
'
'        With qdf
'            'update if contact is in contact_access, otherwise insert new record
'            If iCount > 0 Then 'Me.AccessLevel
'                .SQL = GetTemplate("u_contact_access")
'            Else
'                .SQL = GetTemplate("i_contact_access")
'            End If
'
'            '-- required parameters --
'            .Parameters("ContactID") = Me.ID
'            .Parameters("AccessID") = Me.AccessLevel
'
'            '-- optional parameters --
'
'            .Execute dbFailOnError
'
'            'cleanup
'            .Close
'        End With
'    End With

    Dim template As String
    
    template = "i_contact" '"i_contact_new"
    
    Dim params() As Variant
    
    'dimension for contact
    ReDim params(0 To 12) As Variant

    With Me
        params(0) = "Contact"
        params(1) = .FirstName
        params(2) = .LastName
        params(3) = .Email
        params(4) = .Username
        params(5) = .Organization
        
        params(6) = IIf(Len(.MiddleInitial) > 0, .MiddleInitial, Null)
        params(7) = IIf(Len(.PosTitle) > 0, .PosTitle, Null)
        params(8) = IIf(.WorkPhone = 0, Null, .WorkPhone)
        params(9) = IIf(.WorkExtension = 0, Null, .WorkExtension)
        params(10) = IIf(.IsActive > 0, .IsActive, 0) 'Null)
        
        If IsUpdate Then
            template = "u_contact"
            params(11) = .ID
        End If
        
        .ID = SetRecord(template, params)
    End With

    'set the person's role (update if contact template was update)
    template = IIf(Left(template, 1) = "u", "u_contact_access", "i_contact_access")
    
    'dimension for role
    ReDim params(0 To 3) As Variant

    With Me
        params(0) = "Contact_Access"
        params(1) = .ID
        params(2) = .AccessLevel
        
        'ID not generated here
        SetRecord template, params
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