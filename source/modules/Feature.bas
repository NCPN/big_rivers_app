Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Feature
' Level:        Framework class
' Version:      1.01
'
' Description:  Feature object related properties, events, functions & procedures
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
Private m_ID As Integer
Private m_LocationID As Integer
Private m_Name As String
Private m_Description As String
Private m_Directions As String
Private m_Sequence As Integer

'---------------------
' Events
'---------------------
Public Event InvalidID()
Public Event InvalidName(Name As String)
Public Event InvalidDescription(Description As String)
Public Event InvalidDirections(Directions As String)
Public Event InvalidSequence(Sequence As Integer)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
   ID = m_ID
End Property

Public Property Let LocationID(Value As Long)
    m_LocationID = Value
End Property

Public Property Get LocationID() As Long
   LocationID = m_LocationID
End Property

Public Property Let Name(Value As String)
    'feature names are 1 or 2 characters (letters only)
    If Len(Trim(Value)) < 3 And IsAlpha(Value) Then
        m_Name = Value
    Else
        RaiseEvent InvalidName(Value)
    End If
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Description(Value As String)
    'descriptions - 255 characters or less
    If Len(Trim(Value)) < 256 Then
        m_Description = Value
    Else
        RaiseEvent InvalidDescription(Value)
    End If
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Directions(Value As String)
    m_Directions = Value
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let Sequence(Value As Integer)
    If Value > -1 Then
        m_Sequence = Value
    Else
        RaiseEvent InvalidSequence(Value)
    End If
End Property

Public Property Get Sequence() As Integer
    Sequence = m_Sequence
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Feature])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Feature])"
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'
'    Set db = CurrentDb
'
'    'events must have: start date, site ID, location ID, protocol ID
'    strSQL = "INSERT INTO Feature(Location_ID, Feature, FeatureDescription, FeatureDirections) VALUES " _
'                & "(" & Me.LocationID & ",'" & Me.Name & "','" _
'                & Me.Description & "','" & Me.Directions & "');"
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim Template As String
    
    Template = "i_feature"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "Feature"
        Params(1) = .LocationID
        Params(2) = .Name
        Params(3) = .Description
        Params(4) = .Directions
        
        If IsUpdate Then
            Template = "u_feature"
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
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Feature])"
    End Select
    Resume Exit_Handler
End Sub