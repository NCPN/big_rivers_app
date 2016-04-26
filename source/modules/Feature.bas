Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Feature
' Level:        Framework class
' Version:      1.00
'
' Description:  Feature object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
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
Public Event InvalidSequence(sequence As Integer)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
   ID = m_ID
End Property

Public Property Let LocationID(value As Long)
    m_LocationID = value
End Property

Public Property Get LocationID() As Long
   LocationID = m_LocationID
End Property

Public Property Let Name(value As String)
    'feature names are 1 or 2 characters (letters only)
    If Len(Trim(value)) < 3 And IsAlpha(value) Then
        m_Name = value
    Else
        RaiseEvent InvalidName(value)
    End If
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Description(value As String)
    'descriptions - 255 characters or less
    If Len(Trim(value)) < 256 Then
        m_Description = value
    Else
        RaiseEvent InvalidDescription(value)
    End If
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Directions(value As String)
    m_Directions = value
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let sequence(value As Integer)
    If value > -1 Then
        m_Sequence = value
    Else
        RaiseEvent InvalidSequence(value)
    End If
End Property

Public Property Get sequence() As Integer
    sequence = m_Sequence
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb()
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    'events must have: start date, site ID, location ID, protocol ID
    strSQL = "INSERT INTO Feature(Location_ID, Feature, FeatureDescription, FeatureDirections) VALUES " _
                & "(" & Me.LocationID & ",'" & Me.Name & "','" _
                & Me.Description & "','" & Me.Directions & "');"

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

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