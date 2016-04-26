Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        ActionDate
' Level:        Framework class
' Version:      1.00
'
' Description:  Action date object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, November 10, 2015
' References:   -
' Revisions:    BLC - 11/10/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_FirstName As String
Private m_LastName As String
Private m_Name As String
Private m_Email As String
Private m_Role As String
Private m_Record As String
Private m_Contact As String
Private m_DateValue As Date
Private m_ActionType As String

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

Public Property Let FirstName(value As String)
    m_FirstName = value
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let Record(value As String)
    m_Record = value
End Property

Public Property Get Record() As String
    Record = m_Record
End Property

Public Property Let Contact(value As Person)
    m_Contact = value
End Property

Public Property Get Contact() As Person
    Contact = m_Contact
End Property

Public Property Let DateValue(value As Date)
    m_DateValue = value
End Property

Public Property Get DateValue() As Date
    DateValue = m_DateValue
End Property

Public Property Let ActionType(value As String)
    Select Case value
        Case "Sample"
        Case "DataEntry"
        Case "Verification"
        Case "Download"
        Case "Change"
    End Select
    m_ActionType = value
End Property

Public Property Get ActionType() As String
    ActionType = m_ActionType
End Property