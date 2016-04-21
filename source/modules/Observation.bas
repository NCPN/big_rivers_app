Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Observation
' Level:        Framework class
' Version:      1.00
'
' Description:  Observation object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_ObservationType As String

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

Public Property Let ObservationType(Value As String)
    Select Case Value
        Case "WCC"  'Woody Canopy Cover
        Case "U"        'Understory
    End Select
    m_ObservationType = Value
End Property

Public Property Get ObservationType() As String
    ObservationType = m_ObservationType
End Property