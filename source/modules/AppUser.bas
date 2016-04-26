Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        AppUser
' Level:        Framework class
' Version:      1.00
'
' Description:  Application User object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Dim AppUser As New Person

Private m_Username As String
Private m_Password As String
Private m_Logins As Integer

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let UserName(value As String)
    m_Username = value
End Property

Public Property Get UserName() As String
    UserName = m_Username
End Property

Public Property Let Password(value As String)
    m_Password = value
End Property

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Logins(value As Integer)
    m_Logins = value
End Property

Public Property Get Logins() As Integer
    Logins = m_Logins
End Property