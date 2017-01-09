Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Link
' Level:        Framework class
' Version:      1.00
'
' Description:  Link object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/30/2015
' References:
'  Maciej Los, April 5, 2011
'  http://www.codeproject.com/Questions/167323/Using-a-VS-Custom-Control-in-VBA-NOT-VB
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Text As String
Private m_Action As String
Private m_LinkFontColor As Long
Private m_LinkBgColor As Long
Private m_LinkVisible As Byte
Private m_LinkSeparatorVisible As Byte

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

Public Property Let Text(value As String)
    m_Text = value
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let action(value As String)
    m_Action = value
End Property

Public Property Get action() As String
    action = m_Action
End Property

Public Property Let LinkFontColor(value As Long)
    m_LinkFontColor = value
End Property

Public Property Get LinkFontColor() As Long
    LinkFontColor = m_LinkFontColor
End Property

Public Property Let LinkBgColor(value As Long)
    If Len(Trim(value)) < 0 Then value = vbGreen '"#3F3F3F"
    m_LinkBgColor = value
    
    'set font color to match
    Select Case value
        Case vbGreen
            Me.LinkFontColor = vbBlack
        Case vbRed, vbBlue
            Me.LinkFontColor = vbWhite
    End Select
End Property

Public Property Get LinkBgColor() As Long
    LinkBgColor = m_LinkBgColor 'FormHeader.BackColor
End Property

Public Property Let LinkVisible(value As Byte)
    m_LinkVisible = value
End Property

Public Property Get LinkVisible() As Byte
    LinkVisible = m_LinkVisible
End Property

Public Property Let LinkSeparatorVisible(value As Byte)
    m_LinkSeparatorVisible = value
End Property

Public Property Get LinkSeparatorVisible() As Byte
    LinkSeparatorVisible = m_LinkSeparatorVisible
End Property