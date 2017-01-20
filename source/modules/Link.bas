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
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let text(Value As String)
    m_Text = Value
End Property

Public Property Get text() As String
    text = m_Text
End Property

Public Property Let action(Value As String)
    m_Action = Value
End Property

Public Property Get action() As String
    action = m_Action
End Property

Public Property Let LinkFontColor(Value As Long)
    m_LinkFontColor = Value
End Property

Public Property Get LinkFontColor() As Long
    LinkFontColor = m_LinkFontColor
End Property

Public Property Let LinkBgColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_LinkBgColor = Value
    
    'set font color to match
    Select Case Value
        Case vbGreen
            Me.LinkFontColor = vbBlack
        Case vbRed, vbBlue
            Me.LinkFontColor = vbWhite
    End Select
End Property

Public Property Get LinkBgColor() As Long
    LinkBgColor = m_LinkBgColor 'FormHeader.BackColor
End Property

Public Property Let LinkVisible(Value As Byte)
    m_LinkVisible = Value
End Property

Public Property Get LinkVisible() As Byte
    LinkVisible = m_LinkVisible
End Property

Public Property Let LinkSeparatorVisible(Value As Byte)
    m_LinkSeparatorVisible = Value
End Property

Public Property Get LinkSeparatorVisible() As Byte
    LinkSeparatorVisible = m_LinkSeparatorVisible
End Property