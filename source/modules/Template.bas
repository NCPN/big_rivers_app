Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Template
' Level:        Framework class
' Version:      1.00
'
' Description:  Template object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/4/2016
' References:   -
' Revisions:    BLC - 10/4/2016 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long

Private m_EventID As Long

Private m_TemplateName As String '255
Private m_Context As String '255
Private m_Syntax As String '10
Private m_TemplateSQL As String 'memo
Private m_Params As String '255
Private m_Version As String '10
Private m_IsSupported As Integer
Private m_Remarks As String '255
Private m_EffectiveDate As Date 'date
Private m_RetireDate As Date 'date

'creator/modifier
Private m_ContactID As Long

'---------------------
' Events
'---------------------
Public Event InvalidTemplateSQL(Value As String)
Public Event InvalidSyntax(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let TemplateName(Value As String)
    m_TemplateName = Value
End Property

Public Property Get TemplateName() As String
    TemplateName = m_TemplateName
End Property

Public Property Let Context(Value As String)
    m_Context = Value
End Property

Public Property Get Context() As String
    Context = m_Context
End Property

Public Property Let TemplateSQL(Value As String)
    m_TemplateSQL = Value
    
    'set params property
    If Len(Me.Syntax) <> Len(Replace(Me.Syntax, "SQL", "")) Then
        Me.Params = GetParamsFromSQL(Me.TemplateSQL)
    End If
    
End Property

Public Property Get TemplateSQL() As String
    TemplateSQL = m_TemplateSQL
End Property

Public Property Let Syntax(Value As String)
    m_Syntax = Value
End Property

Public Property Get Syntax() As String
    Syntax = m_Syntax
End Property

Public Property Let Params(Value As String)
    m_Params = Value
End Property

Public Property Get Params() As String
    Params = m_Params
End Property

Public Property Let Version(Value As String)
    m_Version = Value
End Property

Public Property Get Version() As String
    Version = m_Version
End Property

Public Property Let Remarks(Value As String)
    m_Remarks = Value
End Property

Public Property Get Remarks() As String
    Remarks = m_Remarks
End Property

Public Property Let IsSupported(Value As Integer)
    m_IsSupported = Value
End Property

Public Property Get IsSupported() As Integer
    IsSupported = m_IsSupported
End Property

Public Property Let EffectiveDate(Value As Date)
    m_EffectiveDate = Format(Value, "mm/dd/yyyy")
End Property

Public Property Get EffectiveDate() As Date
    EffectiveDate = m_EffectiveDate
End Property

Public Property Let RetireDate(Value As Date)
    m_RetireDate = Format(Value, "mm/dd/yyyy")
End Property

Public Property Get RetireDate() As Date
    RetireDate = m_RetireDate
End Property

Public Property Let ContactID(Value As Long)
    m_ContactID = Value
End Property

Public Property Get ContactID() As Long
    ContactID = m_ContactID
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_Template])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_Template])"
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

    Dim Template As String
    
    Template = "i_template"
    
    Dim Params(0 To 12) As Variant

    With Me
        Params(0) = "Template"
        Params(1) = .TemplateName
        Params(2) = .Context
        Params(3) = .TemplateSQL
        Params(4) = .Remarks
        Params(5) = .EffectiveDate
        Params(6) = .ContactID
        Params(7) = .Params
        Params(8) = .Syntax
        Params(9) = .Version
        Params(10) = .IsSupported
        Params(11) = .RetireDate
    
        If IsUpdate Then
            Template = "u_template"
            Params(12) = .ID
        End If

        .ID = SetRecord(Template, Params)
    End With
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_Template])"
    End Select
    Resume Exit_Handler
End Sub