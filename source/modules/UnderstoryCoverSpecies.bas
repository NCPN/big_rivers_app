Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        UnderstoryCoverSpecies
' Level:        Framework class
' Version:      1.03
'
' Description:  UnderstoryCover species object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 4/19/2016
' References:   -
' Revisions:    BLC - 4/19/2016 - 1.00 - initial version
'               BLC - 6/10/2016 - 1.01 - revised booleans to byte (values 0 & 1 vs. 0 & -1)
'               BLC - 6/11/2016 - 1.02 - updated to use GetTemplate()
'               BLC - 8/8/2016  - 1.03 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
' =================================

'---------------------
' Declarations
'---------------------
Private m_CoverSpecies As New CoverSpecies

Private m_IsSeedling As Byte

'---------------------
' Events
'---------------------
Public Event InvalidIsSeedling(value As Byte)

'-- base events (coverspecies)
Public Event InvalidVegPlotID(value As String)
Public Event InvalidPercentCover(value As Integer)

'-- base events (species) --
Public Event InvalidMasterPlantCode(value As String)
Public Event InvalidLUCode(value As String)
Public Event InvalidFamily(value As String)
Public Event InvalidSpecies(value As String)
Public Event InvalidCode(value As String)

'---------------------
' Properties
'---------------------
Public Property Let IsSeedling(value As Byte)
    If varType(value) = vbByte Then
        m_IsSeedling = value
    Else
        RaiseEvent InvalidIsSeedling(value)
    End If
End Property

Public Property Get IsSeedling() As Byte
    IsSeedling = m_IsSeedling
End Property

' ---------------------------
' -- base class properties --
' ---------------------------
' NOTE: required since VBA does not support direct inheritance
'       or polymorphism like other OOP languages
' ---------------------------
' base class = Cover Species
' ---------------------------
Public Property Let VegPlotID(value As Long)
    m_CoverSpecies.VegPlotID = value
End Property

Public Property Get VegPlotID() As Long
    VegPlotID = m_CoverSpecies.VegPlotID
End Property

Public Property Let PercentCover(value As Integer)
    If IsBetween(value, 0, 100, True) Then
        m_CoverSpecies.PercentCover = value
    Else
        RaiseEvent InvalidPercentCover(value)
    End If
End Property

Public Property Get PercentCover() As Integer
    PercentCover = m_CoverSpecies.PercentCover
End Property

' ---------------------------
' base class = Species
' ---------------------------
Public Property Let ID(value As Long)
    m_CoverSpecies.ID = value
End Property

Public Property Get ID() As Long
    ID = m_CoverSpecies.ID
End Property

Public Property Let MasterPlantCode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_CoverSpecies.MasterPlantCode = value
    Else
        RaiseEvent InvalidMasterPlantCode(value)
    End If
End Property

Public Property Get MasterPlantCode() As String
    MasterPlantCode = m_CoverSpecies.MasterPlantCode
End Property

Public Property Let COfamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.COfamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get COfamily() As String
    COfamily = m_CoverSpecies.COfamily
End Property

Public Property Let UTfamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.UTfamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get UTfamily() As String
    UTfamily = m_CoverSpecies.UTfamily
End Property

Public Property Let WYfamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.WYfamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get WYfamily() As String
    WYfamily = m_CoverSpecies.WYfamily
End Property

Public Property Let COspecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.COspecies = value
    Else
        RaiseEvent InvalidSpecies(value)
    End If
End Property

Public Property Get COspecies() As String
    COspecies = m_CoverSpecies.COspecies
End Property

Public Property Let UTspecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.UTspecies = value
    Else
        RaiseEvent InvalidSpecies(value)
    End If
End Property

Public Property Get UTspecies() As String
    UTspecies = m_CoverSpecies.UTspecies
End Property

Public Property Let WYspecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.WYspecies = value
    Else
        RaiseEvent InvalidSpecies(value)
    End If
End Property

Public Property Get WYspecies() As String
    WYspecies = m_CoverSpecies.WYspecies
End Property

Public Property Let LUcode(value As String)
    'valid length varchar(25) but 6-letter lookup
    If Not IsNull(value) And IsBetween(Len(value), 1, 6, True) Then
        m_CoverSpecies.LUcode = value
    Else
        RaiseEvent InvalidLUCode(value)
    End If
End Property

Public Property Get LUcode() As String
    LUcode = m_CoverSpecies.LUcode
End Property

Public Property Let MasterFamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.MasterFamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get MasterFamily() As String
    MasterFamily = m_CoverSpecies.MasterFamily
End Property

Public Property Let MasterCode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_CoverSpecies.MasterCode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get MasterCode() As String
    MasterCode = m_CoverSpecies.MasterCode
End Property

Public Property Let MasterSpecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.MasterSpecies = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get MasterSpecies() As String
    MasterSpecies = m_CoverSpecies.MasterSpecies
End Property

Public Property Let UTcode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_CoverSpecies.UTcode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get UTcode() As String
    UTcode = m_CoverSpecies.UTcode
End Property

Public Property Let COcode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_CoverSpecies.COcode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get COcode() As String
    COcode = m_CoverSpecies.COcode
End Property

Public Property Let WYcode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_CoverSpecies.WYcode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get WYcode() As String
    WYcode = m_CoverSpecies.WYcode
End Property

Public Property Let MasterCommonName(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_CoverSpecies.MasterCommonName = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get MasterCommonName() As String
    MasterCommonName = m_CoverSpecies.MasterCommonName
End Property

Public Property Let Lifeform(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_CoverSpecies.Lifeform = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get Lifeform() As String
    Lifeform = m_CoverSpecies.Lifeform
End Property

Public Property Let Duration(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_CoverSpecies.Duration = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get Duration() As String
    Duration = m_CoverSpecies.Duration
End Property

Public Property Let Nativity(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_CoverSpecies.Nativity = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get Nativity() As String
    Nativity = m_CoverSpecies.Nativity
End Property


'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

'    MsgBox "Initializing...", vbOKOnly
    
    Set m_CoverSpecies = New CoverSpecies

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[cls_UnderstoryCoverSpecies])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  Class termination (closing) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler
    
'    MsgBox "Terminating...", vbOKOnly
        
    Set m_CoverSpecies = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_UnderstoryCoverSpecies])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup understory cover species based on the lookup code
' Parameters:   luCode - species 6-character lookup code from NCPN master plants (string)
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub Init(LUcode As String)
On Error GoTo Err_Handler
    
            m_CoverSpecies.Init (LUcode)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[cls_UnderstoryCoverSpecies])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save cover species based to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'   BLC, 6/11/2016 - revised to GetTemplate()
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
'    'record actions must have:
''    strSQL = "INSERT INTO UnderstorySpecies(VegPlot_ID, Master_PLANT_Code, PercentCover, IsSeedling) VALUES " _
''                & "(" & Me.VegPlotID & ",'" & Me.MasterPlantCode & "'," _
''                & Me.PercentCover & "," & Me.IsSeedling & ");"
'    strSQL = GetTemplate("i_understory_species", _
'                "vegplotID" & PARAM_SEPARATOR & Me.VegPlotID & _
'                "|masterplantcode" & PARAM_SEPARATOR & Me.MasterPlantCode & _
'                "|pctcover" & PARAM_SEPARATOR & Me.PercentCover & _
'                "|isseedling" & PARAM_SEPARATOR & Me.IsSeedling)
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim Template As String
    
    Template = "i_understory_species"
    
    Dim Params(0 To 6) As Variant

    With Me
        Params(0) = "UnderstorySpecies"
        Params(1) = .VegPlotID
        Params(2) = .MasterPlantCode
        Params(3) = .PercentCover
        Params(4) = .IsSeedling
        
        If IsUpdate Then
            Template = "u_understory_species"
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
                "Error encounter (#" & Err.Number & " - Init[cls_UnderstoryCoverSpecies])"
    End Select
    Resume Exit_Handler
End Sub