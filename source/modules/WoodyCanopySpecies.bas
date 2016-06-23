Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        WoodyCanopySpecies
' Level:        Framework class
' Version:      1.02
'
' Description:  Woody Canopy cover species object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 4/19/2016
' References:   -
' Revisions:    BLC - 4/19/2016 - 1.00 - initial version
'               BLC - 6/10/2016 - 1.01 - revised booleans to byte (values 0 & 1 vs. 0 & -1)
'               BLC - 6/11/2016 - 1.01 - updated to use GetTemplate()
' =================================

'---------------------
' Declarations
'---------------------
Private m_CoverSpecies As New CoverSpecies

Private m_IsSeedling As Byte

'---------------------
' Events
'---------------------
'-- base events (coverspecies)
Public Event InvalidVegPlotID(Value As String)
Public Event InvalidPercentCover(Value As Integer)
Public Event InvalidIsSeedling(Value As Byte)

'-- base events (species) --
Public Event InvalidMasterPlantCode(Value As String)
Public Event InvalidLUCode(Value As String)
Public Event InvalidFamily(Value As String)
Public Event InvalidSpecies(Value As String)
Public Event InvalidCode(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let IsSeedling(Value As Byte)
    If varType(Value) = vbByte Then
        m_IsSeedling = Value
    Else
        RaiseEvent InvalidIsSeedling(Value)
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
Public Property Let VegPlotID(Value As Long)
    m_CoverSpecies.VegPlotID = Value
End Property

Public Property Get VegPlotID() As Long
    VegPlotID = m_CoverSpecies.VegPlotID
End Property

Public Property Let PercentCover(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_CoverSpecies.PercentCover = Value
    Else
        RaiseEvent InvalidPercentCover(Value)
    End If
End Property

Public Property Get PercentCover() As Integer
    PercentCover = m_CoverSpecies.PercentCover
End Property

' ---------------------------
' base class = Species
' ---------------------------
Public Property Let ID(Value As Long)
    m_CoverSpecies.ID = Value
End Property

Public Property Get ID() As Long
    ID = m_CoverSpecies.ID
End Property

Public Property Let MasterPlantCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.MasterPlantCode = Value
    Else
        RaiseEvent InvalidMasterPlantCode(Value)
    End If
End Property

Public Property Get MasterPlantCode() As String
    MasterPlantCode = m_CoverSpecies.MasterPlantCode
End Property

Public Property Let COfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.COfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get COfamily() As String
    COfamily = m_CoverSpecies.COfamily
End Property

Public Property Let UTfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.UTfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get UTfamily() As String
    UTfamily = m_CoverSpecies.UTfamily
End Property

Public Property Let WYfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.WYfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get WYfamily() As String
    WYfamily = m_CoverSpecies.WYfamily
End Property

Public Property Let COspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.COspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get COspecies() As String
    COspecies = m_CoverSpecies.COspecies
End Property

Public Property Let UTspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.UTspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get UTspecies() As String
    UTspecies = m_CoverSpecies.UTspecies
End Property

Public Property Let WYspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.WYspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get WYspecies() As String
    WYspecies = m_CoverSpecies.WYspecies
End Property

Public Property Let LUcode(Value As String)
    'valid length varchar(25) but 6-letter lookup
    If Not IsNull(Value) And IsBetween(Len(Value), 1, 6, True) Then
        m_CoverSpecies.LUcode = Value
    Else
        RaiseEvent InvalidLUCode(Value)
    End If
End Property

Public Property Get LUcode() As String
    LUcode = m_CoverSpecies.LUcode
End Property

Public Property Let MasterFamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.MasterFamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterFamily() As String
    MasterFamily = m_CoverSpecies.MasterFamily
End Property

Public Property Let MasterCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.MasterCode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCode() As String
    MasterCode = m_CoverSpecies.MasterCode
End Property

Public Property Let MasterSpecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.MasterSpecies = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterSpecies() As String
    MasterSpecies = m_CoverSpecies.MasterSpecies
End Property

Public Property Let UTcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.UTcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get UTcode() As String
    UTcode = m_CoverSpecies.UTcode
End Property

Public Property Let COcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.COcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get COcode() As String
    COcode = m_CoverSpecies.COcode
End Property

Public Property Let WYcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.WYcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get WYcode() As String
    WYcode = m_CoverSpecies.WYcode
End Property

Public Property Let MasterCommonName(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.MasterCommonName = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCommonName() As String
    MasterCommonName = m_CoverSpecies.MasterCommonName
End Property

Public Property Let Lifeform(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_CoverSpecies.Lifeform = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Lifeform() As String
    Lifeform = m_CoverSpecies.Lifeform
End Property

Public Property Let Duration(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_CoverSpecies.Duration = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Duration() As String
    Duration = m_CoverSpecies.Duration
End Property

Public Property Let Nativity(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_CoverSpecies.Nativity = Value
    Else
        RaiseEvent InvalidCode(Value)
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
            "Error encountered (#" & Err.Number & " - Class_Initialize[cls_WoodyCanopySpecies])"
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
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_WoodyCanopy])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup woody canopy cover species based on the lookup code
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
                "Error encounter (#" & Err.Number & " - Init[cls_WoodyCanopySpecies])"
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb()
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    'record actions must have:
'    strSQL = "INSERT INTO WoodyCanopySpecies(VegPlot_ID, Master_PLANT_Code, PercentCover) VALUES " _
'                & "(" & Me.VegPlotID & ",'" & Me.MasterPlantCode & "'," _
'                & Me.PercentCover & ");"
    strSQL = GetTemplate("i_cover_species", _
                "tbl" & PARAM_SEPARATOR & "WoodyCanopySpecies" & _
                "vegplotID" & PARAM_SEPARATOR & Me.VegPlotID & _
                "|masterplantcode" & PARAM_SEPARATOR & Me.MasterPlantCode & _
                "|pctcover" & PARAM_SEPARATOR & Me.PercentCover)

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[cls_WoodyCanopySpecies])"
    End Select
    Resume Exit_Handler
End Sub