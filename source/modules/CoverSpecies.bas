Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        CoverSpecies
' Level:        Framework class
' Version:      1.00
'
' Description:  Cover Species object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 4/13/2016
' References:   -
' Revisions:    BLC - 4/13/2016 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_Species As New Species

Private m_PercentCover As Integer
Private m_VegPlotID As Long

'---------------------
' Events
'---------------------
Public Event InvalidVegPlotID(Value As String)
Public Event InvalidPercentCover(Value As Integer)

'-- base events --
Public Event InvalidMasterPlantCode(Value As String)
Public Event InvalidLUCode(Value As String)
Public Event InvalidFamily(Value As String)
Public Event InvalidSpecies(Value As String)
Public Event InvalidCode(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let VegPlotID(Value As Long)
    m_VegPlotID = Value
End Property

Public Property Get VegPlotID() As Long
    VegPlotID = m_VegPlotID
End Property

Public Property Let PercentCover(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_PercentCover = Value
    Else
        RaiseEvent InvalidPercentCover(Value)
    End If
End Property

Public Property Get PercentCover() As Integer
    PercentCover = m_PercentCover
End Property

' ---------------------------
' -- base class properties --
' ---------------------------
' NOTE: required since VBA does not support direct inheritance
'       or polymorphism like other OOP languages
' ---------------------------
' base class = Species
' ---------------------------
Public Property Let ID(Value As Long)
    m_Species.ID = Value
End Property

Public Property Get ID() As Long
    ID = m_Species.ID
End Property

Public Property Let MasterPlantCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.MasterPlantCode = Value
    Else
        RaiseEvent InvalidMasterPlantCode(Value)
    End If
End Property

Public Property Get MasterPlantCode() As String
    MasterPlantCode = m_Species.MasterPlantCode
End Property

Public Property Let COfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.COfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get COfamily() As String
    COfamily = m_Species.COfamily
End Property

Public Property Let UTfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.UTfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get UTfamily() As String
    UTfamily = m_Species.UTfamily
End Property

Public Property Let WYfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.WYfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get WYfamily() As String
    WYfamily = m_Species.WYfamily
End Property

Public Property Let COspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.COspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get COspecies() As String
    COspecies = m_Species.COspecies
End Property

Public Property Let UTspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.UTspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get UTspecies() As String
    UTspecies = m_Species.UTspecies
End Property

Public Property Let WYspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.WYspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get WYspecies() As String
    WYspecies = m_Species.WYspecies
End Property

Public Property Let LUcode(Value As String)
    'valid length varchar(25) but 6-letter lookup
    If Not IsNull(Value) And IsBetween(Len(Value), 1, 6, True) Then
        m_Species.LUcode = Value
    Else
        RaiseEvent InvalidLUCode(Value)
    End If
End Property

Public Property Get LUcode() As String
    LUcode = m_Species.LUcode
End Property

Public Property Let MasterFamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.MasterFamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterFamily() As String
    MasterFamily = m_Species.MasterFamily
End Property

Public Property Let MasterCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.MasterCode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCode() As String
    MasterCode = m_Species.MasterCode
End Property

Public Property Let MasterSpecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.MasterSpecies = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterSpecies() As String
    MasterSpecies = m_Species.MasterSpecies
End Property

Public Property Let UTcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.UTcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get UTcode() As String
    UTcode = m_Species.UTcode
End Property

Public Property Let COcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.COcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get COcode() As String
    COcode = m_Species.COcode
End Property

Public Property Let WYcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_Species.WYcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get WYcode() As String
    WYcode = m_Species.WYcode
End Property

Public Property Let MasterCommonName(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_Species.MasterCommonName = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCommonName() As String
    MasterCommonName = m_Species.MasterCommonName
End Property

Public Property Let Lifeform(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Species.Lifeform = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Lifeform() As String
    Lifeform = m_Species.Lifeform
End Property

Public Property Let Duration(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Species.Duration = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Duration() As String
    Duration = m_Species.Duration
End Property

Public Property Let Nativity(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Species.Nativity = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Nativity() As String
    Nativity = m_Species.Nativity
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

    MsgBox "Initializing...", vbOKOnly
    
    Set m_Species = New Species

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[cls_CoverSpecies])"
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
    
    MsgBox "Terminating...", vbOKOnly
        
    Set m_Species = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_CoverSpecies])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup cover species based on the lookup code
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
    
            m_Species.Init (LUcode)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[cls_CoverSpecies])"
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb()
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As dao.Database
    Dim rs As dao.Recordset
    
    Set db = CurrentDb
    
    'record actions must have:
    strSQL = "INSERT INTO CoverSpecies(VegPlot_ID, Master_PLANT_Code, PercentCover, IsSeedling) VALUES " _
                & "('" & Me.VegPlotID & ",'" & Me.MasterPlantCode & "'," _
                & Me.PercentCover & "," & Me & "', Now() );"

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[cls_CoverSpecies])"
    End Select
    Resume Exit_Handler
End Sub