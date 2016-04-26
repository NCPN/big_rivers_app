Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Species
' Level:        Framework class
' Version:      1.00
'
' Description:  Species object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 10/30/2015
' References:   -
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_MasterPlantCode As String '20 or ZLS

Private m_COfamily As String '50 or ZLS
Private m_UTfamily As String '50 or ZLS
Private m_WYfamily As String '50 or ZLS
Private m_COspecies As String '50 or ZLS
Private m_UTspecies As String '50 or ZLS
Private m_WYspecies As String '50 or ZLS
Private m_LUcode As String 'lookup code '25 (NOT NULL!)

Private m_MasterFamily As String '50 or ZLS
Private m_MasterCode As String  '20 or ZLS
Private m_MasterSpecies As String '50 or ZLS

Private m_UTcode As String '20 or ZLS
Private m_COcode As String '20 or ZLS
Private m_WYcode As String '20 or ZLS

Private m_MasterCommonName As String '50 or ZLS

Private m_Lifeform As String '255 or ZLS
Private m_Duration As String '255 or ZLS
Private m_Nativity As String '255 or ZLS

'---------------------
' Events
'---------------------
Public Event InvalidMasterPlantCode(value As String)
Public Event InvalidLUCode(value As String)
Public Event InvalidFamily(value As String)
Public Event InvalidSpecies(value As String)
Public Event InvalidCode(value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let MasterPlantCode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_MasterPlantCode = value
    Else
        RaiseEvent InvalidMasterPlantCode(value)
    End If
End Property

Public Property Get MasterPlantCode() As String
    MasterPlantCode = m_MasterPlantCode
End Property

Public Property Let COfamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_COfamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get COfamily() As String
    COfamily = m_COfamily
End Property

Public Property Let UTfamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_UTfamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get UTfamily() As String
    UTfamily = m_UTfamily
End Property

Public Property Let WYfamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_WYfamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get WYfamily() As String
    WYfamily = m_WYfamily
End Property

Public Property Let COspecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_COspecies = value
    Else
        RaiseEvent InvalidSpecies(value)
    End If
End Property

Public Property Get COspecies() As String
    COspecies = m_COspecies
End Property

Public Property Let UTspecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_UTspecies = value
    Else
        RaiseEvent InvalidSpecies(value)
    End If
End Property

Public Property Get UTspecies() As String
    UTspecies = m_UTspecies
End Property

Public Property Let WYspecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_WYspecies = value
    Else
        RaiseEvent InvalidSpecies(value)
    End If
End Property

Public Property Get WYspecies() As String
    WYspecies = m_WYspecies
End Property

Public Property Let LUcode(value As String)
    'valid length varchar(25) but 6-letter lookup
    If Not IsNull(value) And IsBetween(Len(value), 1, 6, True) Then
        m_LUcode = value
    Else
        RaiseEvent InvalidLUCode(value)
    End If
End Property

Public Property Get LUcode() As String
    LUcode = m_LUcode
End Property

Public Property Let MasterFamily(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_MasterFamily = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get MasterFamily() As String
    MasterFamily = m_MasterFamily
End Property

Public Property Let MasterCode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_MasterCode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get MasterCode() As String
    MasterCode = m_MasterCode
End Property

Public Property Let MasterSpecies(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_MasterSpecies = value
    Else
        RaiseEvent InvalidFamily(value)
    End If
End Property

Public Property Get MasterSpecies() As String
    MasterSpecies = m_MasterSpecies
End Property

Public Property Let UTcode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_UTcode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get UTcode() As String
    UTcode = m_UTcode
End Property

Public Property Let COcode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_COcode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get COcode() As String
    COcode = m_COcode
End Property

Public Property Let WYcode(value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(value), 1, 20, True) Then
        m_WYcode = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get WYcode() As String
    WYcode = m_WYcode
End Property

Public Property Let MasterCommonName(value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(value), 1, 50, True) Then
        m_MasterCommonName = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get MasterCommonName() As String
    MasterCommonName = m_MasterCommonName
End Property

Public Property Let Lifeform(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_Lifeform = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get Lifeform() As String
    Lifeform = m_Lifeform
End Property

Public Property Let Duration(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_Duration = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get Duration() As String
    Duration = m_Duration
End Property

Public Property Let Nativity(value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(value), 1, 255, True) Then
        m_Nativity = value
    Else
        RaiseEvent InvalidCode(value)
    End If
End Property

Public Property Get Nativity() As String
    Nativity = m_Nativity
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


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[cls_Species])"
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
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_Species])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup species based on the lookup code
' Parameters:   luCode - species 6-character lookup code from NCPN master plants (string)
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/18/2016 - for NCPN tools
' Revisions:
'   BLC, 4/18/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub Init(LUcode As String)
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    'species must have:
    strSQL = "SELECT DISTINCT TOP 1 Master_Family, Master_PLANT_Code, Master_Species, " _
            & "UT_Family, CO_Family, WY_Family, Utah_PLant_Code, " _
            & "Utah_Species, CO_PLANT_Code, Co_Species, " _
            & "Wy_PLANT_Code, Wy_Species, Master_Common_Name, " _
            & "LU_Code, Lifeform, Duration, Nativity " _
            & "FROM tlu_NCPN_plants WHERE LU_Code = '" & LUcode & "';"

    Set rs = db.OpenRecordset(strSQL)
    If Not (rs.EOF And rs.BOF) Then
        With rs
            Me.MasterFamily = Nz(.Fields("Master_Family"), "")
            Me.MasterPlantCode = Nz(.Fields("Master_PLANT_Code"), "")
            Me.MasterCode = Me.MasterPlantCode
            Me.MasterSpecies = Nz(.Fields("Master_Species"), "")
            Me.UTfamily = Nz(.Fields("UT_Family"), "")
            Me.COfamily = Nz(.Fields("CO_Family"), "")
            Me.WYfamily = Nz(.Fields("WY_Family"), "")
            Me.UTcode = Nz(.Fields("Utah_Plant_Code"), "")
            Me.UTspecies = Nz(.Fields("Utah_Species"), "")
            Me.COcode = Nz(.Fields("CO_PLANT_Code"), "")
            Me.COspecies = Nz(.Fields("Co_Species"), "")
            Me.WYcode = Nz(.Fields("Wy_PLANT_code"), "")
            Me.WYspecies = Nz(.Fields("Wy_Species"), "")
            Me.MasterCommonName = Nz(.Fields("Master_Common_Name"), "")
            Me.LUcode = .Fields("LU_Code")
            Me.Lifeform = Nz(.Fields("Lifeform"), "")
            Me.Duration = Nz(.Fields("Duration"), "")
            Me.Nativity = Nz(.Fields("Nativity"), "")
        End With
    Else
        RaiseEvent InvalidLUCode(LUcode)
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[cls_Species])"
    End Select
    Resume Exit_Handler
End Sub