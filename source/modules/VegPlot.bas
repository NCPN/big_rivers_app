Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegPlot
' Level:        Framework class
' Version:      1.02
'
' Description:  VegPlot object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/8/2016   - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               BLC - 1/12/2017 - 1.02 - added % litter, woody debris (all parks)
'                                        replaced FilamentousAlgae (boolean) w/
'                                        PctFilamentousAlgae (%)
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_EventID As Long
Private m_SiteID As Long
Private m_FeatureID As Long
Private m_VegTransectID As Long
Private m_PlotNumber As Integer
Private m_PlotDistance As Integer
Private m_ModalSedimentSize As String '3
Private m_PercentFines As Integer
Private m_PercentWater As Integer
Private m_UnderstoryRootedPctCover As Integer
Private m_PctFilamentousAlgae As Integer
Private m_PercentLitter As Integer
Private m_PercentWoodyDebris As Integer
Private m_PlotDensity As Integer
Private m_NoCanopyVeg As Boolean
Private m_NoRootedVeg As Boolean
Private m_HasSocialTrail As Boolean
Private m_NoIndicatorSpecies As Boolean
'---------------------
' Events
'---------------------
Public Event InvalidSizeClass(Value As String)
Public Event InvalidPlotDensity(Value As Integer)
Public Event InvalidPercent(Value As Integer)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let EventID(Value As Long)
    m_EventID = Value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let SiteID(Value As Long)
    m_SiteID = Value
End Property

Public Property Get SiteID() As Long
    SiteID = m_SiteID
End Property

Public Property Let FeatureID(Value As Long)
    m_FeatureID = Value
End Property

Public Property Get FeatureID() As Long
    FeatureID = m_FeatureID
End Property

Public Property Let VegTransectID(Value As Long)
    m_VegTransectID = Value
End Property

Public Property Get VegTransectID() As Long
    VegTransectID = m_VegTransectID
End Property

Public Property Let PlotNumber(Value As Integer)
    m_PlotNumber = Value
End Property

Public Property Get PlotNumber() As Integer
    PlotNumber = m_PlotNumber
End Property

Public Property Let PlotDistance(Value As Integer)
    m_PlotDistance = Value
End Property

Public Property Get PlotDistance() As Integer
    PlotDistance = m_PlotDistance
End Property

Public Property Let ModalSedimentSize(Value As String)
    'determine if valid ModWentworthClassSize
    Dim i As Integer
    For i = ModWentworthClassSize.[_First] To ModWentworthClassSize.[_Last]
'        If (ModWentworthClassSize(i) = Value) Then
            m_ModalSedimentSize = Value
'            Exit For
'        End If
    Next
    'catch invalid values
    If Len(m_ModalSedimentSize) = 0 Then RaiseEvent InvalidSizeClass(Value)
End Property

Public Property Get ModalSedimentSize() As String
    ModalSedimentSize = m_ModalSedimentSize
End Property

Public Property Let PercentFines(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_PercentFines = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PercentFines() As Integer
    PercentFines = m_PercentFines
End Property

Public Property Let PercentWater(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_PercentWater = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PercentWater() As Integer
    PercentWater = m_PercentWater
End Property

Public Property Let UnderstoryRootedPctCover(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_UnderstoryRootedPctCover = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get UnderstoryRootedPctCover() As Integer
    UnderstoryRootedPctCover = m_UnderstoryRootedPctCover
End Property

Public Property Let PctFilamentousAlgae(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_PctFilamentousAlgae = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctFilamentousAlgae() As Integer
    PctFilamentousAlgae = m_PctFilamentousAlgae
End Property

Public Property Let PercentLitter(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_PercentLitter = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PercentLitter() As Integer
    PercentLitter = m_PercentLitter
End Property


Public Property Let PercentWoodyDebris(Value As Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_PercentWoodyDebris = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PercentWoodyDebris() As Integer
    PercentWoodyDebris = m_PercentWoodyDebris
End Property

Public Property Let PlotDensity(Value As Integer)
    Dim aryDensity() As String
    aryDensity = Split(PLOT_DENSITIES, ",")
    If IsInArray(CStr(Value), aryDensity) Then
        m_PlotDensity = CInt(Value)
    Else
        RaiseEvent InvalidPlotDensity(Value)
    End If
End Property

Public Property Get PlotDensity() As Integer
    PlotDensity = m_PlotDensity
End Property

Public Property Let NoCanopyVeg(Value As Boolean)
    m_NoCanopyVeg = Value
End Property

Public Property Get NoCanopyVeg() As Boolean
    NoCanopyVeg = m_NoCanopyVeg
End Property

Public Property Let NoRootedVeg(Value As Boolean)
    m_NoRootedVeg = Value
End Property

Public Property Get NoRootedVeg() As Boolean
    NoRootedVeg = m_NoRootedVeg
End Property

Public Property Let HasSocialTrail(Value As Boolean)
    m_HasSocialTrail = Value
End Property

Public Property Get HasSocialTrail() As Boolean
    HasSocialTrail = m_HasSocialTrail
End Property

Public Property Let NoIndicatorSpecies(Value As Boolean)
    m_NoIndicatorSpecies = Value
End Property

Public Property Get NoIndicatorSpecies() As Boolean
    NoIndicatorSpecies = m_NoIndicatorSpecies
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[cls_VegPlot])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[cls_VegPlot])"
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
'   BLC, 1/12/2017 - added % litter, woody debris
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'
'    Set db = CurrentDb
'
'    'record VegPlots must have:
'    strSQL = "INSERT INTO VegPlot(Event_ID, Site_ID, Feature_ID, " _
'                & "VegTransect_ID, PlotNumber, PlotDistance_m, " _
'                & "ModalSedimentSize, PercentFine, PercentWater, " _
'                & "UnderstoryRootedPctCover, PlotDensity, NoCanopyVeg, " _
'                & "NoRootedVeg, HasSocialTrail, FilamentousAlgae, " _
'                & "NoIndicatorSpecies) VALUES " _
'                & "(" & Me.EventID & "," & Me.SiteID & "," _
'                & Me.FeatureID & "," & Me.VegTransectID & "," _
'                & Me.PlotNumber & "," & Me.PlotDistance & ",'" _
'                & Me.ModalSedimentSize & "'," & Me.PercentFines & "," _
'                & Me.PercentWater & "," & Me.UnderstoryRootedPctCover & "," _
'                & Me.PlotDensity & "," & Me.NoCanopyVeg & "," _
'                & Me.NoRootedVeg & "," & Me.HasSocialTrail & "," _
'                & Me.FilamentousAlgae & "," & Me.NoIndicatorSpecies & ");"
'
'    db.Execute strSQL, dbFailOnError
'    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

    Dim Template As String
    
    Template = "i_vegplot"
    
    Dim Params(0 To 19) As Variant

    With Me
        Params(0) = "VegPlot"
        Params(1) = .EventID
        Params(2) = .SiteID
        Params(3) = .FeatureID
        Params(4) = .VegTransectID
        Params(5) = .PlotNumber
        Params(6) = .PlotDistance
        Params(7) = .ModalSedimentSize
        Params(8) = .PercentFines
        Params(9) = .PercentWater
        Params(10) = .UnderstoryRootedPctCover
        Params(11) = .PctFilamentousAlgae
        Params(12) = .PercentLitter
        Params(13) = .PercentWoodyDebris
        Params(14) = .PlotDensity
        Params(15) = .NoCanopyVeg
        Params(16) = .NoRootedVeg
        Params(17) = .HasSocialTrail
        Params(18) = .NoIndicatorSpecies
        
        If IsUpdate Then
            Template = "u_vegplot"
            Params(19) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_VegPlot])"
    End Select
    Resume Exit_Handler
End Sub