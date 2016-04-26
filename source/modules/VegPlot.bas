Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegPlot
' Level:        Framework class
' Version:      1.00
'
' Description:  VegPlot object related properties, events, functions & procedures
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
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
Private m_PlotDensity As Integer
Private m_NoCanopyVeg As Boolean
Private m_NoRootedVeg As Boolean
Private m_HasSocialTrail As Boolean
Private m_FilamentousAlgae As Boolean
Private m_NoIndicatorSpecies As Boolean

'Private m_SiteID As Long
'Private m_RiverID As Long
'Private m_LocationID As Long
'Private m_Observer As String
'Private m_Recorder As String
'Private m_ObserverID As Integer
'Private m_RecorderID As Integer
'Private m_CommentID As Integer
'Private m_Comment As String

'---------------------
' Events
'---------------------
Public Event InvalidSizeClass(value As String)
Public Event InvalidPlotDensity(value As Integer)
Public Event InvalidPercent(value As Integer)

'---------------------
' Properties
'---------------------
Public Property Let ID(value As Long)
    m_ID = value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let EventID(value As Long)
    m_EventID = value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let SiteID(value As Long)
    m_SiteID = value
End Property

Public Property Get SiteID() As Long
    SiteID = m_SiteID
End Property

Public Property Let FeatureID(value As Long)
    m_FeatureID = value
End Property

Public Property Get FeatureID() As Long
    FeatureID = m_FeatureID
End Property

Public Property Let VegTransectID(value As Long)
    m_VegTransectID = value
End Property

Public Property Get VegTransectID() As Long
    VegTransectID = m_VegTransectID
End Property

Public Property Let PlotNumber(value As Integer)
    m_PlotNumber = value
End Property

Public Property Get PlotNumber() As Integer
    PlotNumber = m_PlotNumber
End Property

Public Property Let PlotDistance(value As Integer)
    m_PlotDistance = value
End Property

Public Property Get PlotDistance() As Integer
    PlotDistance = m_PlotDistance
End Property

Public Property Let ModalSedimentSize(value As String)
    'determine if valid ModWentworthClassSize
    Dim i As Integer
    For i = ModWentworthClassSize.[_First] To ModWentworthClassSize.[_Last]
'        If (ModWentworthClassSize(i) = Value) Then
            m_ModalSedimentSize = value
'            Exit For
'        End If
    Next
    'catch invalid values
    If Len(m_ModalSedimentSize) = 0 Then RaiseEvent InvalidSizeClass(value)
End Property

Public Property Get ModalSedimentSize() As String
    ModalSedimentSize = m_ModalSedimentSize
End Property

Public Property Let PercentFines(value As Integer)
    If IsBetween(value, 0, 100, True) Then
        m_PercentFines = value
    Else
        RaiseEvent InvalidPercent(value)
    End If
End Property

Public Property Get PercentFines() As Integer
    PercentFines = m_PercentFines
End Property

Public Property Let PercentWater(value As Integer)
    If IsBetween(value, 0, 100, True) Then
        m_PercentWater = value
    Else
        RaiseEvent InvalidPercent(value)
    End If
End Property

Public Property Get PercentWater() As Integer
    PercentWater = m_PercentWater
End Property

Public Property Let UnderstoryRootedPctCover(value As Integer)
    If IsBetween(value, 0, 100, True) Then
        m_UnderstoryRootedPctCover = value
    Else
        RaiseEvent InvalidPercent(value)
    End If
End Property

Public Property Get UnderstoryRootedPctCover() As Integer
    UnderstoryRootedPctCover = m_UnderstoryRootedPctCover
End Property

Public Property Let PlotDensity(value As Integer)
    Dim aryDensity() As Integer
    aryDensity = Split(PLOT_DENSITIES, ",")
    If IsInArray(CStr(value), aryDensity) Then
        m_PlotDensity = value
    Else
        RaiseEvent InvalidPlotDensity(value)
    End If
End Property

Public Property Get PlotDensity() As Integer
    PlotDensity = m_PlotDensity
End Property

Public Property Let NoCanopyVeg(value As Boolean)
    m_NoCanopyVeg = value
End Property

Public Property Get NoCanopyVeg() As Boolean
    NoCanopyVeg = m_NoCanopyVeg
End Property

Public Property Let NoRootedVeg(value As Boolean)
    m_NoRootedVeg = value
End Property

Public Property Get NoRootedVeg() As Boolean
    NoRootedVeg = m_NoRootedVeg
End Property

Public Property Let HasSocialTrail(value As Boolean)
    m_HasSocialTrail = value
End Property

Public Property Get HasSocialTrail() As Boolean
    HasSocialTrail = m_HasSocialTrail
End Property

Public Property Let FilamentousAlgae(value As Boolean)
    m_FilamentousAlgae = value
End Property

Public Property Get FilamentousAlgae() As Boolean
    FilamentousAlgae = m_FilamentousAlgae
End Property

Public Property Let NoIndicatorSpecies(value As Boolean)
    m_NoIndicatorSpecies = value
End Property

Public Property Get NoIndicatorSpecies() As Boolean
    NoIndicatorSpecies = m_NoIndicatorSpecies
End Property


'Public Property Let SiteID(Value As Integer)
'    m_SiteID = Value
'End Property
'
'Public Property Get SiteID() As Integer
'    SiteID = m_SiteID
'End Property
'
'Public Property Let LocationID(Value As Integer)
'    m_LocationID = Value
'End Property
'
'Public Property Get LocationID() As Integer
'    LocationID = m_LocationID
'End Property
'
'Public Property Let ObserverID(Value As Integer)
'    m_ObserverID = Value
'End Property
'
'Public Property Get ObserverID() As Integer
'    ObserverID = m_ObserverID
'End Property
'
'Public Property Let Observer(Value As String)
'    m_Observer = Value
'End Property
'
'Public Property Get Observer() As String
'    Observer = m_Observer
'End Property
'
'Public Property Let RecorderID(Value As Integer)
'    m_RecorderID = Value
'End Property
'
'Public Property Get RecorderID() As Integer
'    RecorderID = m_RecorderID
'End Property
'
'Public Property Let Recorder(Value As String)
'    m_Recorder = Value
'End Property
'
'Public Property Get Recorder() As String
'    Recorder = m_Recorder
'End Property

''---------------------
''change to comment object instead??
''---------------------
'Public Property Let CommentID(Value As Integer)
'    m_CommentID = Value
'End Property
'
'Public Property Get CommentID() As Integer
'    CommentID = m_CommentID
'End Property
'
'Public Property Let Comment(Value As String)
'    m_Comment = Value
'End Property
'
'Public Property Get Comment() As String
'    Comment = m_Comment
'End Property

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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb()
On Error GoTo Err_Handler
    
    Dim strSQL As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    'record VegPlots must have:
    strSQL = "INSERT INTO VegPlot(Event_ID, Site_ID, Feature_ID, " _
                & "VegTransect_ID, PlotNumber, PlotDistance_m, " _
                & "ModalSedimentSize, PercentFine, PercentWater, " _
                & "UnderstoryRootedPctCover, PlotDensity, NoCanopyVeg, " _
                & "NoRootedVeg, HasSocialTrail, FilamentousAlgae, " _
                & "NoIndicatorSpecies) VALUES " _
                & "(" & Me.EventID & "," & Me.SiteID & "," _
                & Me.FeatureID & "," & Me.VegTransectID & "," _
                & Me.PlotNumber & "," & Me.PlotDistance & ",'" _
                & Me.ModalSedimentSize & "'," & Me.PercentFines & "," _
                & Me.PercentWater & "," & Me.UnderstoryRootedPctCover & "," _
                & Me.PlotDensity & "," & Me.NoCanopyVeg & "," _
                & Me.NoRootedVeg & "," & Me.HasSocialTrail & "," _
                & Me.FilamentousAlgae & "," & Me.NoIndicatorSpecies & ");"

    db.Execute strSQL, dbFailOnError
    Me.ID = db.OpenRecordset("SELECT @@IDENTITY")(0)

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