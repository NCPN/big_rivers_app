Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'   Comintern, November 2, 2016
'   http://stackoverflow.com/questions/40386553/long-to-wide-and-duplicate-column-when-row-has-data

'Plot.cls

Private Type PlotData
    PlotID As Long
    VisitDate As Date
    LocationID As Long
    ModalSedimentSize As String
    PercentFine As Double
    PercentWater As Double
    UnderstoryRootedPctCover As Double
    PlotDensity As Integer
    NoCanopyVeg As Byte
    NoRootedVeg As Byte
    HasSocialTrail As Byte
    FilamentousAlgae As Byte
    NoIndicatorSpecies As Byte
    Litter As Double
End Type

Private this As PlotData
Private mCover As Scripting.Dictionary

Private Sub Class_Initialize()
    Set mCover = New Scripting.Dictionary
End Sub

Public Property Get PlotID() As Long
    PlotID = this.PlotID
End Property

Public Property Let PlotID(Value As Long)
    this.PlotID = Value
End Property

Public Property Get VisitDate() As Date
    VisitDate = this.VisitDate
End Property

Public Property Let VisitDate(Value As Date)
    this.VisitDate = Value
End Property

Public Property Get LocationID() As Long
    LocationID = this.LocationID
End Property

Public Property Let LocationID(Value As Long)
    this.LocationID = Value
End Property

Public Property Get ModalSedimentSize() As String
    ModalSedimentSize = this.ModalSedimentSize
End Property

Public Property Let ModalSedimentSize(Value As String)
    this.ModalSedimentSize = Value
End Property

Public Property Get PercentFine() As Double
    PercentFine = this.PercentFine
End Property

Public Property Let PercentFine(Value As Double)
    this.PercentFine = Value
End Property

Public Property Get PercentWater() As Double
    PercentWater = this.PercentWater
End Property

Public Property Let PercentWater(Value As Double)
    this.PercentWater = Value
End Property

Public Property Get UnderstoryRootedPctCover() As Double
    UnderstoryRootedPctCover = this.UnderstoryRootedPctCover
End Property

Public Property Let UnderstoryRootedPctCover(Value As Double)
    this.UnderstoryRootedPctCover = Value
End Property

Public Property Get PlotDensity() As Integer
    PlotDensity = this.PlotDensity
End Property

Public Property Let PlotDensity(Value As Integer)
    this.PlotDensity = Value
End Property

Public Property Get NoCanopyVeg() As Byte
    NoCanopyVeg = this.NoCanopyVeg
End Property

Public Property Let NoCanopyVeg(Value As Byte)
    this.NoCanopyVeg = Value
End Property

Public Property Get NoRootedVeg() As Byte
    NoRootedVeg = this.NoRootedVeg
End Property

Public Property Let NoRootedVeg(Value As Byte)
    this.NoRootedVeg = Value
End Property

Public Property Get HasSocialTrail() As Byte
    HasSocialTrail = this.HasSocialTrail
End Property

Public Property Let HasSocialTrail(Value As Byte)
    this.HasSocialTrail = Value
End Property

Public Property Get FilamentousAlgae() As Double
    FilamentousAlgae = this.FilamentousAlgae
End Property

Public Property Let FilamentousAlgae(Value As Double)
    this.FilamentousAlgae = Value
End Property

Public Property Get NoIndicatorSpecies() As Byte
    NoIndicatorSpecies = this.NoIndicatorSpecies
End Property

Public Property Let NoIndicatorSpecies(Value As Byte)
    this.NoIndicatorSpecies = Value
End Property

Public Sub AddSpeciesCover(species As String, cover As String)
    mCover.Add species, cover
End Sub

Public Property Get Litter() As Double
    Litter = this.Litter
End Property

Public Property Let Litter(Value As Double)
    this.Litter = Value
End Property

'Also in Plot.cls
Public Property Get CsvRows() As String
    Dim key As Variant
    Dim output() As String
    ReDim output(mCover.Count - 1)
    Dim i As Long
    For Each key In mCover.Keys
        Dim temp(16) As String
        temp(0) = this.PlotID
        temp(1) = this.VisitDate
        temp(2) = this.LocationID
        temp(3) = this.ModalSedimentSize
        temp(4) = this.PercentFine
        temp(5) = this.PercentWater
        temp(6) = this.UnderstoryRootedPctCover
        temp(7) = this.PlotDensity
        temp(8) = this.NoCanopyVeg
        temp(9) = this.NoRootedVeg
        temp(10) = this.HasSocialTrail
        temp(11) = this.FilamentousAlgae
        temp(12) = this.NoIndicatorSpecies
        temp(13) = this.Litter
        temp(14) = key
        temp(15) = mCover(key)
        output(i) = Join(temp, ",")
        i = i + 1
    Next key
    CsvRows = Join(output, vbCrLf)
End Property

'Public Sub SampleUsage()
'    Dim plots As New Collection
'
'    With ActiveSheet
'        Dim col As Long
'        For col = 2 To 4
'            Dim current As Plot
'            Set current = New Plot
'            current.PlotId = .Cells(1, col).Value
'            current.DataDate = .Cells(2, col).Value
'            current.Location = .Cells(3, col).Value
'            Dim r As Long
'            For r = 4 To 6
'                Dim cover As String
'                cover = .Cells(r, col).Value
'                If cover <> vbNullString Then
'                    current.AddSpeciesCover .Cells(r, 1).Value, cover
'                End If
'            Next
'            plots.Add current
'        Next
'
'    End With
'
'    For Each current In plots
'        Debug.Print current.CsvRows
'    Next
'End Sub