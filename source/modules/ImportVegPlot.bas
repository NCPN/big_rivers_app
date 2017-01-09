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

Public Property Let PlotID(value As Long)
    this.PlotID = value
End Property

Public Property Get VisitDate() As Date
    VisitDate = this.VisitDate
End Property

Public Property Let VisitDate(value As Date)
    this.VisitDate = value
End Property

Public Property Get LocationID() As Long
    LocationID = this.LocationID
End Property

Public Property Let LocationID(value As Long)
    this.LocationID = value
End Property

Public Property Get ModalSedimentSize() As String
    ModalSedimentSize = this.ModalSedimentSize
End Property

Public Property Let ModalSedimentSize(value As String)
    this.ModalSedimentSize = value
End Property

Public Property Get PercentFine() As Double
    PercentFine = this.PercentFine
End Property

Public Property Let PercentFine(value As Double)
    this.PercentFine = value
End Property

Public Property Get PercentWater() As Double
    PercentWater = this.PercentWater
End Property

Public Property Let PercentWater(value As Double)
    this.PercentWater = value
End Property

Public Property Get UnderstoryRootedPctCover() As Double
    UnderstoryRootedPctCover = this.UnderstoryRootedPctCover
End Property

Public Property Let UnderstoryRootedPctCover(value As Double)
    this.UnderstoryRootedPctCover = value
End Property

Public Property Get PlotDensity() As Integer
    PlotDensity = this.PlotDensity
End Property

Public Property Let PlotDensity(value As Integer)
    this.PlotDensity = value
End Property

Public Property Get NoCanopyVeg() As Byte
    NoCanopyVeg = this.NoCanopyVeg
End Property

Public Property Let NoCanopyVeg(value As Byte)
    this.NoCanopyVeg = value
End Property

Public Property Get NoRootedVeg() As Byte
    NoRootedVeg = this.NoRootedVeg
End Property

Public Property Let NoRootedVeg(value As Byte)
    this.NoRootedVeg = value
End Property

Public Property Get HasSocialTrail() As Byte
    HasSocialTrail = this.HasSocialTrail
End Property

Public Property Let HasSocialTrail(value As Byte)
    this.HasSocialTrail = value
End Property

Public Property Get FilamentousAlgae() As Double
    FilamentousAlgae = this.FilamentousAlgae
End Property

Public Property Let FilamentousAlgae(value As Double)
    this.FilamentousAlgae = value
End Property

Public Property Get NoIndicatorSpecies() As Byte
    NoIndicatorSpecies = this.NoIndicatorSpecies
End Property

Public Property Let NoIndicatorSpecies(value As Byte)
    this.NoIndicatorSpecies = value
End Property

Public Sub AddSpeciesCover(species As String, cover As String)
    mCover.Add species, cover
End Sub

Public Property Get Litter() As Double
    Litter = this.Litter
End Property

Public Property Let Litter(value As Double)
    this.Litter = value
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