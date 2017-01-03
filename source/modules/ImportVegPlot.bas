Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'   Comintern, November 2, 2016
'   http://stackoverflow.com/questions/40386553/long-to-wide-and-duplicate-column-when-row-has-data

'Plot.cls

Private Type PlotMembers
    PlotId As Long
    DataDate As Date
    Location As String
End Type

Private this As PlotMembers
Private mCover As Scripting.Dictionary

Private Sub Class_Initialize()
    Set mCover = New Scripting.Dictionary
End Sub

Public Property Get PlotId() As Long
    PlotId = this.PlotId
End Property

Public Property Let PlotId(inValue As Long)
    this.PlotId = inValue
End Property

Public Property Get DataDate() As Date
    DataDate = this.DataDate
End Property

Public Property Let DataDate(inValue As Date)
    this.DataDate = inValue
End Property

Public Property Get Location() As String
    Location = this.Location
End Property

Public Property Let Location(inValue As String)
    this.Location = inValue
End Property

Public Sub AddSpeciesCover(species As String, cover As String)
    mCover.Add species, cover
End Sub

'Also in Plot.cls
Public Property Get CsvRows() As String
    Dim key As Variant
    Dim output() As String
    ReDim output(mCover.Count - 1)
    Dim i As Long
    For Each key In mCover.Keys
        Dim temp(4) As String
        temp(0) = this.PlotId
        temp(1) = this.DataDate
        temp(2) = this.Location
        temp(3) = key
        temp(4) = mCover(key)
        output(i) = Join(temp, ",")
        i = i + 1
    Next key
    CsvRows = Join(output, vbCrLf)
End Property

Public Sub SampleUsage()
    Dim plots As New Collection

    With ActiveSheet
        Dim col As Long
        For col = 2 To 4
            Dim current As Plot
            Set current = New Plot
            current.PlotId = .Cells(1, col).Value
            current.DataDate = .Cells(2, col).Value
            current.Location = .Cells(3, col).Value
            Dim r As Long
            For r = 4 To 6
                Dim cover As String
                cover = .Cells(r, col).Value
                If cover <> vbNullString Then
                    current.AddSpeciesCover .Cells(r, 1).Value, cover
                End If
            Next
            plots.Add current
        Next

    End With

    For Each current In plots
        Debug.Print current.CsvRows
    Next
End Sub