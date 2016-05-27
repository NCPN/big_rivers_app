Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    Left =7188
    Top =3216
    Right =16068
    Bottom =10764
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xe10d2e93daa8e440
    End
    Caption ="Title"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SaveSplitterBarPosition =0
    SplitFormSplitterBar =0
    SaveSplitterBarPosition =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =255
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Section
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal crKey As Long, _
   ByVal bAlpha As Byte, _
   ByVal dwFlags As Long) As Long
 
Private Const LWA_ALPHA     As Long = &H2

Private Const GWL_EXSTYLE   As Long = -20

Private Const WS_EX_LAYERED As Long = &H80000
 
Public Sub SetFormOpacity(frm As Form, sngOpacity As Single, TColor As Long)
    Dim lngStyle As Long

    ' get the current window style, then set transparency
    lngStyle = GetWindowLong(frm.hWnd, GWL_EXSTYLE)
    SetWindowLong frm.hWnd, GWL_EXSTYLE, lngStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hWnd, TColor, (sngOpacity * 255), LWA_ALPHA
End Sub

Private Sub Form_Load()
Dim Transp As Long
Transp = RGB(0, 0, 0) 'This is the color you want your background to be
 
Me.Detail.backColor = Transp
 
Me.Painting = False
SetFormOpacity Me, 0.5, Transp
Me.Painting = True
End Sub
