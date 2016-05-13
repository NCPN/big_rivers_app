Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    ScrollBars =0
    BorderStyle =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =1
    GridX =24
    GridY =24
    Width =11521
    DatasheetFontHeight =11
    ItemSuffix =151
    Right =16392
    Bottom =7248
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="Percent Cover"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006d01000000000000012d00006801000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin PageHeader
            Height =2364
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Top =360
                    Width =11520
                    Height =288
                    FontSize =8
                    TopMargin =29
                    BackColor =6842733
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblKey"
                    Caption ="Key"
                    GridlineColor =10921638
                    LayoutCachedTop =360
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =648
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Width =11520
                    Height =360
                    FontSize =9
                    FontWeight =700
                    TopMargin =29
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Woody Canopy % Cover"
                    GridlineColor =10921638
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Top =1920
                    Width =11520
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCheckboxRow"
                    Caption ="No Canopy Veg?"
                    GridlineColor =10921638
                    LayoutCachedTop =1920
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =2280
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Top =768
                    Width =11520
                    Height =288
                    FontSize =8
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblTotalCover"
                    Caption ="Total Plot Cover %"
                    GridlineColor =10921638
                    LayoutCachedTop =768
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1056
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Top =1188
                    Width =5760
                    Height =288
                    FontSize =7
                    LeftMargin =288
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblLeftKey"
                    Caption ="left key"
                    GridlineColor =10921638
                    LayoutCachedTop =1188
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1476
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =3
                    Left =5760
                    Top =1188
                    Width =5760
                    Height =288
                    FontSize =7
                    TopMargin =29
                    RightMargin =288
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblRightKey"
                    Caption ="right key"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =1188
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1476
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Top =1560
                    Width =11520
                    Height =288
                    FontSize =8
                    TopMargin =29
                    BackColor =11265523
                    BorderColor =8355711
                    Name ="lblSubTitle"
                    Caption ="Key"
                    GridlineColor =10921638
                    LayoutCachedTop =1560
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1848
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =5760
                    Top =1560
                    Width =5760
                    Height =288
                    FontSize =8
                    TopMargin =29
                    RightMargin =288
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblRightKeySub"
                    Caption ="right key"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =1560
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =1848
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =14869733
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Width =11520
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblRow"
                    Caption ="row"
                    GridlineColor =10921638
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =24
                    Width =960
                    Height =300
                    FontSize =8
                    TopMargin =29
                    BorderColor =10921638
                    Name ="tbxSpecies"
                    ControlSource ="Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =24
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =324
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3276
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol1"
                    Caption ="c1"
                    GridlineColor =10921638
                    LayoutCachedLeft =3276
                    LayoutCachedWidth =3794
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1248
                    Top =12
                    Width =960
                    Height =300
                    FontSize =8
                    TabIndex =1
                    TopMargin =29
                    BorderColor =10921638
                    Name ="lblCode"
                    ControlSource ="LU_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =1248
                    LayoutCachedTop =12
                    LayoutCachedWidth =2208
                    LayoutCachedHeight =312
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3792
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol2"
                    Caption ="c2"
                    GridlineColor =10921638
                    LayoutCachedLeft =3792
                    LayoutCachedWidth =4310
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4308
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol3"
                    Caption ="c3"
                    GridlineColor =10921638
                    LayoutCachedLeft =4308
                    LayoutCachedWidth =4826
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4824
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol4"
                    Caption ="c4"
                    GridlineColor =10921638
                    LayoutCachedLeft =4824
                    LayoutCachedWidth =5342
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5340
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol5"
                    Caption ="c5"
                    GridlineColor =10921638
                    LayoutCachedLeft =5340
                    LayoutCachedWidth =5858
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5856
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol6"
                    Caption ="c6"
                    GridlineColor =10921638
                    LayoutCachedLeft =5856
                    LayoutCachedWidth =6374
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6372
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol7"
                    Caption ="c7"
                    GridlineColor =10921638
                    LayoutCachedLeft =6372
                    LayoutCachedWidth =6890
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6888
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol8"
                    Caption ="c8"
                    GridlineColor =10921638
                    LayoutCachedLeft =6888
                    LayoutCachedWidth =7406
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7380
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol9"
                    Caption ="c9"
                    GridlineColor =10921638
                    LayoutCachedLeft =7380
                    LayoutCachedWidth =7898
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7896
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol10"
                    Caption ="c10"
                    GridlineColor =10921638
                    LayoutCachedLeft =7896
                    LayoutCachedWidth =8414
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8413
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol11"
                    Caption ="c11"
                    GridlineColor =10921638
                    LayoutCachedLeft =8413
                    LayoutCachedWidth =8931
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8929
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol12"
                    Caption ="c12"
                    GridlineColor =10921638
                    LayoutCachedLeft =8929
                    LayoutCachedWidth =9447
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9446
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol13"
                    Caption ="c13"
                    GridlineColor =10921638
                    LayoutCachedLeft =9446
                    LayoutCachedWidth =9964
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9964
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol14"
                    Caption ="c14"
                    GridlineColor =10921638
                    LayoutCachedLeft =9964
                    LayoutCachedWidth =10482
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10483
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol15"
                    Caption ="c15"
                    GridlineColor =10921638
                    LayoutCachedLeft =10483
                    LayoutCachedWidth =11001
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =11003
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCol16"
                    Caption ="c16"
                    GridlineColor =10921638
                    LayoutCachedLeft =11003
                    LayoutCachedWidth =11521
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageFooter
            Height =1020
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Width =11520
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblFA"
                    Caption ="Filamentous Algae"
                    GridlineColor =10921638
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Top =360
                    Width =11520
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblSocialTrails"
                    Caption ="SocialTrails"
                    GridlineColor =10921638
                    LayoutCachedTop =360
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =720
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =1
                    Top =60
                    Width =2880
                    Height =288
                    FontSize =6
                    LeftMargin =72
                    TopMargin =130
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblFAKey"
                    Caption ="FA key"
                    GridlineColor =10921638
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =348
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
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

' =================================
' Report:       PercentCover
' Level:        Application report
' Version:      1.00
'
' Description:  PercentCover report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 12, 2016
' References:
'  Allen Browne, April 2010
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 5/12/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Dim m_Park As String
Dim m_CoverType As String
Dim m_Title As String
Dim m_ShowKey As Boolean
Dim m_ShowSubHeader As Boolean
Dim m_ShowCheckboxes As Boolean
Dim m_ShowTotalPct As Boolean
Dim m_ShowFA As Boolean
Dim m_ShowSocialTrails As Boolean

'---------------------
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let Park(Value As String)
    If Len(Value) = 4 Then
        m_Park = Value
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let CoverType(Value As String)
    If Len(Value) > 0 Then
        m_CoverType = Value
    End If
End Property

Public Property Get CoverType() As String
    CoverType = m_CoverType
End Property

Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let ShowKey(Value As Boolean)
    m_ShowKey = Value
End Property

Public Property Get ShowKey() As Boolean
    ShowKey = m_ShowKey
End Property

Public Property Let ShowSubHeader(Value As Boolean)
    m_ShowSubHeader = Value
End Property

Public Property Get ShowSubHeader() As Boolean
    ShowSubHeader = m_ShowSubHeader
End Property

Public Property Let ShowCheckboxes(Value As Boolean)
    m_ShowCheckboxes = Value
End Property

Public Property Get ShowCheckboxes() As Boolean
    ShowCheckboxes = m_ShowCheckboxes
End Property

Public Property Let ShowTotalPct(Value As Boolean)
    m_ShowTotalPct = Value
End Property

Public Property Get ShowTotalPct() As Boolean
    ShowTotalPct = m_ShowTotalPct
End Property

Public Property Let ShowFA(Value As Boolean)
    m_ShowFA = Value
End Property

Public Property Get ShowFA() As Boolean
    ShowFA = m_ShowFA
End Property

Public Property Let ShowSocialTrails(Value As Boolean)
    m_ShowSocialTrails = Value
End Property

Public Property Get ShowSocialTrails() As Boolean
    ShowSocialTrails = m_ShowSocialTrails
End Property

'---------------------
' Events
'---------------------
' ---------------------------------
' Sub:          Report_Open
' Description:  Report opening event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/4/2016 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim ary() As String, strPark As String
    Dim strCoverType As String ', strWCC As String, strARC As String, strURC As String
    Dim strLabel As String
    Dim i As Integer
    
    'defaults
    
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        strPark = Nz(TempVars("ParkCode"), "")
    Else
        ary = Split(OpenArgs, "|")
        strPark = UCase(ary(0))
    End If
    
    'customizations, if any
    Select Case strPark
        Case "BLCA", ""
            Select Case Me.CoverType
                Case "WCC"
                    strCoverType = "Woody Canopy"
                Case "URS"
                    strCoverType = "Understory Rooted"
            End Select
        Case "CANY"
            Select Case Me.CoverType
                Case "WCC"
                    strCoverType = "Woody Canopy"
                Case "URS"
                    strCoverType = "Understory Rooted"
            End Select
        Case "DINO"
            'Cover type = ARC
            strCoverType = "All Rooted Species"
    End Select
    
    'header
    lblTitle.Caption = strCoverType & "% Cover"
    lblLeftKey.Caption = "R = rooted in plot"
    lblRightKey.Caption = "Rooted && Unrooted > 1.5m " _
                            & ChrW(uBullet) & " nearest 1% " _
                            & ChrW(uBullet) & " T " _
                            & ChrW(uLessThanOrEqual) & " 0.5"
    
    lblKey.Caption = ChrW(&H2264) & " 1.5m height  " _
                    & ChrW(uBullet) & "  to nearest 1%  " _
                    & ChrW(uBullet) & "  T(trace) " _
                    & ChrW(uLessThanOrEqual) & " 0.5  " _
                    & ChrW(uBullet) & "  No dead plants/parts  " _
                    & ChrW(uBullet) & "  No double-counting overlapping areas of cover  " _
                    & ChrW(uBullet) & "  max overall plot cover = 100%"
    
    lblSubTitle.Caption = "Herbaceous Indicator Species"
    
    lblFAKey.Caption = "incl. attached macrophytes & FA < 0.5cm long"
    
    'columns
    For i = 1 To 16
        strLabel = "lblCol" & i
        Me.Controls(strLabel).Caption = ""
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[PercentCover Report])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Function:     Detail_Format
' Description:  report detail formatting actions
' Assumptions:  -
' Parameters:   Cancel - if format action should be cancelled (integer)
'               FormatCount - items to format (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 12, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/12/2016 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[PercentCover Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     Detail_Print
' Description:  report detail printing actions
' Assumptions:  -
' Parameters:   Cancel - if print action should be cancelled (integer)
'               PrintCount - items to print (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 12, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/12/2016 - initial version
' ---------------------------------
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Print[PercentCover Report])"
    End Select
    Resume Exit_Handler
End Sub
