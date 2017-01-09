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
    Width =5520
    DatasheetFontHeight =11
    ItemSuffix =103
    Right =10260
    Bottom =5655
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="Photo Key"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006d01000000000000901500003804000001000000 ,
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
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=1"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1379
            Name ="GroupHeader0"
            AlternateBackColor =16777215
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblKeyHdr"
                    Caption ="KEY"
                    GridlineColor =10921638
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1104
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTransectHdr"
                    Caption ="T"
                    GridlineColor =10921638
                    LayoutCachedLeft =1104
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2208
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFeatureHdr"
                    Caption ="F"
                    GridlineColor =10921638
                    LayoutCachedLeft =2208
                    LayoutCachedWidth =3288
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3312
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblOverviewHdr"
                    Caption ="O"
                    GridlineColor =10921638
                    LayoutCachedLeft =3312
                    LayoutCachedWidth =4392
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Top =384
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPhotoType"
                    Caption ="Photo Type"
                    GridlineColor =10921638
                    LayoutCachedTop =384
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =744
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Top =768
                    Width =1080
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDirFacing"
                    Caption ="Direction Facing"
                    GridlineColor =10921638
                    LayoutCachedTop =768
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =1344
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4416
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRefHdr"
                    Caption ="R"
                    GridlineColor =10921638
                    LayoutCachedLeft =4416
                    LayoutCachedWidth =5496
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1104
                    Top =384
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTransect"
                    Caption ="Transect"
                    GridlineColor =10921638
                    LayoutCachedLeft =1104
                    LayoutCachedTop =384
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =744
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2208
                    Top =384
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFeature"
                    Caption ="Feature"
                    GridlineColor =10921638
                    LayoutCachedLeft =2208
                    LayoutCachedTop =384
                    LayoutCachedWidth =3288
                    LayoutCachedHeight =744
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3312
                    Top =384
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblOverview"
                    Caption ="Overview"
                    GridlineColor =10921638
                    LayoutCachedLeft =3312
                    LayoutCachedTop =384
                    LayoutCachedWidth =4392
                    LayoutCachedHeight =744
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4416
                    Top =384
                    Width =1080
                    Height =360
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReference"
                    Caption ="Reference"
                    GridlineColor =10921638
                    LayoutCachedLeft =4416
                    LayoutCachedTop =384
                    LayoutCachedWidth =5496
                    LayoutCachedHeight =744
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1104
                    Top =768
                    Width =1080
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTransectDirFacing"
                    Caption ="RR | RL"
                    GridlineColor =10921638
                    LayoutCachedLeft =1104
                    LayoutCachedTop =768
                    LayoutCachedWidth =2184
                    LayoutCachedHeight =1344
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2208
                    Top =768
                    Width =1080
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFeatureDirFacing"
                    Caption ="US | DS"
                    GridlineColor =10921638
                    LayoutCachedLeft =2208
                    LayoutCachedTop =768
                    LayoutCachedWidth =3288
                    LayoutCachedHeight =1344
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3312
                    Top =768
                    Width =1080
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblOverviewDirFacing"
                    Caption =" US | DS RR | RL"
                    GridlineColor =10921638
                    LayoutCachedLeft =3312
                    LayoutCachedTop =768
                    LayoutCachedWidth =4392
                    LayoutCachedHeight =1344
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =4416
                    Top =768
                    Width =1080
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRefDirFacing"
                    Caption =" US | DS RR | RL"
                    GridlineColor =10921638
                    LayoutCachedLeft =4416
                    LayoutCachedTop =768
                    LayoutCachedWidth =5496
                    LayoutCachedHeight =1344
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1080
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =12632256
            Begin
                Begin Rectangle
                    Width =5520
                    Height =1080
                    BackColor =0
                    BorderColor =10921638
                    Name ="rctPhotoNaming"
                    GridlineColor =10921638
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1080
                    BackThemeColorIndex =-1
                End
                Begin Rectangle
                    Left =1920
                    Top =60
                    Width =3540
                    Height =960
                    BorderColor =10921638
                    Name ="rctInset"
                    GridlineColor =10921638
                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =1020
                    BackThemeColorIndex =-1
                End
                Begin Line
                    Left =588
                    Top =504
                    Width =288
                    BorderColor =16777215
                    Name ="lnDay"
                    GridlineColor =10921638
                    LayoutCachedLeft =588
                    LayoutCachedTop =504
                    LayoutCachedWidth =876
                    LayoutCachedHeight =504
                    BorderThemeColorIndex =-1
                End
                Begin Line
                    Left =996
                    Top =240
                    Width =720
                    BorderColor =16777215
                    Name ="lnCameraNum"
                    GridlineColor =10921638
                    LayoutCachedLeft =996
                    LayoutCachedTop =240
                    LayoutCachedWidth =1716
                    LayoutCachedHeight =240
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    Left =120
                    Top =240
                    Width =1728
                    Height =312
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblExampleNum"
                    Caption ="P A 0 1 0 3 0 0"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =240
                    LayoutCachedWidth =1848
                    LayoutCachedHeight =552
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2520
                    Top =420
                    Width =564
                    Height =228
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblNamePartDay"
                    Caption ="01-31"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =420
                    LayoutCachedWidth =3084
                    LayoutCachedHeight =648
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2520
                    Top =720
                    Width =2832
                    Height =228
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblNamePartCameraNum"
                    Caption ="4-digit Camera Photo Sequence #"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =720
                    LayoutCachedWidth =5352
                    LayoutCachedHeight =948
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2520
                    Top =120
                    Width =2706
                    Height =228
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblNamePartMonth"
                    Caption ="Jan-Sep = 1-9, Oct-Dec = A-C"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =120
                    LayoutCachedWidth =5226
                    LayoutCachedHeight =348
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2160
                    Top =120
                    Width =288
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    Name ="lblEx1"
                    Caption ="1"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =120
                    LayoutCachedWidth =2448
                    LayoutCachedHeight =372
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2160
                    Top =420
                    Width =288
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    Name ="lblEx2"
                    Caption ="2"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =420
                    LayoutCachedWidth =2448
                    LayoutCachedHeight =672
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2160
                    Top =660
                    Width =288
                    Height =252
                    FontSize =9
                    BorderColor =8355711
                    Name ="lblEx3"
                    Caption ="3"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedTop =660
                    LayoutCachedWidth =2448
                    LayoutCachedHeight =912
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =1260
                    Top =600
                    Width =288
                    Height =312
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumEx3"
                    Caption ="3"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =600
                    LayoutCachedWidth =1548
                    LayoutCachedHeight =912
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =612
                    Top =600
                    Width =288
                    Height =312
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumEx2"
                    Caption ="2"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =612
                    LayoutCachedTop =600
                    LayoutCachedWidth =900
                    LayoutCachedHeight =912
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =252
                    Top =600
                    Width =288
                    Height =312
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumEx1"
                    Caption ="1"
                    FontName ="Verdana"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =252
                    LayoutCachedTop =600
                    LayoutCachedWidth =540
                    LayoutCachedHeight =912
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' Report:       PhotoKey
' Level:        Application report
' Version:      1.00
'
' Description:  PhotoKey report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 10, 2016
' References:
'  Allen Browne, April 2010
'  http://allenbrowne.com/ser-43.html
' Revisions:    BLC - 5/10/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Dim m_Park As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidPark(Park As String)

'---------------------
' Properties
'---------------------
Public Property Let Park(value As String)
    If Len(value) = 4 Then
        m_Park = value
    Else
        RaiseEvent InvalidPark(value)
    End If
End Property

Public Property Get Park() As String
    Park = m_Park
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
    Dim blnShow As Boolean
    Dim LeftPos1 As Double, LeftPos2 As Double, LeftPos3 As Double, LeftPos4 As Double
    Dim ShowWidth As Double
    
    'defaults
    strPark = ""
    ShowWidth = 0.75
    LeftPos1 = 0.7667 * TWIPS_PER_INCH
    LeftPos2 = 1.5333 * TWIPS_PER_INCH
    LeftPos3 = 2.3 * TWIPS_PER_INCH
    LeftPos4 = 3.0667 * TWIPS_PER_INCH
    
    lblOverviewDirFacing.Caption = "US | DS" & vbCrLf & "RR | RL"
    lblRefDirFacing.Caption = "US | DS" & vbCrLf & "RR | RL"
    
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        strPark = Nz(TempVars("ParkCode"), "")
    Else
        ary = Split(OpenArgs, "|")
        strPark = UCase(ary(0))
    End If
        
    Select Case strPark
        Case "BLCA", "CANY", ""
            blnShow = True
            ShowWidth = 0.75 * TWIPS_PER_INCH
            LeftPos3 = LeftPos3
            LeftPos4 = LeftPos4
        Case "DINO"
            blnShow = False
            ShowWidth = 2 * 0.75 * TWIPS_PER_INCH '1.5
            LeftPos4 = LeftPos3
            LeftPos3 = LeftPos1
    End Select
    
    'iterate & position controls
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        With ctrl
            Select Case Left(ctrl.Name, 6)
                Case "lblTra" ' "lblTransect"
                    .Left = LeftPos1
                    .Visible = blnShow
                Case "lblFea" ' "lblFeature"
                    .Left = LeftPos2
                    .Visible = blnShow
                Case "lblOve" ' "lblOverview"
                    .Left = LeftPos3
                    .Width = ShowWidth
                Case "lblRef" ' "lblRef"
                    .Left = LeftPos4
                    .Width = ShowWidth
            End Select
        End With
    Next
    
    'iterate
    For Each ctrl In Me.Controls
        Select Case ctrl.Name
            Case "lblNumEx1"
                ctrl.Caption = ChrW(uCircle1)
            Case "lblNumEx2"
                ctrl.Caption = ChrW(uCircle2)
            Case "lblNumEx3"
                ctrl.Caption = ChrW(uCircle3)
            Case "lblEx1"
                ctrl.Caption = ChrW(uCircleFilled1)
            Case "lblEx2"
                ctrl.Caption = ChrW(uCircleFilled2)
            Case "lblEx3"
                ctrl.Caption = ChrW(uCircleFilled3)
        End Select
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[PhotoKey Report])"
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
' Source/date:  Bonnie Campbell, May 10, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/10/2016 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    Dim ctrl As Control
    Dim strValue As String
    
    'iterate
    For Each ctrl In Me.Controls
        
        Select Case ctrl.Name
            Case "lblNumEx1"
                ctrl.Caption = ChrW(uCircle1)
            Case "lblNumEx2"
                ctrl.Caption = ChrW(uCircle2)
            Case "lblNumEx3"
                ctrl.Caption = ChrW(uCircle3)
            Case "lblEx1"
                ctrl.Caption = ChrW(uCircleFilled1)
            Case "lblEx2"
                ctrl.Caption = ChrW(uCircleFilled2)
            Case "lblEx3"
                ctrl.Caption = ChrW(uCircleFilled3)
        End Select
        
'        If Left(ctrl.Name, 8) = "lblNumEx" Then
'            strValue = "&H277" & Right(ctrl.Name, 1)
'            ctrl.Caption = ChrW(&H2460) 'ChrW(&H2771)
'
'        End If
    Next
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[PhotoKey Report])"
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
' Source/date:  Bonnie Campbell, May 10, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/10/2016 - initial version
' ---------------------------------
Private Sub Detail_Print(Cancel As Integer, PrintCount As Integer)
On Error GoTo Err_Handler

    'CircleControl Me.lblNumEx1

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Print[PhotoKey Report])"
    End Select
    Resume Exit_Handler
End Sub
