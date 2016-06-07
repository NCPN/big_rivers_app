Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    PictureType =1
    GridX =24
    GridY =24
    Width =15120
    DatasheetFontHeight =11
    ItemSuffix =142
    Right =15015
    Bottom =10500
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="Photos"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006d01000000000000103b00008c01000001000000 ,
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
        Begin Subform
            BorderLineStyle =0
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
            SortOrder = NotDefault
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
            Height =3600
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Top =600
                    Width =15120
                    Height =360
                    BackColor =8355711
                    BorderColor =10921638
                    Name ="rctUnderHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =600
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =960
                    BackThemeColorIndex =0
                    BackTint =50.0
                End
                Begin Label
                    Top =975
                    Width =1332
                    Height =639
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblRiver"
                    Caption ="River (circle):"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedTop =975
                    LayoutCachedWidth =1332
                    LayoutCachedHeight =1614
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =5055
                    Top =975
                    Width =2220
                    Height =639
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblSamplingDate"
                    Caption ="Date: ____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =5055
                    LayoutCachedTop =975
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =1614
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =1320
                    Top =975
                    Width =1620
                    Height =639
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblRiverSegments"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =975
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1614
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =180
                    Top =60
                    Width =1740
                    Height =324
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =6447974
                    Name ="lblTitle"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =384
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =6720
                    Top =60
                    Width =1620
                    Height =540
                    BorderColor =8355711
                    Name ="lblEntry"
                    Caption ="Data entered by: Date entered:"
                    GridlineColor =10921638
                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =10140
                    Top =60
                    Width =1620
                    Height =540
                    BorderColor =8355711
                    Name ="lblVerify"
                    Caption ="Data verified by: Date verified:"
                    GridlineColor =10921638
                    LayoutCachedLeft =10140
                    LayoutCachedTop =60
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =8640
                    Top =300
                    Width =1620
                    Height =300
                    BorderColor =8355711
                    Name ="lblEntryDate"
                    Caption ="____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =8640
                    LayoutCachedTop =300
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =12060
                    Top =300
                    Width =1620
                    Height =300
                    BorderColor =8355711
                    Name ="lblVerifyDate"
                    Caption ="____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =12060
                    LayoutCachedTop =300
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =13452
                    Width =1620
                    Height =420
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblPageOf"
                    Caption ="Page ____ of ____"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =13452
                    LayoutCachedWidth =15072
                    LayoutCachedHeight =420
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =8640
                    Top =630
                    Width =6435
                    Height =345
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblProtocolVersion"
                    Caption ="Big River Monitoring Protocol - SOP#7 - Version 1.01 - December 2015"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =8640
                    LayoutCachedTop =630
                    LayoutCachedWidth =15075
                    LayoutCachedHeight =975
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2940
                    Top =975
                    Width =837
                    Height =639
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblSiteID"
                    Caption ="Sentinel Site:"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =2940
                    LayoutCachedTop =975
                    LayoutCachedWidth =3777
                    LayoutCachedHeight =1614
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =3735
                    Top =975
                    Width =1380
                    Height =639
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblSite"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =3735
                    LayoutCachedTop =975
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =1614
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =2520
                    Top =60
                    Width =2310
                    Height =585
                    BorderColor =8355711
                    Name ="lblDownload"
                    Caption ="Photos downloaded by: Date downloaded:"
                    GridlineColor =10921638
                    LayoutCachedLeft =2520
                    LayoutCachedTop =60
                    LayoutCachedWidth =4830
                    LayoutCachedHeight =645
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =4860
                    Top =300
                    Width =1620
                    Height =300
                    BorderColor =8355711
                    Name ="lblDownloadDate"
                    Caption ="____/____/_____"
                    GridlineColor =10921638
                    LayoutCachedLeft =4860
                    LayoutCachedTop =300
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =7260
                    Top =960
                    Width =1437
                    Height =654
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblCameraDateTimeChk"
                    Caption ="Camera date/time checked?"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =7260
                    LayoutCachedTop =960
                    LayoutCachedWidth =8697
                    LayoutCachedHeight =1614
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Left =9240
                    Top =1080
                    Width =5820
                    Height =2159
                    Name ="PhotoKey"
                    SourceObject ="Report.PhotoKey"
                    GridlineColor =10921638

                    LayoutCachedLeft =9240
                    LayoutCachedTop =1080
                    LayoutCachedWidth =15060
                    LayoutCachedHeight =3239
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =1620
                    Width =6855
                    Height =1741
                    TabIndex =1
                    Name ="LocationKey"
                    SourceObject ="Report.LocationKey"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6975
                    LayoutCachedHeight =3361
                End
                Begin Label
                    Left =8100
                    Top =1260
                    Width =717
                    Height =414
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblChkBox1"
                    Caption ="ChkBox"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedLeft =8100
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8817
                    LayoutCachedHeight =1674
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Top =636
                    Width =2460
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMonitoring"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638
                    LayoutCachedTop =636
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =960
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2580
            Name ="GroupHeader0"
            AlternateBackColor =16777215
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1140
                    Top =585
                    Width =1140
                    Height =1983
                    FontWeight =700
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblDirFacing"
                    Caption ="Direction facing"
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedTop =585
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Top =585
                    Width =1140
                    Height =1983
                    FontWeight =700
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotoType"
                    Caption ="Photo Type"
                    GridlineColor =10921638
                    LayoutCachedTop =585
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6000
                    Top =585
                    Width =2040
                    Height =1983
                    FontWeight =700
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotog"
                    Caption ="Photographer"
                    GridlineColor =10921638
                    LayoutCachedLeft =6000
                    LayoutCachedTop =585
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =11580
                    Top =585
                    Width =3540
                    Height =1983
                    FontWeight =700
                    LeftMargin =288
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblComment"
                    Caption ="Comments"
                    GridlineColor =10921638
                    LayoutCachedLeft =11580
                    LayoutCachedTop =585
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =11580
                    Width =3540
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblALL5"
                    Caption ="ALL"
                    GridlineColor =10921638
                    LayoutCachedLeft =11580
                    LayoutCachedWidth =15120
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =10320
                    Width =1260
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblR"
                    Caption ="R"
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8040
                    Width =2280
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblALL4"
                    Caption ="ALL"
                    GridlineColor =10921638
                    LayoutCachedLeft =8040
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6000
                    Width =2040
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblALL3"
                    Caption ="ALL"
                    GridlineColor =10921638
                    LayoutCachedLeft =6000
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2280
                    Width =1620
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblALL"
                    Caption ="ALL"
                    GridlineColor =10921638
                    LayoutCachedLeft =2280
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =11640
                    Top =1785
                    Width =930
                    Height =315
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lblSkipped"
                    Caption ="Skipped?"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =11640
                    LayoutCachedTop =1785
                    LayoutCachedWidth =12570
                    LayoutCachedHeight =2100
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    Left =11640
                    Top =2145
                    Width =1740
                    Height =285
                    FontSize =10
                    BorderColor =8355711
                    Name ="lblSkippedReason"
                    Caption ="(if so, give reason)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =11640
                    LayoutCachedTop =2145
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =2430
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =6000
                    Top =1305
                    Width =2055
                    Height =285
                    FontSize =10
                    BorderColor =8355711
                    Name ="lblPhotogHint"
                    Caption ="(first initial, last name)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6000
                    LayoutCachedTop =1305
                    LayoutCachedWidth =8055
                    LayoutCachedHeight =1590
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =1380
                    Top =1545
                    Width =660
                    Height =285
                    FontSize =10
                    BorderColor =8355711
                    Name ="lblCircle"
                    Caption ="(circle)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1380
                    LayoutCachedTop =1545
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1830
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Top =585
                    Width =2100
                    Height =1983
                    FontWeight =700
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblSubjectLoc"
                    Caption ="Subject Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =3900
                    LayoutCachedTop =585
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Left =75
                    Top =1275
                    Width =975
                    Height =300
                    FontSize =10
                    BorderColor =8355711
                    Name ="lblCircleOne"
                    Caption ="(circle one)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =75
                    LayoutCachedTop =1275
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =1575
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =3900
                    Width =2100
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblOR"
                    Caption ="O    R"
                    GridlineColor =10921638
                    LayoutCachedLeft =3900
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2280
                    Top =585
                    Width =1620
                    Height =1983
                    FontWeight =700
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotogLoc"
                    Caption ="Photographer Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =2280
                    LayoutCachedTop =585
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =10950
                    Top =585
                    Width =630
                    Height =1983
                    FontWeight =700
                    LeftMargin =144
                    RightMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblReplacement"
                    Caption ="Replacement?"
                    GridlineColor =10921638
                    LayoutCachedLeft =10950
                    LayoutCachedTop =585
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8040
                    Top =585
                    Width =2280
                    Height =1983
                    FontWeight =700
                    TopMargin =288
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotoNum"
                    Caption ="Photo #"
                    GridlineColor =10921638
                    LayoutCachedLeft =8040
                    LayoutCachedTop =585
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1140
                    Width =1140
                    Height =576
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPulledDownloaded"
                    Caption ="ALL"
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Width =1140
                    Height =576
                    FontWeight =700
                    LeftMargin =58
                    RightMargin =29
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTransducer"
                    Caption ="Photo Type"
                    GridlineColor =10921638
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =576
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Left =10320
                    Top =585
                    Width =630
                    Height =1983
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblCloseUp"
                    Caption ="Close Up?"
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedTop =585
                    LayoutCachedWidth =10950
                    LayoutCachedHeight =2568
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =396
            Name ="Detail"
            AlternateBackColor =12632256
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Width =1140
                    Height =396
                    FontSize =12
                    LeftMargin =29
                    TopMargin =72
                    RightMargin =29
                    BorderColor =8355711
                    Name ="lblTFOR"
                    Caption ="TFOR"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =1140
                    Width =1140
                    Height =396
                    FontSize =8
                    LeftMargin =216
                    RightMargin =216
                    BorderColor =8355711
                    Name ="lblOptions"
                    Caption ="US  DS RR  RL"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =2280
                    Width =1620
                    Height =396
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblBlank1"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2280
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =3900
                    Width =2100
                    Height =396
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblBlank2"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3900
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =6000
                    Width =2040
                    Height =396
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblBlank3"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6000
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =8040
                    Width =2280
                    Height =396
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblBlank4"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =8040
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =13680
                    Top =60
                    Width =480
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCount"
                    ControlSource ="RecCount"
                    GridlineColor =10921638

                    LayoutCachedLeft =13680
                    LayoutCachedTop =60
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =375
                End
                Begin Label
                    OldBorderStyle =1
                    Left =11580
                    Width =3539
                    Height =396
                    BackColor =-2
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblChkBox4"
                    GridlineColor =10921638
                    LayoutCachedLeft =11580
                    LayoutCachedWidth =15119
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeTint =75.0
                End
                Begin Label
                    OldBorderStyle =1
                    Left =10950
                    Width =630
                    Height =396
                    BackColor =-2
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblChkBox3"
                    GridlineColor =10921638
                    LayoutCachedLeft =10950
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeTint =75.0
                End
                Begin Label
                    OldBorderStyle =1
                    Left =10320
                    Width =630
                    Height =396
                    BackColor =-2
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblChkBox2"
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedWidth =10950
                    LayoutCachedHeight =396
                    BackThemeColorIndex =-1
                    ForeTint =75.0
                End
            End
        End
        Begin PageFooter
            Height =1800
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =3420
                    Top =660
                    Width =3420
                    Height =540
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblOtherBlank1"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3420
                    LayoutCachedTop =660
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =6840
                    Top =660
                    Width =3240
                    Height =540
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblOtherBlank2"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6840
                    LayoutCachedTop =660
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =10080
                    Top =660
                    Width =4920
                    Height =540
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblOtherBlank3"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10080
                    LayoutCachedTop =660
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =735
                    Width =660
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblAP1"
                    Caption ="Animals Plants"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =735
                    LayoutCachedWidth =780
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =780
                    Top =735
                    Width =720
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblCS1"
                    Caption ="Cultural Scenic"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =780
                    LayoutCachedTop =735
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =1500
                    Top =735
                    Width =960
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblDW1"
                    Caption ="Disturbance Weather"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1500
                    LayoutCachedTop =735
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =2460
                    Top =735
                    Width =840
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblFO1"
                    Caption ="Field Work Other"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2460
                    LayoutCachedTop =735
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1140
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =60
                    Top =300
                    Width =3360
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =144
                    TopMargin =29
                    RightMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblCategory"
                    Caption ="Category (select one)"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =660
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3420
                    Top =300
                    Width =3420
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =144
                    TopMargin =29
                    RightMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblPhotographer"
                    Caption ="Photographer"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3420
                    LayoutCachedTop =300
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =660
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6840
                    Top =300
                    Width =3240
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =144
                    TopMargin =29
                    RightMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblOtherPhotoNumHdr"
                    Caption ="Photo #"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6840
                    LayoutCachedTop =300
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =660
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10080
                    Top =300
                    Width =4920
                    Height =360
                    FontSize =10
                    FontWeight =700
                    LeftMargin =144
                    TopMargin =29
                    RightMargin =144
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblDescription"
                    Caption ="Description"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10080
                    LayoutCachedTop =300
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =660
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Width =1503
                    Height =314
                    FontSize =10
                    FontWeight =700
                    LeftMargin =29
                    TopMargin =29
                    BackColor =12632256
                    BorderColor =8355711
                    Name ="lblOtherPhotos"
                    Caption ="OTHER PHOTOS:"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedWidth =1503
                    LayoutCachedHeight =314
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =60
                    Top =1200
                    Width =3360
                    Height =540
                    FontSize =8
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblCats2"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =1200
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =3420
                    Top =1200
                    Width =3420
                    Height =540
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblOtherBlank4"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =3420
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =6840
                    Top =1200
                    Width =3240
                    Height =540
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblOtherBlank5"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =6840
                    LayoutCachedTop =1200
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =4
                    Left =10080
                    Top =1200
                    Width =4920
                    Height =540
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblOtherBlank6"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =10080
                    LayoutCachedTop =1200
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =1740
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =1275
                    Width =660
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblAP2"
                    Caption ="Animals Plants"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =1275
                    LayoutCachedWidth =780
                    LayoutCachedHeight =1680
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =780
                    Top =1275
                    Width =720
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblCS2"
                    Caption ="Cultural Scenic"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =780
                    LayoutCachedTop =1275
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =1680
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =1500
                    Top =1275
                    Width =960
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblDW2"
                    Caption ="Disturbance Weather"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =1500
                    LayoutCachedTop =1275
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1680
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =2460
                    Top =1275
                    Width =840
                    Height =405
                    FontSize =8
                    BorderColor =8355711
                    Name ="lblFO2"
                    Caption ="Field Work Other"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =2460
                    LayoutCachedTop =1275
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1680
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =60
                    Top =660
                    Width =3360
                    Height =540
                    FontSize =8
                    LeftMargin =144
                    RightMargin =144
                    BorderColor =8355711
                    Name ="lblCats1"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =1200
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
' Report:       Photo
' Level:        Application report
' Version:      1.00
'
' Description:  Photo Report object related properties, events, functions & procedures for UI display
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
' Event Declarations
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let Park(Value As String)
    If Len(Value) = 4 Then
        m_Park = Value
    Else
        RaiseEvent InvalidPark(Value)
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
' References:
'   Jemima Chadwick, October 8, 2001
'   http://www.vbforums.com/showthread.php?109169-Converting-Array-to-recordset
' Source/date:  Bonnie Campbell, May 4, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/4/2016 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim ary() As String, strPark As String, strSegments As String
    Dim strSQL As String
    Dim arySegments() As Variant, aryProtocol() As Variant
    Dim i As Integer
    Dim sopnum() As String
    
    'exit if no park
    If Len(Nz(TempVars("ParkCode"), "")) = 0 Then GoTo Exit_Handler
    
    'defaults
    Me.Park = Nz(TempVars("ParkCode"), "")
    strPark = ""
    strSegments = ""
    i = 0
    
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        strPark = TempVars("ParkCode") '""
    Else
        ary = Split(OpenArgs, "|")
        strPark = UCase(ary(0))
    End If
        
    'set title
    Me.lblTitle.Caption = strPark & " Photos"

    'protocol version
    aryProtocol = GetProtocolVersion
    sopnum = GetSOPMetadata("Photo")
    If IsArray(sopnum) Then i = CInt(sopnum(0, 0))
    
    lblMonitoring.Caption = "NCPN " & aryProtocol(0, 0)
    lblProtocolVersion.Caption = aryProtocol(0, 0) & " - " & "SOP #" & i & " - Version " & aryProtocol(1, 0) & " - " & Format(aryProtocol(2, 0), "mmm yyyy")

    'set river segment(s)
    arySegments = GetRiverSegments(strPark)
    For i = 0 To UBound(arySegments, 2)
        strSegments = strSegments & arySegments(0, i) & Space(4)
    Next
    strSegments = Left(strSegments, Len(strSegments) - 1)
    
    Me.lblRiverSegments.Caption = strSegments
    
    'set key values
    Me!LocationKey.Report.Park = Me.Park
    Me!PhotoKey.Report.Park = Me.Park

    'prep checkboxes
    Me.lblChkBox1.Caption = Space(2) & ChrW(uCheckboxEmpty)
    Me.lblChkBox2.Caption = Space(2) & ChrW(uCheckboxEmpty)
    Me.lblChkBox3.Caption = Space(2) & ChrW(uCheckboxEmpty)
    Me.lblChkBox4.Caption = Space(2) & ChrW(uCheckboxEmpty)

    'prepare data source --> we simply want to print the detail multiple times, so create rs to handle it
    '                        rs("RecordCount") is datasource for tbxCount which invisibly generates rows

    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Photo Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Details_Print
' Description:  Detail print event actions
' Assumptions:  -
' Parameters:   Cancel - whether printing is canceled or not (boolean)
'               PrintCount - # of times to print (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/26/2016 - initial version
' ---------------------------------
Private Sub Details_Print(Cancel As Integer, PrintCount As Integer)
On Error GoTo Err_Handler



Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Details_Print[Photo Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Report_Close
' Description:  Closing event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 2, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/2/2016 - initial version
' ---------------------------------
Private Sub Report_Close()
On Error GoTo Err_Handler

    'unhide modal Main form
    Forms("Main").visible = True

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[Photo Report])"
    End Select
    Resume Exit_Handler
End Sub


'---------------------
' Methods
'---------------------
