Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
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
    Width =11460
    DatasheetFontHeight =11
    ItemSuffix =185
    Right =25395
    Bottom =11790
    DatasheetGridlinesColor =14806254
    OnNoData ="=NoData([Report])"
    RecSrcDt = Begin
        0x32f638d6d6a9e440
    End
    Caption ="Percent Cover"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006d01000000000000c42c00006801000001000000 ,
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="=1"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=2"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=3"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=4"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="=5"
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
            Height =360
            Name ="GroupHeader0"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Width =11448
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
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Top =36
                    Width =5760
                    Height =288
                    FontSize =7
                    LeftMargin =288
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLeftKey"
                    Caption ="left key"
                    GridlineColor =10921638
                    LayoutCachedTop =36
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =324
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =5685
                    Top =30
                    Width =5760
                    Height =288
                    FontSize =7
                    TopMargin =29
                    RightMargin =288
                    BackColor =14869733
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblRightKey"
                    Caption ="right key"
                    GridlineColor =10921638
                    LayoutCachedLeft =5685
                    LayoutCachedTop =30
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =318
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            BreakLevel =1
            Name ="GroupHeader1"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Width =11448
                    Height =288
                    FontSize =8
                    TopMargin =29
                    BackColor =6842733
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblKey"
                    Caption ="Key"
                    GridlineColor =10921638
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =288
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =363
            BreakLevel =2
            Name ="GroupHeader2"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Width =11448
                    Height =360
                    FontSize =8
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblTotalCover"
                    Caption ="Total Plot Cover %"
                    GridlineColor =10921638
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3195
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT1"
                    Caption ="c1"
                    GridlineColor =10921638
                    LayoutCachedLeft =3195
                    LayoutCachedWidth =3713
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3711
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT2"
                    Caption ="c2"
                    GridlineColor =10921638
                    LayoutCachedLeft =3711
                    LayoutCachedWidth =4229
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4227
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT3"
                    Caption ="c3"
                    GridlineColor =10921638
                    LayoutCachedLeft =4227
                    LayoutCachedWidth =4745
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4743
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT4"
                    Caption ="c4"
                    GridlineColor =10921638
                    LayoutCachedLeft =4743
                    LayoutCachedWidth =5261
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5259
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT5"
                    Caption ="c5"
                    GridlineColor =10921638
                    LayoutCachedLeft =5259
                    LayoutCachedWidth =5777
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5775
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT6"
                    Caption ="c6"
                    GridlineColor =10921638
                    LayoutCachedLeft =5775
                    LayoutCachedWidth =6293
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6291
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT7"
                    Caption ="c7"
                    GridlineColor =10921638
                    LayoutCachedLeft =6291
                    LayoutCachedWidth =6809
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6807
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT8"
                    Caption ="c8"
                    GridlineColor =10921638
                    LayoutCachedLeft =6807
                    LayoutCachedWidth =7325
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7299
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT9"
                    Caption ="c9"
                    GridlineColor =10921638
                    LayoutCachedLeft =7299
                    LayoutCachedWidth =7817
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7815
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT10"
                    Caption ="c10"
                    GridlineColor =10921638
                    LayoutCachedLeft =7815
                    LayoutCachedWidth =8333
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8332
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT11"
                    Caption ="c11"
                    GridlineColor =10921638
                    LayoutCachedLeft =8332
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8848
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT12"
                    Caption ="c12"
                    GridlineColor =10921638
                    LayoutCachedLeft =8848
                    LayoutCachedWidth =9366
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9365
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT13"
                    Caption ="c13"
                    GridlineColor =10921638
                    LayoutCachedLeft =9365
                    LayoutCachedWidth =9883
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9883
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT14"
                    Caption ="c14"
                    GridlineColor =10921638
                    LayoutCachedLeft =9883
                    LayoutCachedWidth =10401
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10402
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT15"
                    Caption ="c15"
                    GridlineColor =10921638
                    LayoutCachedLeft =10402
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10922
                    Width =518
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColT16"
                    Caption ="c16"
                    GridlineColor =10921638
                    LayoutCachedLeft =10922
                    LayoutCachedWidth =11440
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =288
            BreakLevel =3
            Name ="GroupHeader3"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =2
                    Width =11448
                    Height =288
                    FontSize =8
                    TopMargin =29
                    BackColor =11265523
                    BorderColor =8355711
                    Name ="lblSubTitle"
                    Caption ="Key"
                    GridlineColor =10921638
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =288
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =3
                    Left =5685
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
                    LayoutCachedLeft =5685
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =288
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =363
            BreakLevel =4
            Name ="GroupHeader4"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Top =3
                    Width =11448
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCheckboxRow"
                    Caption ="No Canopy Veg?"
                    GridlineColor =10921638
                    LayoutCachedTop =3
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =363
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3195
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC1"
                    Caption ="c1"
                    GridlineColor =10921638
                    LayoutCachedLeft =3195
                    LayoutCachedWidth =3713
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3711
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC2"
                    Caption ="c2"
                    GridlineColor =10921638
                    LayoutCachedLeft =3711
                    LayoutCachedWidth =4229
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4227
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC3"
                    Caption ="c3"
                    GridlineColor =10921638
                    LayoutCachedLeft =4227
                    LayoutCachedWidth =4745
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4743
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC4"
                    Caption ="c4"
                    GridlineColor =10921638
                    LayoutCachedLeft =4743
                    LayoutCachedWidth =5261
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5259
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC5"
                    Caption ="c5"
                    GridlineColor =10921638
                    LayoutCachedLeft =5259
                    LayoutCachedWidth =5777
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5775
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC6"
                    Caption ="c6"
                    GridlineColor =10921638
                    LayoutCachedLeft =5775
                    LayoutCachedWidth =6293
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6291
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC7"
                    Caption ="c7"
                    GridlineColor =10921638
                    LayoutCachedLeft =6291
                    LayoutCachedWidth =6809
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6807
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC8"
                    Caption ="c8"
                    GridlineColor =10921638
                    LayoutCachedLeft =6807
                    LayoutCachedWidth =7325
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7299
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC9"
                    Caption ="c9"
                    GridlineColor =10921638
                    LayoutCachedLeft =7299
                    LayoutCachedWidth =7817
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7815
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC10"
                    Caption ="c10"
                    GridlineColor =10921638
                    LayoutCachedLeft =7815
                    LayoutCachedWidth =8333
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8332
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC11"
                    Caption ="c11"
                    GridlineColor =10921638
                    LayoutCachedLeft =8332
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8848
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC12"
                    Caption ="c12"
                    GridlineColor =10921638
                    LayoutCachedLeft =8848
                    LayoutCachedWidth =9366
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9365
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC13"
                    Caption ="c13"
                    GridlineColor =10921638
                    LayoutCachedLeft =9365
                    LayoutCachedWidth =9883
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9883
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC14"
                    Caption ="c14"
                    GridlineColor =10921638
                    LayoutCachedLeft =9883
                    LayoutCachedWidth =10401
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10402
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC15"
                    Caption ="c15"
                    GridlineColor =10921638
                    LayoutCachedLeft =10402
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10922
                    Width =518
                    Height =360
                    FontSize =14
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblColC16"
                    Caption ="c16"
                    GridlineColor =10921638
                    LayoutCachedLeft =10922
                    LayoutCachedWidth =11440
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =4
                    Top =3
                    Width =11448
                    Height =360
                    FontSize =14
                    FontWeight =600
                    LeftMargin =3384
                    TopMargin =29
                    RightMargin =144
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblCheckboxes"
                    Caption ="checkboxes"
                    GridlineColor =10921638
                    LayoutCachedTop =3
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =363
                    ThemeFontIndex =-1
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
            Height =360
            Name ="Detail"
            AutoHeight =255
            AlternateBackColor =14869733
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Width =11448
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblRow"
                    Caption ="row"
                    GridlineColor =10921638
                    LayoutCachedWidth =11448
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
                    Width =960
                    Height =360
                    FontSize =8
                    TopMargin =29
                    BorderColor =10921638
                    Name ="tbxSpecies"
                    ControlSource ="Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3195
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
                    LayoutCachedLeft =3195
                    LayoutCachedWidth =3713
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
                    Width =960
                    Height =360
                    FontSize =8
                    TabIndex =1
                    TopMargin =29
                    BorderColor =10921638
                    Name ="lblCode"
                    ControlSource ="LU_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =1248
                    LayoutCachedWidth =2208
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =3711
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
                    LayoutCachedLeft =3711
                    LayoutCachedWidth =4229
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4227
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
                    LayoutCachedLeft =4227
                    LayoutCachedWidth =4745
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =4743
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
                    LayoutCachedLeft =4743
                    LayoutCachedWidth =5261
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5259
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
                    LayoutCachedLeft =5259
                    LayoutCachedWidth =5777
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =5775
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
                    LayoutCachedLeft =5775
                    LayoutCachedWidth =6293
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6291
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
                    LayoutCachedLeft =6291
                    LayoutCachedWidth =6809
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =6807
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
                    LayoutCachedLeft =6807
                    LayoutCachedWidth =7325
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7299
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
                    LayoutCachedLeft =7299
                    LayoutCachedWidth =7817
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =7815
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
                    LayoutCachedLeft =7815
                    LayoutCachedWidth =8333
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8332
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
                    LayoutCachedLeft =8332
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =8848
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
                    LayoutCachedLeft =8848
                    LayoutCachedWidth =9366
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9365
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
                    LayoutCachedLeft =9365
                    LayoutCachedWidth =9883
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =9883
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
                    LayoutCachedLeft =9883
                    LayoutCachedWidth =10401
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10402
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
                    LayoutCachedLeft =10402
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Left =10922
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
                    LayoutCachedLeft =10922
                    LayoutCachedWidth =11440
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =720
            Name ="GroupFooter1"
            AutoHeight =255
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    TextAlign =1
                    Width =11448
                    Height =360
                    FontSize =8
                    FontWeight =600
                    TopMargin =29
                    BackColor =14869733
                    BorderColor =8355711
                    Name ="lblFA"
                    Caption ="Filamentous Algae"
                    GridlineColor =10921638
                    LayoutCachedWidth =11448
                    LayoutCachedHeight =360
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =1
                    Top =360
                    Width =11448
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
                    LayoutCachedWidth =11448
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
' Declarations
'---------------------
Dim m_Park As String
Dim m_CoverType As String
Dim m_CoverTypeName As String
Dim m_Title As String
Dim m_ShowTitleKey As Boolean
Dim m_ShowKey As Boolean
Dim m_ShowSubTitle As Boolean
Dim m_ShowCheckboxes As Boolean
Dim m_ShowTotalPct As Boolean
Dim m_ShowFA As Boolean
Dim m_ShowSocialTrails As Boolean

'---------------------
' Event Declarations
'---------------------
Public Event InvalidPark(Park As String)
Public Event InvalidCoverType(CoverType As String)

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

Public Property Let CoverType(Value As String)
    Dim aryCoverType() As String
    aryCoverType = Split(COVER_TYPES, ",")
    
    If Len(Value) > 0 And IsInArray(Value, aryCoverType) Then
        m_CoverType = Value
        Select Case UCase(Value)
            Case "WCC"
                CoverTypeName = "Woody Canopy"
                ShowTitleKey = True
                ShowKey = False
                ShowTotalPct = False
                ShowSubTitle = False
                lblCheckboxes.Caption = "No Canopy Veg?"
            Case "URC"
                CoverTypeName = "Understory Rooted Plant"
                ShowTitleKey = True
                ShowKey = True
                ShowTotalPct = True
                ShowSubTitle = True
                Select Case Park
                    Case "BLCA"
                        lblCheckboxes.Caption = "No Indicator Species?"
                        lblRightKeySub.Caption = ChrW(uLessThanOrEqual) & " 1.5m " _
                                                    & ChrW(uBullet) & "nearest 1%"
                    Case "CANY"
                        lblCheckboxes.Caption = "No Rooted Veg?"
                        lblRightKeySub.Caption = ChrW(uLessThanOrEqual) & " 1.5m " _
                                                    & ChrW(uBullet) & " % Cover "
                    Case "DINO" 'N/A
                End Select
            Case "ARS"
                CoverTypeName = "All Rooted Species"
                ShowTitleKey = False
                ShowKey = False
                ShowTotalPct = False
                ShowSubTitle = False
                lblCheckboxes.Caption = "No Rooted Veg?"
        End Select
    Else
        RaiseEvent InvalidCoverType(Value)
    End If
End Property

Public Property Get CoverType() As String
    CoverType = m_CoverType
End Property

Public Property Let CoverTypeName(Value As String)
    If Len(Value) > 0 Then
        m_CoverTypeName = Value
    End If
End Property

Public Property Get CoverTypeName() As String
    CoverTypeName = m_CoverTypeName
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
    'set display
    ToggleRow "key", Me.ShowKey
End Property

Public Property Get ShowKey() As Boolean
    ShowKey = m_ShowKey
End Property

Public Property Let ShowTitleKey(Value As Boolean)
    m_ShowTitleKey = Value
    'set display
    ToggleRow "titlekey", Me.ShowTitleKey
End Property

Public Property Get ShowTitleKey() As Boolean
    ShowTitleKey = m_ShowTitleKey
End Property

Public Property Let ShowSubTitle(Value As Boolean)
    m_ShowSubTitle = Value
    'set display
    ToggleRow "subtitle", Me.ShowSubTitle
End Property

Public Property Get ShowSubTitle() As Boolean
    ShowSubTitle = m_ShowSubTitle
End Property

Public Property Let ShowCheckboxes(Value As Boolean)
    m_ShowCheckboxes = Value
End Property

Public Property Get ShowCheckboxes() As Boolean
    ShowCheckboxes = m_ShowCheckboxes
End Property

Public Property Let ShowTotalPct(Value As Boolean)
    m_ShowTotalPct = Value
    'set display
    ToggleRow "total", Me.ShowTotalPct
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

    Dim ary() As String
    Dim strLabel As String, strCheckboxes As String, strSQL As String
    Dim i As Integer
    
    'defaults
    strCheckboxes = ""
    
    If Len(Nz(OpenArgs, "")) > 0 Or IsNull(OpenArgs) Then
        Park = Nz(TempVars("ParkCode"), "")
    Else
        ary = Split(OpenArgs, "|")
        Park = UCase(ary(0))
    End If
    
    'default
'    If Len(covertype) = 0 Then Me.covertype = "ARS"
    
    'customizations, if any
    Select Case Park
        Case "BLCA", ""
            Select Case Me.CoverType
                Case "WCC"
                    strSQL = ""
                Case "URS"
                    strSQL = ""
            End Select
        Case "CANY"
            Select Case Me.CoverType
                Case "WCC"
                    strSQL = ""
                Case "URS"
                    strSQL = ""
            End Select
        Case "DINO"
            strSQL = ""
    End Select

    'headers & keys
    lblTitle.Caption = Me.CoverTypeName & " % Cover"
    lblLeftKey.Caption = "R = rooted in plot"
    lblRightKey.Caption = "Rooted && Unrooted > 1.5m " _
                            & ChrW(uBullet) & " nearest 1% " _
                            & ChrW(uBullet) & " T " _
                            & ChrW(uLessThanOrEqual) & " 0.5"
    
    lblKey.Caption = ChrW(uLessThanOrEqual) & " 1.5m height  " _
                    & ChrW(uBullet) & "  to nearest 1%  " _
                    & ChrW(uBullet) & "  T(trace) " _
                    & ChrW(uLessThanOrEqual) & " 0.5  " _
                    & ChrW(uBullet) & "  No dead plants/parts  " _
                    & ChrW(uBullet) & "  No double-counting overlapping areas of cover  " _
                    & ChrW(uBullet) & "  max overall plot cover = 100%"
    
    lblSubTitle.Caption = "Herbaceous Indicator Species"
    
    lblFAKey.Caption = "incl. attached macrophytes & FA < 0.5cm long"
    
    'columns (total %, checkboxes & species)
    For i = 1 To 16
        strLabel = "lblColT" & i
        Me.Controls(strLabel).Caption = ""
    
        strLabel = "lblColC" & i
        Me.Controls(strLabel).Caption = ""
        
        strLabel = "lblCol" & i
        Me.Controls(strLabel).Caption = ""
        
        strCheckboxes = strCheckboxes & ChrW(uCheckboxEmpty)
    Next
    
    lblCheckboxes.Caption = strCheckboxes
    
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
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/28/2015 - initial version
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
            "Error encountered (#" & Err.Number & " - Class_Initialize[PercentCover report])"
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
' Source/date:  Bonnie Campbell, October 28, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/28/2015 - initial version
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
            "Error encountered (#" & Err.Number & " - Class_Terminate[PercentCover report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetCoverType
' Description:  Sets report cover type value
' Assumptions:  -
' Parameters:   covertype - cover type to display (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 25, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/25/2016 - initial version
' ---------------------------------
Private Sub SetCoverType(CoverType As String)
On Error GoTo Err_Handler

    Me.CoverType = CoverType
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetCoverType[PercentCover Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ToggleRow
' Description:  Shows/Hides the row for key & subtitle
' Assumptions:  Group headers contain header rows
'                   GroupHeader0 - title
'                   GroupHeader1 - title key
'                   GroupHeader2 - total % cover
'                   GroupHeader3 - subtitle
'                   GroupHeader4 - checkboxes
' Parameters:   row - row of controls to hide (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/13/2016 - initial version
' ---------------------------------
Private Sub ToggleRow(row As String, show As Boolean)
On Error GoTo Err_Handler

    Dim strLabel As String
    Dim i As Integer
    
    Select Case row
        Case "titlekey"  'key on title row
            lblLeftKey.visible = show
            lblRightKey.visible = show
        Case "key"  'key row below title
            Me.GroupHeader1.visible = show
            lblKey.visible = show
            
        Case "total"    'total plot % cover
            Me.GroupHeader2.visible = show
            lblTotalCover.visible = show
            
            'show/hide columns
            For i = 1 To 16
                strLabel = "lblColT" & i
                Me.Controls(strLabel).visible = show
            Next
                    
        Case "subtitle" 'subtitle row above checkboxes
            Me.GroupHeader3.visible = show
            lblSubTitle.visible = show
            lblRightKeySub.visible = show
            
    End Select
        
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleRow[PercentCover Report])"
    End Select
    Resume Exit_Handler
End Sub
