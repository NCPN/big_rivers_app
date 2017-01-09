Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    SubdatasheetExpanded = NotDefault
    ScrollBars =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13584
    DatasheetFontHeight =11
    ItemSuffix =64
    Left =3120
    Top =3465
    Right =16965
    Bottom =9645
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x0a915c95ff94e440
    End
    Caption ="Species Search"
    OnClose ="[Event Procedure]"
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
    OrderByOnLoad =0
    SplitFormDatasheet =1
    OrderByOnLoad =0
    SplitFormDatasheet =1
    ShowPageMargins =0
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
            SpecialEffect =3
            BackStyle =0
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin OptionButton
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
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
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =5535
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =4020
                    Width =13536
                    Height =1140
                    BackColor =14806254
                    BorderColor =10921638
                    Name ="boxCurrTgtArea"
                    GridlineColor =10921638
                    LayoutCachedTop =4020
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =5160
                    BackThemeColorIndex =3
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =840
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSearchHdr"
                    Caption ="Search"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =432
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =840
                    Top =720
                    Width =8154
                    Height =1380
                    ColumnOrder =0
                    BorderColor =10921638
                    Name ="optgSpeciesType"
                    GridlineColor =10921638

                    LayoutCachedLeft =840
                    LayoutCachedTop =720
                    LayoutCachedWidth =8994
                    LayoutCachedHeight =2100
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =960
                            Top =600
                            Width =1596
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSpeciesTypes"
                            Caption ="What to Search..."
                            GridlineColor =10921638
                            LayoutCachedLeft =960
                            LayoutCachedTop =600
                            LayoutCachedWidth =2556
                            LayoutCachedHeight =900
                            BackThemeColorIndex =-1
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =1260
                    Top =1620
                    Width =240
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    Name ="cbxUT"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Utah species"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =1620
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =1560
                            Top =1560
                            Width =525
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblUtah"
                            Caption ="Utah"
                            ControlTipText ="Utah species"
                            GridlineColor =10921638
                            LayoutCachedLeft =1560
                            LayoutCachedTop =1560
                            LayoutCachedWidth =2085
                            LayoutCachedHeight =1875
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =2520
                    Top =1620
                    Width =240
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    Name ="cbxCO"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Utah species"
                    GridlineColor =10921638

                    LayoutCachedLeft =2520
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =2820
                            Top =1560
                            Width =900
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCO"
                            Caption ="Colorado"
                            ControlTipText ="Colorado species"
                            GridlineColor =10921638
                            LayoutCachedLeft =2820
                            LayoutCachedTop =1560
                            LayoutCachedWidth =3720
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =4140
                    Top =1620
                    Width =240
                    ColumnOrder =3
                    TabIndex =3
                    BorderColor =10921638
                    Name ="cbxWY"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Wyoming species"
                    GridlineColor =10921638

                    LayoutCachedLeft =4140
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4500
                            Top =1560
                            Width =936
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblWY"
                            Caption ="Wyoming"
                            ControlTipText ="Wyoming species"
                            GridlineColor =10921638
                            LayoutCachedLeft =4500
                            LayoutCachedTop =1560
                            LayoutCachedWidth =5436
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =5820
                    Top =1620
                    Width =240
                    ColumnOrder =4
                    TabIndex =4
                    BorderColor =10921638
                    Name ="cbxITIS"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="ITIS species"
                    GridlineColor =10921638

                    LayoutCachedLeft =5820
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6180
                            Top =1560
                            Width =405
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblITIS"
                            Caption ="ITIS"
                            GridlineColor =10921638
                            LayoutCachedLeft =6180
                            LayoutCachedTop =1560
                            LayoutCachedWidth =6585
                            LayoutCachedHeight =1875
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =7080
                    Top =1620
                    Width =240
                    ColumnOrder =5
                    TabIndex =5
                    BorderColor =10921638
                    Name ="cbxCommon"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Common name"
                    GridlineColor =10921638

                    LayoutCachedLeft =7080
                    LayoutCachedTop =1620
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =7440
                            Top =1560
                            Width =900
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCommon"
                            Caption ="Common"
                            ControlTipText ="Common name"
                            GridlineColor =10921638
                            LayoutCachedLeft =7440
                            LayoutCachedTop =1560
                            LayoutCachedWidth =8340
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =1200
                    Top =1020
                    Width =5700
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblChooseSpeciesType"
                    Caption ="Choose at least one species type or common name to search."
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =1335
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =2760
                    Width =6540
                    Height =360
                    ColumnOrder =6
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSearchFor"
                    DefaultValue ="\"\""
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1200
                    LayoutCachedTop =2760
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =3120
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =900
                            Top =2280
                            Width =4476
                            Height =300
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSearchFor"
                            Caption ="Enter the name or portion of name to search for."
                            GridlineColor =10921638
                            LayoutCachedLeft =900
                            LayoutCachedTop =2280
                            LayoutCachedWidth =5376
                            LayoutCachedHeight =2580
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =120
                    Top =4140
                    Width =1716
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSearchResults"
                    Caption ="Search Results"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =4140
                    LayoutCachedWidth =1836
                    LayoutCachedHeight =4512
                End
                Begin Label
                    OverlapFlags =215
                    Left =300
                    Top =4620
                    Width =7083
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSearchResultInstructions"
                    Caption ="Double click the species code you'd like to add to your target species listing. "
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =4620
                    LayoutCachedWidth =7383
                    LayoutCachedHeight =4920
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =87
                    Top =4020
                    Width =13536
                    BorderColor =8355711
                    Name ="lineCurrTgtAreaTop"
                    GridlineColor =10921638
                    LayoutCachedTop =4020
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =4020
                    BorderTint =50.0
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =87
                    Top =5160
                    Width =13536
                    BorderColor =8355711
                    Name ="lineCurrTgtAreaBtm"
                    GridlineColor =10921638
                    LayoutCachedTop =5160
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =5160
                    BorderTint =50.0
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =7320
                    Top =3300
                    Width =2220
                    TabIndex =7
                    ForeColor =16711680
                    Name ="btnSearch"
                    Caption ="Search!"
                    StatusBarText ="Add new target area"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7320
                    LayoutCachedTop =3300
                    LayoutCachedWidth =9540
                    LayoutCachedHeight =3660
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =6750156
                    BackThemeColorIndex =-1
                    BackTint =100.0
                    BorderColor =52377
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    HoverColor =3407769
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =52224
                    PressedThemeColorIndex =-1
                    PressedShade =100.0
                    HoverForeColor =2375487
                    HoverForeThemeColorIndex =-1
                    HoverForeTint =100.0
                    PressedForeColor =6750156
                    PressedForeThemeColorIndex =-1
                    PressedForeTint =100.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =93
                    Left =120
                    Top =5220
                    Width =1728
                    Height =300
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCodeHdr"
                    Caption ="Code"
                    Tag ="*"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =5220
                    LayoutCachedWidth =1848
                    LayoutCachedHeight =5520
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =93
                    Left =1920
                    Top =5220
                    Width =2304
                    Height =300
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblITISHdr"
                    Caption ="ITIS"
                    Tag ="*"
                    GridlineColor =10921638
                    LayoutCachedLeft =1920
                    LayoutCachedTop =5220
                    LayoutCachedWidth =4224
                    LayoutCachedHeight =5520
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =93
                    Left =4260
                    Top =5220
                    Width =2304
                    Height =300
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUTHdr"
                    Caption ="UT"
                    Tag ="*"
                    GridlineColor =10921638
                    LayoutCachedLeft =4260
                    LayoutCachedTop =5220
                    LayoutCachedWidth =6564
                    LayoutCachedHeight =5520
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =93
                    Left =6600
                    Top =5220
                    Width =2304
                    Height =300
                    BackColor =15788753
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCOHdr"
                    Caption ="CO"
                    Tag ="*"
                    GridlineColor =10921638
                    LayoutCachedLeft =6600
                    LayoutCachedTop =5220
                    LayoutCachedWidth =8904
                    LayoutCachedHeight =5520
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =8940
                    Top =5220
                    Width =2304
                    Height =300
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWYHdr"
                    Caption ="WY"
                    Tag ="*"
                    GridlineColor =10921638
                    LayoutCachedLeft =8940
                    LayoutCachedTop =5220
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =5520
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =93
                    Left =11280
                    Top =5220
                    Width =2304
                    Height =300
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCommonHdr"
                    Caption ="Common"
                    Tag ="*"
                    GridlineColor =10921638
                    LayoutCachedLeft =11280
                    LayoutCachedTop =5220
                    LayoutCachedWidth =13584
                    LayoutCachedHeight =5520
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =87
                    Top =5520
                    Width =13536
                    BorderColor =8355711
                    Name ="lineResultsTop"
                    GridlineColor =10921638
                    LayoutCachedTop =5520
                    LayoutCachedWidth =13536
                    LayoutCachedHeight =5520
                    BorderTint =50.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    Left =1980
                    Top =4140
                    Width =408
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFor"
                    Caption ="for"
                    GridlineColor =10921638
                    LayoutCachedLeft =1980
                    LayoutCachedTop =4140
                    LayoutCachedWidth =2388
                    LayoutCachedHeight =4512
                End
                Begin Label
                    OverlapFlags =215
                    Left =2700
                    Top =4140
                    Width =2448
                    Height =372
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblSearchForValue"
                    Caption ="species"
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =4140
                    LayoutCachedWidth =5148
                    LayoutCachedHeight =4512
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    Left =5820
                    Top =4140
                    Width =2448
                    Height =372
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblSpeciesFound"
                    Caption ="0 species found"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =4140
                    LayoutCachedWidth =8268
                    LayoutCachedHeight =4512
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =120
                    Width =1728
                    Height =300
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLUCode"
                    ControlSource ="LU_Code"
                    OnDblClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1848
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Width =2304
                    Height =300
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxMasterSpecies"
                    ControlSource ="Master_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedWidth =4224
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4260
                    Width =2304
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUTSpecies"
                    ControlSource ="Utah_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =4260
                    LayoutCachedWidth =6564
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =6600
                    Width =2304
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCOSpecies"
                    ControlSource ="CO_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =8904
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =8940
                    Width =2304
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxWYSpecies"
                    ControlSource ="WY_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =8940
                    LayoutCachedWidth =11244
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11280
                    Width =2304
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCmnName"
                    ControlSource ="Master_Common_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =11280
                    LayoutCachedWidth =13584
                    LayoutCachedHeight =300
                End
                Begin Label
                    Visible = NotDefault
                    FontItalic = NotDefault
                    OverlapFlags =247
                    Left =5520
                    Width =2304
                    Height =300
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblNoRecords"
                    Caption ="-- No species found --"
                    GridlineColor =10921638
                    LayoutCachedLeft =5520
                    LayoutCachedWidth =7824
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =10560
                    Width =2304
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxMasterPlantCode"
                    ControlSource ="Master_PLANT_Code"
                    Tag ="#"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedWidth =12864
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
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
' MODULE:       Form_SpeciesSearch
' Description:  Species search functions & procedures
' Version:      1.05
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015  - 1.00 - initial version
'               BLC - 6/26/2015 - 1.01 - added LU_Code to search
'               BLC - 6/30/2015 - 1.02 - removed unused subroutines
'                                        btnSearch_Enter() and SpeciesSearch()
'                                        both handled by btnSearch_Click()
'               BLC - 7/22/2015 - 1.03 - fixed search header highlighting
'               BLC - 7/5/2016  - 1.04 - adapted for big rivers, renamed SpeciesSearch & changed
'                                        Exit_Sub to Exit_Handler, added Form_Close()
'               BLC - 8/2/2016  - 1.05 - use Me.CallingForm, update documentation
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidCallingForm(value As String)


'=================================================================
'  Properties
'=================================================================
' ---------------------------------
' PROPERTY:     Maximized
' Description:  Indicates if form is maximized or not by checking IsZoomed()
' Assumptions:  none
' Parameters:   N/A
' Returns:      True(1) - form is maximized
'               False(0) - form is not maximized
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Get Maximized() As Boolean
     Maximized = IsZoomed(Me.hWnd) * 1
End Property

' ---------------------------------
' PROPERTY:     Minimized
' Description:  Indicates if form is minimized or not by checking IsIconic()
' Assumptions:  none
' Parameters:   N/A
' Returns:      True(1) - form is minimized
'               False(0) - form is not minimized
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Get Minimized() As Boolean
     Minimized = IsIconic(Me.hWnd) * 1
End Property

' ---------------------------------
' PROPERTY LET: Maximized
' Description:  Sets custom form property 'Maximized'
' Assumptions:
' Note:         The IsMax argument must be defined as the same data type
'               returned by the corresponding Property Get procedure for
'               the same custom property.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Let Maximized(IsMax As Boolean)
     If IsMax Then
         Me.SetFocus
         DoCmd.Maximize
     Else
         Me.SetFocus
         DoCmd.Restore
     End If
End Property

' ---------------------------------
' PROPERTY LET: Minimized
' Description:  Sets custom form property 'Minimized'
' Assumptions:
' Note:         The IsMin argument must be defined as the same data type
'               returned by the corresponding Property Get procedure for
'               the same custom property.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://support2.microsoft.com/?kbid=210190
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/23/2015  - initial version
' ---------------------------------
Public Property Let Minimized(IsMin As Boolean)
     If IsMin Then
         Me.SetFocus
         DoCmd.Minimize
     Else
         Me.SetFocus
         DoCmd.Restore
     End If
End Property

'---------------------
' Properties
'---------------------
Public Property Let CallingForm(value As String)
    If Len(value) > 0 Then
        m_CallingForm = value
    Else
        RaiseEvent InvalidCallingForm(value)
    End If
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
' SUB:          Form_Load
' Description:  Search form preparation action
' Assumptions:  none
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015  - initial version
'   BLC - 2/20/2015 - cleared selections & updated documentation
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'   BLC - 7/5/2016 - adjusted for big rivers & new form name SpeciesSearch vs. frmSpeciesSearch
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    
    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'set form caller
    TempVars("originForm") = Forms!SpeciesSearch.OpenArgs
    
'    Initialize
       
    'species type selections
    TempVars.Add "speciestype", ""
    
    'disable search until something is entered
    btnSearch.Enabled = False
    'DisableControl btnSearch
    
    'clear selections
    ClearFields Me

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxSearchFor_LostFocus
' Description:  Actions to take when search for textbox is not empty
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/10/2015 - initial version
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for enabled control
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub tbxSearchFor_LostFocus()
On Error GoTo Err_Handler
    
    If Len(tbxSearchFor.value) > 0 Then
        'check if species list is identified
        If Len(TempVars("speciestype")) > 0 Then
            'enable the search "button"
            btnSearch.Enabled = True
            'EnableControl btnSearch, CTRL_ADD_ENABLED, TEXT_ENABLED
        Else
            MsgBox "Please choose at least one species list to search.", vbOKOnly, "Oops! Missing Species List to Search"
        End If
    Else
        'disable the search "button"
        btnSearch.Enabled = False
        DisableControl btnSearch
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSearchFor_LostFocus[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxCO_Click
' Description:  actions on checkbox click
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxCO_Click()
On Error GoTo Err_Handler

If cbxCO = True Then

    cbxAddToList "speciestype", "CO", ";"

Else

    cbxRemoveFromList "speciestype", "CO", ";"

End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxCO_Click[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxUT_Click
' Description:  actions on checkbox click
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxUT_Click()
On Error GoTo Err_Handler

If cbxUT = True Then
    
    cbxAddToList "speciestype", "UT", ";"

Else
    
    cbxRemoveFromList "speciestype", "UT", ";"

End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUT_Click[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxWY_Click
' Description:  actions on checkbox click
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxWY_Click()
On Error GoTo Err_Handler

If cbxWY = True Then

    cbxAddToList "speciestype", "WY", ";"

Else

    cbxRemoveFromList "speciestype", "WY", ";"

End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxWY_Click[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxITIS_Click
' Description:  actions on checkbox click
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxITIS_Click()
On Error GoTo Err_Handler

If cbxITIS = True Then

    cbxAddToList "speciestype", "ITIS", ";"

Else

    cbxRemoveFromList "speciestype", "ITIS", ";"

End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxITIS_Click[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxCommon_Click
' Description:  actions on checkbox click
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
' ---------------------------------
Private Sub cbxCommon_Click()
On Error GoTo Err_Handler

If cbxCommon = True Then

    cbxAddToList "speciestype", "CMN", ";"
Else

    cbxRemoveFromList "speciestype", "CMN", ";"
End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxCO_Click[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxAddToList
' Description:  Add an item to a list
' Assumptions:  -
' Parameters:   list - listbox name
'               cbxValue - value to add
'               separator - delimiter for values
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub cbxAddToList(list As String, cbxValue As String, separator As String)
On Error GoTo Err_Handler
    
    'if list exists and item is in it, exit
    If Len(TempVars(list)) > 0 Then
        If CountInString(TempVars(list), cbxValue) > 0 Then
            GoTo Exit_Handler
        End If
    End If
        
    'add item if it's not already in list
    TempVars(list) = TempVars(list) & cbxValue & separator
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxAddToList[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxRemoveFromList
' Description:  Remove an item from a list
' Assumptions:  -
' Parameters:   list - listbox name
'               cbxValue - value to add
'               separator - delimiter for values
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/9/2015 - initial version
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub cbxRemoveFromList(list As String, cbxValue As String, separator As String)
On Error GoTo Err_Handler
    
    TempVars(list) = Replace(Replace(TempVars(list), cbxValue, ""), separator & separator, separator)
    
    'clear if only = separator
    If Len(TempVars(list)) = 1 And TempVars(list) = separator Then TempVars(list) = ""

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxRemoveFromList[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxLUCode_DblClick
' Description:  Add an item to the listbox if it is not a duplicate of items already listed
' Assumptions:  Assumes duplicates are not desired in the listbox
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015 - initial version
'   BLC - 2/21/2015 - fixed Runtime Error 451: Property let procedure not defined and property get Procedure did not return an object.
'                     changed from .ListIndex(i) to .Column(2,i) when iterating through list items
'   BLC - 2/23/2015 - added lblTgtSpeciesCount update
'   BLC - 5/27/2015 - added Transect_Only and Tgt_Area_ID values to item (";0;0")
'                     added check for missing LUCode
'   BLC - 5/29/2015 - added tbxMasterPlantCode, changed tbxResultCode to tbxLUCode
'                     swapped tbxMasterSpecies (ITIS) for tbxUTSpecies
'                     renamed tbxResultCode to tbxLUCode
'                     changed order of tbxMasterPlantCode and tbxLUCode to populate listbox
'                     (bugfix for search species missing proper LUcode)
'   BLC - 6/4/2015  - added handling to disable double-click for path coming from Search tab
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Private Sub tbxLUCode_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    Dim item As String
    Dim i As Integer
    Dim lbx As ListBox
    
    'check if coming from search tab & disable
    If TempVars("originForm") = "DisableDoubleClick" Then GoTo Exit_Handler
    
    'check for empty values --> tbxResultCode, tbxUTSpecies, tbxMasterSpecies - cannot be empty!
    If IsNull(tbxLUCode) Or Len(Trim(tbxLUCode)) = 0 Then
        MsgBox "Species " & tbxMasterSpecies & " is missing a lookup code (LU_Code). " & _
            vbCrLf & vbCrLf & "This code is required before the species can be added to a target list. " & _
            vbCrLf & vbCrLf & "Please determine the appropriate code and enter it into the master " & _
            "plant species list." & _
            vbCrLf & vbCrLf & "Contact the project ecologist/data manager to add the species. ", _
            vbExclamation, "Missing Lookup Code!"

        'email species desired
        
        GoTo Exit_Handler

    End If
    
    'add components of item (code, species (UT or whatever), & ITIS) to listbox

    'prepare item for listbox value
    item = tbxMasterPlantCode & ";" & tbxMasterSpecies & ";" & tbxLUCode & ";0;0"
    
    'iterate through listbox (use .Column(x,i) vs .ListIndex(i) which results in error 451 property let not defined, property get...)
    If IsListDuplicate(Forms("frm_Tgt_Species").Controls("lbxTgtSpecies"), 2, tbxLUCode) Then
        'duplicate, so exit
        GoTo Exit_Handler
    End If
    
    Set lbx = Forms("frm_Tgt_Species").Controls("lbxTgtSpecies")
    
    With lbx
        'add item if not duplicate
        .AddItem item
    
        'update target species count
        Forms("frm_Tgt_Species").Controls("lblTgtSpeciesCount").Caption = .ListCount - 1 & " species"

    End With
    
    'minimize search form
    DoCmd.SelectObject acForm, Me.Name, False
    DoCmd.Minimize
    
    'return focus to calling form
    Dim origin As String
    origin = TempVars("originForm")
    If Len(origin) > 0 Then
        DoCmd.SelectObject acForm, origin, False
        DoCmd.Restore
    End If
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxResultCode_DblClick[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnSearch_Click
' Description:  Search for the name or portion of a name in the species/common names listed & return a result list
' Assumptions:
' Note:         Returns all species/common names from tlu_NCPN_Plants that contain the search string.
'               The string may be found at the beginning, middle or end of a name to be included.
'               Special search strings like "*" (not including quotes) will return ALL species in the table.
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' http://codevba.com/msaccess/status_bar_and_progress_meter.htm#.VNb9X_lM4_4
' Adapted:      Bonnie Campbell, February 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 2/20/2015 - added header highlighting
'   BLC - 2/23/2015 - fixed duplicate results (SELECT DISTINCT...)
'   BLC - 5/13/2015 - revised to use global constants vs. tempvars for enabled control
'   BLC - 5/14/2015 - revised to leave checkbox list intact to avoid error message @ choosing a species type
'                     when checkbox was left checked
'   BLC - 5/29/2015 - added Master_PLANT_Code to selection (bugfix for search species missing proper LUcode)
'                     renamed tbxResultCode to tbxLUCode
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
'   BLC - 6/26/2015 - added LU_Code to search to enable search against code to find species
'   BLC - 7/22/2015 - fixed search header highlighting
' ---------------------------------
Private Sub btnSearch_Click()
On Error GoTo Err_Handler
    
    Dim speciestype As Variant
    Dim strSearch As String, strSpecies As String, strWhere As String, strSQL As String
    Dim i As Integer

    'ignore if disabled
    If btnSearch.Enabled = False Then GoTo Exit_Handler

    strSearch = Trim(tbxSearchFor.value)
            
    'check strSearch is alpha numeric
    
    'check if species list is selected
    If Len(TempVars("speciestype")) > 0 Then
        'enable the search "button"
        btnSearch.Enabled = True
        'EnableControl btnSearch, CTRL_ADD_ENABLED, TEXT_ENABLED
    Else
        MsgBox "Please choose at least one species list to search.", vbOKOnly, "Oops! Missing Species List to Search"
        GoTo Exit_Handler
    End If
    
    'determine which species names are to be searched (ITIS, UT, CO, WY, Common)
    strWhere = " WHERE "
        
    'reset headers
    ResetHeaders Me, True, "*", False, 0, 8355711 ', vbWhite '#7F7F7F rgb(127,127,127)
            
    'determine which species names to check
    Dim listTypes() As String
    'add the 6-letter code to the search matches
    If Len(TempVars("speciestype")) > 0 Then
        'TempVars("speciestype") = TempVars("speciestype") & "CODE;"
        cbxAddToList "speciestype", "CODE", ";"
    'Else
    '    TempVars("speciestype") = "CODE;"
    End If
    listTypes = Split(TempVars("speciestype"), ";")
    
    For Each speciestype In listTypes
        
        If Len(speciestype) > 0 Then
            
            'If CountInString(speciestype, ";") > 1 Then
            i = i + 1
            If i > 1 Then
                strWhere = strWhere & " OR "
            
            End If
        
            'forecolor 16737792 '#0066FF rgb(0,102,255)
            'backcolor 15788753 '#D1EAF0 rgb(209,234,240)
            Select Case speciestype
                Case "CO"   'Colorado
                    strSpecies = "CO_Species"
                    ResetHeaders Me, False, "*", True, 1, 16737792, 15788753, lblCOHdr
                Case "UT"   'Utah
                    strSpecies = "Utah_Species"
                    ResetHeaders Me, False, "*", True, 1, 16737792, 15788753, lblUTHdr
                Case "WY"   'Wyoming
                    strSpecies = "WY_Species"
                    ResetHeaders Me, False, "*", True, 1, 16737792, 15788753, lblWYHdr
                Case "ITIS" 'Master
                    strSpecies = "Master_Species"
                    ResetHeaders Me, False, "*", True, 1, 16737792, 15788753, lblITISHdr
                Case "CMN"  'Common
                    strSpecies = "Master_Common_Name"
                    ResetHeaders Me, False, "*", True, 1, 16737792, 15788753, lblCommonHdr
                Case "CODE"
                    strSpecies = "LU_Code"
            End Select
                    
            strWhere = strWhere & " " & strSpecies & " LIKE '*" & strSearch & "*'"
            
        End If
    Next
    
    'prep WHERE clause
    If Len(Replace(strWhere, "WHERE", "")) = 0 Then strWhere = ""
    
    'build SQL statement
    strSQL = "SELECT DISTINCT LU_Code, Master_Species, Utah_Species, CO_Species, WY_Species, " _
            & "Master_Common_Name, Master_PLANT_Code " _
            & "FROM tlu_NCPN_Plants " _
            & strWhere & ";"
               
    'run search
    Dim rs As DAO.Recordset
      
    'fetch data
    Set rs = CurrentDb.OpenRecordset(strSQL) ', dbOpenSnapshot)

    'set form results
    Set Me.Recordset = rs
    tbxLUCode.ControlSource = "LU_Code"
    tbxMasterSpecies.ControlSource = "Master_Species" 'ITIS
    tbxUTSpecies.ControlSource = "Utah_Species"
    tbxCOSpecies.ControlSource = "CO_Species"
    tbxWYSpecies.ControlSource = "WY_Species"
    tbxCmnName.ControlSource = "Master_Common_Name"
    tbxMasterPlantCode.ControlSource = "Master_PLANT_Code"

    'turn fields on (includes lblNoRecords, controls w/o & w/ * tags)
    ShowControls Me, True, "", True
    ShowControls Me, True, "*", True
        
    ' determine record count
    Dim Count As Integer
    If Not rs.EOF Then
        rs.MoveLast
        Count = rs.RecordCount
        rs.MoveFirst
        
        'hide no records
        lblNoRecords.Visible = False
    Else
        lblNoRecords.Visible = True
    End If
        
    'set # species found
    lblSpeciesFound.Caption = Count & " species found"
        
    'set search for caption
    lblSearchForValue.Caption = """" & strSearch & """"
    
    'extend form if species count > 0
    If Count > 0 Then
        SetWindowSize Me, 8000, Me.Width
    End If
    
    'set statusbar notice
    Dim varReturn As Variant
    varReturn = SysCmd(acSysCmdSetStatus, "Searching for " & strSearch & "...")
    
    'clear fields
    ClearFields Me
    
    'leave last selections for checkboxes (don't clear TempVars.item("speciestype"))
    'must clear to clear highlighting & reset speciestypes
    TempVars.item("speciestype") = ""

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSearch_Click[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 5, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/5/2016 - initial version
'   BLC - 8/2/2016 - use Me.CallingForm
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[SpeciesSearch form])"
    End Select
    Resume Exit_Handler
End Sub
