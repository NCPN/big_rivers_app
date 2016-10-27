Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7860
    DatasheetFontHeight =11
    ItemSuffix =69
    Left =3360
    Top =2775
    Right =17595
    Bottom =14625
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x9116deeeb5cfe440
    End
    Caption ="VegPlot"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin ComboBox
            AddColon = NotDefault
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
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =1395
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="VegPlot"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =420
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Enter the plot information and click save.\015\012Add cover species via buttons "
                        "at right."
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =735
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3480
                    Top =1080
                    Width =2025
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblModalSedSize"
                    Caption ="Modal Sediment Size"
                    GridlineColor =10921638
                    LayoutCachedLeft =3480
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =900
                    Width =720
                    ForeColor =16711680
                    Name ="btnComment"
                    Caption ="í ½í·©"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =900
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =-1
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1920
                    Top =1080
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDistance"
                    Caption ="Distance (m)"
                    GridlineColor =10921638
                    LayoutCachedLeft =1920
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3165
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1080
                    Top =1080
                    Width =600
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumber"
                    Caption ="Plot #"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =3600
                    Top =60
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="CANY  >  Green  >  HC  >  "
                    GridlineColor =10921638
                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5820
                    Top =900
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnSetObserverRecorder"
                    Caption ="í ½í±¥"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5820
                    LayoutCachedTop =900
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1260
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =8160
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5820
                    Top =1320
                    Width =1860
                    Height =1860
                    FontSize =14
                    LeftMargin =72
                    TopMargin =144
                    BackColor =12444887
                    Name ="lblCover"
                    Caption ="Cover\015\012Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =1320
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3180
                    BackThemeColorIndex =6
                    BackTint =40.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =2040
                    Width =5580
                    Height =1140
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =11916796
                    Name ="lblChkboxes"
                    Caption ="âœ”"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =3180
                    BackThemeColorIndex =9
                    BackTint =40.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =540
                    Width =4140
                    Height =1380
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =16051931
                    Name ="lblPct"
                    Caption ="%"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =540
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1920
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =16
                    ForeColor =4210752
                    Name ="btnSave"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Save Record"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0687050c06860ffb05850ffa05050ffa05050ff ,
                        0xa05050ff904850ff904840ff904840ff804040ff803840ff803840ff703840ff ,
                        0x703830ff0000000000000000d06870fff09090ffe08080ffb04820ff403020ff ,
                        0xc0b8b0ffc0b8b0ffd0c0c0ffd0c8c0ff505050ffa04030ffa04030ffa03830ff ,
                        0x703840ff0000000000000000d07070ffff98a0fff08880ffe08080ff705850ff ,
                        0x404030ff907870fff0e0e0fff0e8e0ff908070ffa04030ffa04040ffa04030ff ,
                        0x803840ff0000000000000000d07870ffffa0a0fff09090fff08880ff705850ff ,
                        0x000000ff404030fff0d8d0fff0e0d0ff807860ffb04840ffb04840ffa04040ff ,
                        0x804040ff0000000000000000d07880ffffa8b0ffffa0a0fff09090ff705850ff ,
                        0x705850ff705850ff705850ff706050ff806860ffc05850ffb05050ffb04840ff ,
                        0x804040ff0000000000000000e08080ffffb0b0ffffb0b0ffffa0a0fff09090ff ,
                        0xf08880ffe08080ffe07880ffd07070ffd06870ffc06060ffc05850ffb05050ff ,
                        0x904840ff0000000000000000e08890ffffb8c0ffffb8b0ffd06060ffc06050ff ,
                        0xc05850ffc05040ffb05030ffb04830ffa04020ffa03810ffc06060ffc05850ff ,
                        0x904840ff0000000000000000e09090ffffc0c0ffd06860ffffffffffffffffff ,
                        0xfff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffe0c8c0ffa03810ffc06060ff ,
                        0x904850ff0000000000000000e098a0ffffc0c0ffd07070ffffffffffffffffff ,
                        0xfffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffe0d0c0ffa04020ffd06860ff ,
                        0xa05050ff0000000000000000f0a0a0ffffc0c0ffe07870ffffffffffffffffff ,
                        0xfffffffffffffffffff8f0fff0f0f0fff0e8e0fff0d8d0ffb04830ffd07070ff ,
                        0xa05050ff0000000000000000f0a8a0ffffc0c0ffe08080ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffff8f0fff0f0f0fff0e8e0ffb05030ffe07880ff ,
                        0xa05050ff0000000000000000f0b0b0ffffc0c0fff08890ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffff8f0fff0f0f0ffc05040ff603030ff ,
                        0xb05850ff0000000000000000f0b0b0ffffc0c0ffff9090ffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff8f0ffc05850ffb05860ff ,
                        0xb05860ff0000000000000000f0b8b0fff0b8b0fff0b0b0fff0b0b0fff0a8b0ff ,
                        0xf0a0a0ffe098a0ffe09090ffe09090ffe08890ffe08080ffd07880ffd07870ff ,
                        0xd07070ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6660
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =420
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =3660
                    Width =7650
                    Height =4380
                    TabIndex =17
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.VegPlotList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3660
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =8040
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =18
                    ForeColor =4210752
                    Name ="btnUndo"
                    Caption ="Edit"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Undo/Clear values"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000f0906060d0784080b0583010000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000e0785040f08850ffd07040ffa05830500000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000f0906020d0704060f08050ffd07050f0a050300000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000c06840d0f08850ffc078508000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0c0b01000000000000000000000000090482040e07840ffe08860ffe0a08000 ,
                        0x00000000000000000000000000000000d07040ffd07040ffc06840ffb06030ff ,
                        0xb05830ff905030ff0000000000000000b0603020c06840ffe08050ffd0886080 ,
                        0x00000000000000000000000000000000d07850ffe07030fff08050fff09870ff ,
                        0xe09060fff0a08040000000000000000080402000c06840ffe07840f0e09870c0 ,
                        0x00000000000000000000000000000000d08050ffe08050fff09060fff0a070ff ,
                        0x904830b0b0603040000000000000000080402000c06840ffd07040f0e09870d0 ,
                        0x00000000000000000000000000000000d08860ffe09060fff09870fff08850f0 ,
                        0xb06040ffb06040ffb060307000000000b0805020a05830f0d07840f0e09070d0 ,
                        0x000000000000000000000000e0b09010c08060ffd09870e0d0886090d09070ff ,
                        0xd08050ffc07040ffc06840ffb06030c0b07040e0a06040ffe08050ffd0a080e0 ,
                        0x00000000000000000000000000000000c08860ffd0a0804000000000d08860c0 ,
                        0xd08860ffd08050f0c06840ffb06840ffb06030f0e07840f0e0a080f0d09880e0 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0a880c0e09880ffe09870f0e09070f0e09070e0e0a080f0e0a890f0f0b8a020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000f0b89060f0b090c0f0b8a0e0f0c0a0c0f0c0a090f0c0b02000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =420
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =3540
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3540
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =8160
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7560
                    Top =105
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =19
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =7560
                    LayoutCachedTop =105
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =60
                    Width =720
                    Height =315
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDistance"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2160
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =720
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNumber"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1020
                    Top =855
                    Height =315
                    TabIndex =3
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctFines"
                    ValidationRule ="Not Like \"[0-9]*.[0-9]*\""
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =855
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1170
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2580
                    Top =840
                    Height =315
                    TabIndex =4
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctWater"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =840
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1155
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1020
                    Top =1500
                    Height =315
                    TabIndex =5
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctURC"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =1500
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1815
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4800
                    Top =840
                    Width =960
                    Height =315
                    TabIndex =7
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPlotDensity"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Set plot density 1/x where X = 1, 2, 4, or 8"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedTop =840
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1155
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =2580
                    Top =1500
                    Height =315
                    TabIndex =6
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctFA"
                    ValidationRule ="Is Null Or \"T\" Or Between 0 and 101"
                    ValidationText ="Values may be whole percentages (0-100), 0.5, or T"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =1500
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =1815
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =2580
                    Top =1260
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFA"
                    Caption ="FA"
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =1260
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1500
                End
                Begin Label
                    OverlapFlags =255
                    Left =4440
                    Top =540
                    Width =1200
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPlotDensity"
                    Caption ="Plot Density"
                    GridlineColor =10921638
                    LayoutCachedLeft =4440
                    LayoutCachedTop =540
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =855
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =600
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFines"
                    Caption ="Fines"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =600
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =215
                    Left =2580
                    Top =600
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblWater"
                    Caption ="Water"
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =600
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =840
                End
                Begin Label
                    OverlapFlags =215
                    Left =1020
                    Top =1260
                    Width =1020
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblURC"
                    Caption ="Total URC"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1500
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =1140
                    Top =2280
                    Width =360
                    Height =360
                    TabIndex =8
                    BorderColor =10921638
                    Name ="chkNoCanopyVeg"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =2280
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =2640
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1380
                            Top =2220
                            Width =1485
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNoCanopyVeg"
                            Caption ="No Canopy Veg"
                            GridlineColor =10921638
                            LayoutCachedLeft =1380
                            LayoutCachedTop =2220
                            LayoutCachedWidth =2865
                            LayoutCachedHeight =2535
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =1140
                    Top =2700
                    Width =360
                    Height =360
                    TabIndex =9
                    BorderColor =10921638
                    Name ="chkNoRootedVeg"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =2700
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =3060
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1380
                            Top =2640
                            Width =1470
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNoRootedVeg"
                            Caption ="No Rooted Veg"
                            GridlineColor =10921638
                            LayoutCachedLeft =1380
                            LayoutCachedTop =2640
                            LayoutCachedWidth =2850
                            LayoutCachedHeight =2955
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =3300
                    Top =2280
                    Width =360
                    Height =360
                    TabIndex =10
                    BorderColor =10921638
                    Name ="chkNoIndicatorSpecies"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =2280
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =2640
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            Left =3540
                            Top =2220
                            Width =1965
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNoIndicatorSpecies"
                            Caption ="No Indicator Species"
                            GridlineColor =10921638
                            LayoutCachedLeft =3540
                            LayoutCachedTop =2220
                            LayoutCachedWidth =5505
                            LayoutCachedHeight =2535
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =3300
                    Top =2700
                    Width =360
                    Height =360
                    TabIndex =11
                    BorderColor =10921638
                    Name ="chkHasSocialTrails"
                    AfterUpdate ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =2700
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =3060
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3540
                            Top =2640
                            Width =1530
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblHasSocialTrails"
                            Caption ="Has Social Trails"
                            GridlineColor =10921638
                            LayoutCachedLeft =3540
                            LayoutCachedTop =2640
                            LayoutCachedWidth =5070
                            LayoutCachedHeight =2955
                        End
                    End
                End
                Begin Label
                    OverlapFlags =119
                    Left =4440
                    Top =855
                    Width =315
                    Height =315
                    BorderColor =8355711
                    Name ="lblFraction"
                    Caption ="1 /"
                    GridlineColor =10921638
                    LayoutCachedLeft =4440
                    LayoutCachedTop =855
                    LayoutCachedWidth =4755
                    LayoutCachedHeight =1170
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =20
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =840
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6840
                    Top =2100
                    Width =720
                    TabIndex =14
                    ForeColor =4210752
                    Name ="btnURC"
                    Caption ="URC"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Understory Rooted Cover Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2100
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =2460
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6840
                    Top =1560
                    Width =720
                    TabIndex =13
                    ForeColor =4210752
                    Name ="btnWCC"
                    Caption ="WCC"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Woody Canopy Cover Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1560
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =1920
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =6840
                    Top =2640
                    Width =720
                    TabIndex =15
                    ForeColor =4210752
                    Name ="btnARC"
                    Caption ="ARC"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit All Rooted Cover Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =6840
                    LayoutCachedTop =2640
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =3000
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5820
                    Top =540
                    Width =1860
                    Height =660
                    FontSize =20
                    LeftMargin =72
                    TopMargin =72
                    BackColor =12835293
                    Name ="lblTagline"
                    Caption ="í ½í³"
                    ControlTipText ="Add/Edit Tagline Measurements"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =540
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =3
                    BackShade =90.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6840
                    Top =630
                    Width =720
                    Height =480
                    TabIndex =12
                    ForeColor =4210752
                    Name ="btnTaglines"
                    Caption ="Tagline"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add/Edit Tagline Measurements"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a890ff604830ff604830ff604830ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a890fffff0e0ffffe0d0ffffe0c0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffff8f0ff000000ff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffff8f0fffff8f0ffffe8d0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffff8fffffff8f0ff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffffffffffff8fffffff0e0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0ffffffffff000000ff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0fffffffffffffffffff0f0f0ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0a8a0ffffffffffffffffff000000ff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0b0a0ffffffffffffffffffffffffff604830ff00000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000c0b0a0ffa08870ff806040ff705040ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff00000000c0b0a0fffffffffff0f8f0fff0f8f0ff705040ffffffffff ,
                        0x000000fff0f0f0ff000000ffffe8d0ff000000ffffe0c0ff000000ffffd8c0ff ,
                        0x604830ff00000000c0b0a0fffffffffffffffffff0f8f0ff805840ffffffffff ,
                        0xffffffffffffffff000000fffff8f0fffff0e0ffffe8e0ff000000ffffd8c0ff ,
                        0x604830ff00000000c0b0a0ffffffffffffffffffffffffffa08070ffffffffff ,
                        0xfffffffffffffffffffffffffff8f0fffff8f0fffff0e0ffffe8e0ffffe8d0ff ,
                        0x604830ff00000000c0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ffc0b0a0ff ,
                        0xc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a8a0ffc0a890ff ,
                        0xc0a890ff00000000
                    End

                    LayoutCachedLeft =6840
                    LayoutCachedTop =630
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =1110
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =3300
                    Top =60
                    Width =2304
                    Height =315
                    TabIndex =2
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxModalSedSize"
                    RowSourceType ="Table/Query"
                    RowSource ="PARAMETERS etype Text ( 255 ); SELECT DISTINCT id, label, summary, label & ' - '"
                        " & summary AS display, Sequence FROM AppEnum WHERE EnumType = ModWentworthClassS"
                        "ize ORDER BY Sequence; "
                    ColumnWidths ="0;0;0;2"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Size class"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3300
                    LayoutCachedTop =60
                    LayoutCachedWidth =5604
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff200001a0000005b00 ,
                        0x7400620078004d006f00640061006c00530065006400530069007a0065005d00 ,
                        0x2e00560061006c00750065003d00220022000000000000000000000000000000 ,
                        0x0000000000000000000000030000000100000000000000ffffff000200000022 ,
                        0x002200000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =300
                    Top =1440
                    Width =600
                    Height =420
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintPct"
                    Caption ="Nearest 1% or T"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =1440
                    LayoutCachedWidth =900
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =3300
                    Width =7860
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =3300
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =3615
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =3120
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16772541
                    Name ="lblMsgIcon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =3120
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =3720
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin FormFooter
            Height =0
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
' Form:         VegPlot
' Level:        Application form
' Version:      1.04
' Basis:        Dropdown form
'
' Description:  Vegplot form object related properties, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 8/23/2016 - 1.01 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 9/8/2016  - 1.02 - added SetObserverRecorder button
'               BLC - 10/3/2016 - 1.03 - disable taglines for CANY & DINO
'               BLC - 10/25/2016 - 1.04 - add CallingForm property & remove ButtonCaption,
'                                         SelectedID, SelectedValue properties
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:
'   Certain fields are not applicable across all parks as noted below.
'       % Water (inundation) - DINO & CANY
'       % Total URC - BLCA & CANY
'       Plot Density - BLCA & CANY
'       No Canopy Veg - BLCA & CANY
'       No Indicator Species - BLCA only
'       No Rooted Veg - DINO & CANY
'       Has Social Trail - BLCA & CANY
'
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 7/13/2016 - added validation, hints
'   BLC - 9/8/2016  - added SetObserverRecorder button
'   BLC - 10/24/2016 - revised to use CallingForm property, GetContext()
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Me.OpenArgs) > 0 Then Me.CallingForm = Me.OpenArgs
    
    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
                 
    Title = "VegPlot"
    Me.lblTitle.Caption = "" 'clear header title
    Directions = "Enter the plot information and click save." _
                & vbCrLf & "Add cover species via buttons at right."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    lblChkboxes.Caption = StringFromCodepoint(uCheck)
    lblCover.Caption = "Cover" & vbCrLf & "Species"
    lblTagline.Caption = StringFromCodepoint(uRuler)
    btnSetObserverRecorder.Caption = StringFromCodepoint(uUsers)
    btnSetObserverRecorder.ControlTipText = "Set the selected veg plot's observer & recorder"
    
    'hints
    lblHintPct.Caption = "Nearest 1% or T"
    lblHintPct.ForeColor = lngBlue
    
    'validation
    Dim strPctValidation As String, strValidationText As String
    strPctValidation = "Is Null Or ""T"" Or Between 0 and 101"
    tbxPctFines.ValidationRule = "Not Like ""[0-9]*.[0-9]*"""
    tbxPctWater.ValidationRule = strPctValidation
    tbxPctURC.ValidationRule = strPctValidation
    tbxPctFA.ValidationRule = strPctValidation
    
    strValidationText = "Values may be whole percentages (0-100), 0.5, or T"
    tbxPctFines.ValidationText = strValidationText
    tbxPctWater.ValidationText = strValidationText
    tbxPctURC.ValidationText = strValidationText
    tbxPctFA.ValidationText = strValidationText
    
    'set hover
    btnSetObserverRecorder.HoverColor = lngGreen
    btnComment.HoverColor = lngGreen
    btnTaglines.HoverColor = lngGreen
    btnWCC.HoverColor = lngGreen
    btnURC.HoverColor = lngGreen
    btnARC.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    tbxNumber.BackColor = lngYellow
    tbxDistance.BackColor = lngYellow
    cbxModalSedSize.BackColor = lngYellow
    tbxPctFines.BackColor = lngYellow
    tbxPctWater.BackColor = lngYellow
    tbxPctURC.BackColor = lngYellow
    tbxPlotDensity.BackColor = lngYellow
    btnSetObserverRecorder.Enabled = False
    
    'populate modal sediment size
    ' -------------------------------------------------------------------------------------
    ' NOTE: s_enums_for_type *MUST* include "DISTINCT" for the combobox autoexpand to work!(Access bug)
    '       Dan Some, August 7, 2011
    '       http://answers.microsoft.com/en-us/office/forum/office_2007-access/combo-box-property-auto-expand-yes-doesnt-seem-to/05fa61af-853e-4c9d-a3e3-2f51aa094668
    ' -------------------------------------------------------------------------------------
    cbxModalSedSize.RowSource = GetTemplate("s_enums_for_type", "etype" & PARAM_SEPARATOR & "ModWentworthClassSize")
    cbxModalSedSize.BoundColumn = 1 'bind to label (not ID)
    cbxModalSedSize.ColumnCount = 4
    cbxModalSedSize.ColumnWidths = "0,0,0,1.5" 'display the display column (combines label - summary)
    
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0
  
    'defaults --> turn off items
    lblWater.Visible = False
    tbxPctWater.Visible = False
    lblURC.Visible = False
    tbxPctURC.Visible = False
    lblPlotDensity.Visible = False
    lblFraction.Visible = False
    tbxPlotDensity.Visible = False
    chkNoCanopyVeg.Visible = False
    lblNoCanopyVeg.Visible = False
    chkNoIndicatorSpecies.Visible = False
    lblNoIndicatorSpecies.Visible = False
    chkNoRootedVeg.Visible = False
    lblNoRootedVeg.Visible = False
    chkHasSocialTrails.Visible = False
    lblHasSocialTrails.Visible = False
    btnWCC.Visible = False
    btnURC.Visible = False
    btnARC.Visible = False
    btnTaglines.Enabled = False
    
    'default plot density = 3 starting in 2015 (i.e. 1/3 density)
    tbxPlotDensity.Value = 3
    
    'adjust UI based on park
    Select Case TempVars("ParkCode")
        
        Case "BLCA"
            lblURC.Visible = True
            tbxPctURC.Visible = True
            lblPlotDensity.Visible = True
            lblFraction.Visible = True
            tbxPlotDensity.Visible = True
            chkNoCanopyVeg.Visible = True
            lblNoCanopyVeg.Visible = True
            chkNoIndicatorSpecies.Visible = True
            lblNoIndicatorSpecies.Visible = True
            chkHasSocialTrails.Visible = True
            lblHasSocialTrails.Visible = True
            btnWCC.Visible = True
            btnURC.Visible = True
            btnTaglines.Enabled = True
            
        Case "CANY"
            lblURC.Visible = True
            tbxPctURC.Visible = True
            lblWater.Visible = True
            tbxPctWater.Visible = True
            lblPlotDensity.Visible = True
            lblFraction.Visible = True
            tbxPlotDensity.Visible = True
            tbxPctURC.Visible = True
            chkNoCanopyVeg.Visible = True
            lblNoCanopyVeg.Visible = True
            chkNoRootedVeg.Visible = True
            lblNoRootedVeg.Visible = True
            chkHasSocialTrails.Visible = True
            lblHasSocialTrails.Visible = True
            btnWCC.Visible = True
            btnURC.Visible = True
            
        Case "DINO"
            lblWater.Visible = True
            tbxPctWater.Visible = True
            chkNoRootedVeg.Visible = True
            lblNoRootedVeg.Visible = True
            btnARC.Visible = True
    
    End Select
    
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'initialize values
    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
              
'      If tbxID > 0 Then btnComment.Enabled = True
    btnSetObserverRecorder.Enabled = IIf(tbxID.Value > 0, True, False)

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxNumber_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxNumber_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxNumber.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNumber_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxDistance_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxDistance_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxDistance.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDistance_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxModalSedSize_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
'   BLC - 7/13/2016 - revised to combobox
' ---------------------------------
Private Sub cbxModalSedSize_AfterUpdate()
On Error GoTo Err_Handler

    If Len(cbxModalSedSize.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxModalSedSize_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctFines_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPctFines_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctFines.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctFines_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctWater_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPctWater_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctWater.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctWater_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctURC_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPctURC_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctURC.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctURC_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPctFA_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPctFA_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPctFA.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPctFA_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxPlotDensity_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub tbxPlotDensity_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxPlotDensity.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPlotDensity_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          chkNoCanopyVeg_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub chkNoCanopyVeg_AfterUpdate()
On Error GoTo Err_Handler

    If Len(chkNoCanopyVeg.Value) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkNoCanopyVeg_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          chkNoRootedVeg_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub chkNoRootedVeg_AfterUpdate()
On Error GoTo Err_Handler

    If Len(chkNoRootedVeg.Value) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkNoRootedVeg_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          chkNoIndicatorSpecies_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub chkNoIndicatorSpecies_AfterUpdate()
On Error GoTo Err_Handler

    If Len(chkNoIndicatorSpecies.Value) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkNoIndicatorSpecies_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          chkHasSocialTrails_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub chkHasSocialTrails_AfterUpdate()
On Error GoTo Err_Handler

    If Len(chkHasSocialTrails.Value) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkHasSocialTrails_AfterUpdate[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnUndo_Click
' Description:  Undo button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 6/27/2016 - revised to use ClearForm()
' ---------------------------------
Private Sub btnUndo_Click()
On Error GoTo Err_Handler
    
    ClearForm Me
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUndo_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSave_Click
' Description:  Save button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    Dim vp As New VegPlot
    
    With vp
        'values passed into form


'        .CollectionSourceName = "T"

'        .CreateDate = ""
'        .CreatedByID = 0
'        .LastModified = ""
'        .LastModifiedByID = 0
        
        '.ProtocolID = 1
        '.SiteID = 1
        
        'form values
        
'        .EventID = cbxEvent.Column(0)
'        .FeatureID = cbxFeature.Column(0)
        
        .PlotNumber = tbxNumber.Value
        .PlotDistance = tbxDistance.Value
        .ModalSedimentSize = cbxModalSedSize.Value
        
        .PlotDensity = tbxPlotDensity.Value
        
        'pct values
        .PercentFines = tbxPctFines.Value
        .PercentWater = tbxPctWater.Value
        .UnderstoryRootedPctCover = tbxPctURC.Value
        
        'chk values
        .NoCanopyVeg = chkNoCanopyVeg.Value
        .NoRootedVeg = chkNoRootedVeg.Value
        .NoIndicatorSpecies = chkNoIndicatorSpecies.Value
        .HasSocialTrail = chkHasSocialTrails.Value
        
'        .LocationName = tbxName.Value
'        .LocationType = "" 'cbxLocationType.SelText
'
'        .HeadtoOrientDistance = tbxDistance.Value
'        .HeadtoOrientBearing = tbxBearing.Value
        
        .ID = tbxID.Value '0 if new, edit if > 0
        .SaveToDb
        
        'set the tbxID.value
        tbxID = .ID
        
    End With
    
    'clear values & refresh display
    
    ReadyForSave
    
    PopulateForm Me, tbxID.Value
    
    If tbxID.Value > 0 Then
        'highlight SetObserverRecorder button
        btnSetObserverRecorder.borderColor = lngRed
        lblMsg.ForeColor = lngYellow
        lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
        lblMsg.Caption = "Don't forget to set observer & recorder!"
    End If
    'refresh list
    Me.list.Requery
    
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSetObserverRecorder_Click
' Description:  Set observer recorder button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 8, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/8/2016 - initial version
' ---------------------------------
Private Sub btnSetObserverRecorder_Click()
On Error GoTo Err_Handler
    
    DoCmd.OpenForm "SetObserverRecorder", acNormal, , , , , "VegPlot|" & Me.tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSetObserverRecorder_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnTaglines_Click
' Description:  Tagline button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub btnTaglines_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "Tagline", acNormal, , , , , Me.Name & "|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTaglines_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnWCC_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnWCC_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegWalk", acNormal, , , , , "WoodyCanopySpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnWCC_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnURC_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnURC_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegWalk", acNormal, , , , , "UnderstoryRootedSpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnURC_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnARC_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnARC_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "VegWalk", acNormal, , , , , "AllRootedSpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnARC_Click[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnComment_Click
' Description:  Undo button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "VegPlot|" & tbxID.Text
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[VegPlot form])"
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 6/27/2016 - revised to use ToggleForm()
'   BLC - 10/24/2016 - revised to use CallingForm property
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore CallingForm
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyForSave
' Description:  Check if form values are ready to save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires:  EventID, SiteID, FeatureID, VegTransectID, PlotDistance_m,
    '           ModalSedSize, PctFine, PctWater, PctURC, PlotDensity,
    '           NoCanopyVeg, NoRootedVeg, HasSocialTrail, FA
    '           BLCA only: NoIndicatorSpecies
    If Nz(tbxDistance.Value, "") > 0 _
        And Nz(cbxModalSedSize.Value, "") > -1 _
        And Nz(tbxPctFines.Value, "") > -1 _
        And Nz(tbxPctWater.Value, "") > -1 _
        And Nz(tbxPctURC.Value, "") > -1 _
        And Nz(tbxPlotDensity.Value, "") > -1 _
        And Nz(chkNoCanopyVeg.Value, "") > -1 _
        And Nz(chkNoRootedVeg.Value, "") > -1 _
        And Nz(chkHasSocialTrails.Value, "") > -1 Then
        
        Select Case TempVars("ParkCode")
            Case "BLCA"
                'requires NoIndicatorSpecies
                If Nz(chkNoIndicatorSpecies.Value, "") > -1 Then GoTo Exit_Handler
            Case "CANY"
            Case "DINO"
        End Select
        
        isOK = True
        
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    btnSetObserverRecorder.Enabled = IIf(tbxID.Value > 0, True, False)
    
    'refresh form
    Me.Requery
   
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[VegPlot form])"
    End Select
    Resume Exit_Handler
End Sub
