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
    ItemSuffix =67
    Left =6315
    Top =3780
    Right =24300
    Bottom =14775
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x236ab60a61c3e440
    End
    Caption ="VegPlot Sampling"
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
            Height =1380
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =60
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Events (Sampling Visits)"
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
                    Caption ="Enter the sampling start date."
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
                    Left =4260
                    Top =1065
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblModalSedSize"
                    Caption ="Modal Sediment Size"
                    GridlineColor =10921638
                    LayoutCachedLeft =4260
                    LayoutCachedTop =1065
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6660
                    Top =900
                    Width =720
                    ForeColor =4210752
                    Name ="btnComment"
                    Caption ="comment"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =900
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =1260
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
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
                    Left =2700
                    Top =1065
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDistance"
                    Caption ="Distance (m)"
                    GridlineColor =10921638
                    LayoutCachedLeft =2700
                    LayoutCachedTop =1065
                    LayoutCachedWidth =3945
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1065
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblNumber"
                    Caption ="Plot #"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1065
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7920
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
                    Caption ="cover"
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
                    Caption ="chk"
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
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =3
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
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4140
                    Top =60
                    Height =315
                    TabIndex =2
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxModalSedSize"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Size class"
                    ConditionalFormat = Begin
                        0x01000000a4000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001d000000200000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000000000220022000000 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
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
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =3420
                    Width =7650
                    Height =4380
                    TabIndex =4
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.VegPlotList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3420
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =7800
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =5
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
                    HoverColor =15060409
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
                    Top =3300
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3300
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =7920
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
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
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
                    Left =2580
                    Top =60
                    Height =315
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDistance"
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
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
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
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNumber"
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
                    LayoutCachedWidth =2460
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
                    TabIndex =7
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctFines"
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
                    TabIndex =8
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctWater"
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
                    TabIndex =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctURC"
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
                    TabIndex =10
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPlotDensity"
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
                    TabIndex =11
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPctFA"
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
                    TabIndex =12
                    BorderColor =10921638
                    Name ="chkNoCanopyVeg"
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
                    TabIndex =13
                    BorderColor =10921638
                    Name ="chkNoRootedVeg"
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
                    OverlapFlags =215
                    Left =3300
                    Top =2280
                    Width =360
                    Height =360
                    TabIndex =14
                    BorderColor =10921638
                    Name ="chkNoIndicatorSpecies"
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =2280
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =2640
                    Begin
                        Begin Label
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
                    TabIndex =15
                    BorderColor =10921638
                    Name ="chkHasSocialTrails"
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =16
                    BorderColor =8355711
                    ForeColor =690698
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
                    TabIndex =17
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
                    HoverColor =15060409
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
                    TabIndex =18
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
                    HoverColor =15060409
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
                    Top =2640
                    Width =720
                    TabIndex =19
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
                    HoverColor =15060409
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
                    Caption ="tag"
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
                    TabIndex =20
                    ForeColor =4210752
                    Name ="btnTaglines"
                    Caption ="Tagline"
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
                    HoverColor =15060409
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
' Form:         Location
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  List form object related properties, Location, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_ButtonCaption
Private m_SelectedID As Integer
Private m_SelectedValue As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)

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

Public Property Let ButtonCaption(Value As String)
    If Len(Value) > 0 Then
        m_ButtonCaption = Value

        'set the form button caption
        Me.btnSave.Caption = m_ButtonCaption
    Else
        RaiseEvent InvalidCaption(Value)
    End If
End Property

Public Property Get ButtonCaption() As String
    ButtonCaption = m_ButtonCaption
End Property

Public Property Let SelectedID(Value As Integer)
        m_SelectedID = Value
End Property

Public Property Get SelectedID() As Integer
    SelectedID = m_SelectedID
End Property

Public Property Let SelectedValue(Value As String)
        m_SelectedValue = Value
End Property

Public Property Get SelectedValue() As String
    SelectedValue = m_SelectedValue
End Property

'---------------------
' Methods
'---------------------

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
            "Error encountered (#" & Err.Number & " - Form_Load[Location form])"
    End Select
    Resume Exit_Handler
End Sub

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
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Title = "VegPlot (Sampling)"
    Directions = "Enter the plot information and click save." _
                & vbCrLf & "Add cover species via buttons at right."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.forecolor = lngLtBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.forecolor = lngBlue
    lblChkboxes.Caption = StringFromCodepoint(uCheck)
    lblCover.Caption = "Cover" & vbCrLf & "Species"
    lblTagline.Caption = StringFromCodepoint(uRuler)
    'Me.btnTaglines.Caption = StringFromCodepoint(uTag)
    
    'set hover
    btnComment.hoverColor = lngGreen
    btnTaglines.hoverColor = lngGreen
    btnWCC.hoverColor = lngGreen
    btnURC.hoverColor = lngGreen
    btnARC.hoverColor = lngGreen
    btnSave.hoverColor = lngGreen
    btnUndo.hoverColor = lngGreen
      
    'defaults
    tbxIcon.forecolor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    tbxNumber.backcolor = lngYellow
    tbxDistance.backcolor = lngYellow
    tbxModalSedSize.backcolor = lngYellow
    tbxPctFines.backcolor = lngYellow
    tbxPctWater.backcolor = lngYellow
    tbxPctURC.backcolor = lngYellow
    tbxPlotDensity.backcolor = lngYellow
    
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0
  
    'defaults --> turn off items
    lblWater.visible = False
    tbxPctWater.visible = False
    lblURC.visible = False
    tbxPctURC.visible = False
    lblPlotDensity.visible = False
    lblFraction.visible = False
    tbxPlotDensity.visible = False
    chkNoCanopyVeg.visible = False
    lblNoCanopyVeg.visible = False
    chkNoIndicatorSpecies.visible = False
    lblNoIndicatorSpecies.visible = False
    chkNoRootedVeg.visible = False
    lblNoRootedVeg.visible = False
    chkHasSocialTrails.visible = False
    lblHasSocialTrails.visible = False
    btnWCC.visible = False
    btnURC.visible = False
    btnARC.visible = False
    
    'adjust UI based on park
    Select Case TempVars("ParkCode")
        Case "BLCA"
            lblURC.visible = True
            tbxPctURC.visible = True
            lblPlotDensity.visible = True
            lblFraction.visible = True
            tbxPlotDensity.visible = True
            chkNoCanopyVeg.visible = True
            lblNoCanopyVeg.visible = True
            chkNoIndicatorSpecies.visible = True
            lblNoIndicatorSpecies.visible = True
            chkHasSocialTrails.visible = True
            lblHasSocialTrails.visible = True
            btnWCC.visible = True
            btnURC.visible = True
        
        Case "CANY"
            lblURC.visible = True
            tbxPctURC.visible = True
            lblWater.visible = True
            tbxPctWater.visible = True
            lblPlotDensity.visible = True
            lblFraction.visible = True
            tbxPlotDensity.visible = True
            tbxPctURC.visible = True
            chkNoCanopyVeg.visible = True
            lblNoCanopyVeg.visible = True
            chkNoRootedVeg.visible = True
            lblNoRootedVeg.visible = True
            chkHasSocialTrails.visible = True
            lblHasSocialTrails.visible = True
            btnWCC.visible = True
            
        Case "DINO"
            lblWater.visible = True
            tbxPctWater.visible = True
            chkNoRootedVeg.visible = True
            lblNoRootedVeg.visible = True
            btnARC.visible = True
    End Select
  
    'hide modal Main form
    Forms("Main").visible = False
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Location form])"
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
              
      If tbxID > 0 Then btnComment.Enabled = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Location form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxStartDate_Change
' Description:  Dropdown change actions
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
Private Sub tbxStartDate_Change()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxStartDate_Change[Location form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxStartDate_LostFocus
' Description:  Dropdown change actions
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
Private Sub tbxStartDate_LostFocus()
On Error GoTo Err_Handler

    'ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxStartDate_LostFocus[Location form])"
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
' ---------------------------------
Private Sub btnUndo_Click()
On Error GoTo Err_Handler
    
    btnSave.Enabled = False
    
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUndo_Click[Location form])"
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
'        .LocationName = tbxName.Value
'        .LocationType = "" 'cbxLocationType.SelText
'
'        .HeadtoOrientDistance = tbxDistance.Value
'        .HeadtoOrientBearing = tbxBearing.Value
        
        .ID = tbxID.Value '0 if new, edit if > 0
        .SaveToDb
    End With
    
    'clear values & refresh display
    Me.RecordSource = ""
    
'    tbxDistance.ControlSource = ""
'    tbxBearing.ControlSource = ""
'    tbxNotes.ControlSource = ""
    
    tbxID.ControlSource = ""
    tbxID.Value = 0
    
    ReadyForSave
    
    'refresh list
    Me.list.Requery
    
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[Location form])"
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
    
    'open comment form
    DoCmd.OpenForm "SpeciesList", acNormal, , , , , "WoodyCanopySpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnWCC_Click[Location form])"
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
    
    'open comment form
    DoCmd.OpenForm "SpeciesList", acNormal, , , , , "UnderstoryRootedSpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnURC_Click[Location form])"
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
    
    'open comment form
    DoCmd.OpenForm "SpeciesList", acNormal, , , , , "AllRootedSpecies|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnARC_Click[Location form])"
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
    DoCmd.OpenForm "Comment", acNormal, , , , , "location|" & tbxID.Text
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[Location form])"
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
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    Forms("Main").Form.visible = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[Location form])"
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
' ---------------------------------
Private Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
'    If tbxDistance.Value > 0 And tbxBearing.Value <> "" Then
        isOK = True
'    End If
    
    tbxIcon.forecolor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[Location form])"
    End Select
    Resume Exit_Handler
End Sub
