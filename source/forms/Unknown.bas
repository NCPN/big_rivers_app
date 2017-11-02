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
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10140
    DatasheetFontHeight =11
    ItemSuffix =81
    Left =3360
    Top =2730
    Right =13755
    Bottom =16275
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x236ab60a61c3e440
    End
    Caption ="Events (Sampling Visits)"
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
                    Caption ="Unknown Species"
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
                    Caption ="description"
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
                    Left =4320
                    Top =1080
                    Width =1740
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPlantType"
                    Caption ="Plant Type"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1080
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1080
                    Width =1890
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUnknownCode"
                    Caption ="Unknown Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9240
                    Top =960
                    Width =720
                    ForeColor =4210752
                    Name ="btnComment"
                    Caption ="comment"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =9240
                    LayoutCachedTop =960
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =1320
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
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5820
                    Top =60
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777184
                    Name ="lblContext"
                    Caption ="Context"
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8340
                    Top =960
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnConfirmUnknown"
                    Caption ="Identify Unknown"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Identify/Confirm Unknown"
                    GridlineColor =10921638

                    LayoutCachedLeft =8340
                    LayoutCachedTop =960
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =1320
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
                    OverlapFlags =85
                    Left =6360
                    Top =960
                    Width =1800
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnSpeciesSearch"
                    Caption ="í ½í´Ž  Species"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Lookup species name"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =960
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =1320
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
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =12120
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
                    Left =180
                    Top =5940
                    Width =9840
                    Height =1020
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =16769505
                    Name ="lblCollection"
                    Caption ="collected?"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =5940
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =6960
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8760
                    Top =60
                    Width =720
                    TabIndex =2
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

                    LayoutCachedLeft =8760
                    LayoutCachedTop =60
                    LayoutCachedWidth =9480
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
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =75
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =75
                    LayoutCachedWidth =960
                    LayoutCachedHeight =375
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7860
                    Top =60
                    Width =720
                    TabIndex =4
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

                    LayoutCachedLeft =7860
                    LayoutCachedTop =60
                    LayoutCachedWidth =8580
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
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9660
                    Top =105
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="tbxID"
                    GridlineColor =10921638

                    LayoutCachedLeft =9660
                    LayoutCachedTop =105
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =405
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =180
                    Top =1620
                    Width =4620
                    Height =720
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDescription"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter plant general description"
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =2340
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =1320
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDescription"
                            Caption ="Description:"
                            GridlineColor =10921638
                            LayoutCachedLeft =60
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1170
                    Top =60
                    Width =2850
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUnknownCode"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000009e000000020000000100000000000000000000001a00000001000000 ,
                        0x00000000fff2000000000000030000001b0000001e0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074006200780055006e006b006e006f0077006e0043006f00640065005d00 ,
                        0x2e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1170
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000190000005b00 ,
                        0x74006200780055006e006b006e006f0077006e0043006f00640065005d002e00 ,
                        0x560061006c00750065003d002200220000000000000000000000000000000000 ,
                        0x000000000000000000030000000100000000000000ffffff0002000000220022 ,
                        0x00000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =2940
                    Top =4200
                    Width =2340
                    Height =1680
                    FontSize =14
                    LeftMargin =72
                    TopMargin =72
                    BackColor =12444887
                    Name ="lblFlower"
                    Caption ="flower"
                    ControlTipText ="Add/Edit flower charateristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =2940
                    LayoutCachedTop =4200
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =5880
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
                    Left =4980
                    Top =2445
                    Width =5040
                    Height =1680
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =11916796
                    Name ="lblLeaf"
                    Caption ="leaf"
                    GridlineColor =10921638
                    LayoutCachedLeft =4980
                    LayoutCachedTop =2445
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =4125
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
                    Top =2445
                    Width =4620
                    Height =1680
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =16051931
                    Name ="lblGrass"
                    Caption ="grass"
                    ControlTipText ="Add/Edit grass characteristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =2445
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =4125
                    BackThemeColorIndex =8
                    BackTint =20.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =7620
                    Width =9855
                    Height =4380
                    TabIndex =6
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.UnknownList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =7620
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =12000
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =7500
                    Width =10140
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =7500
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =12120
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =5805
                    Top =2565
                    Width =525
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLeafType"
                    Caption ="Type"
                    GridlineColor =10921638
                    LayoutCachedLeft =5805
                    LayoutCachedTop =2565
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =2880
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =3600
                    Top =6600
                    Width =360
                    Height =360
                    TabIndex =7
                    BorderColor =10921638
                    Name ="chkHasPhotos"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =3600
                    LayoutCachedTop =6600
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =6960
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3840
                            Top =6540
                            Width =1485
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblHasPhotos"
                            Caption ="Photos Taken?"
                            GridlineColor =10921638
                            LayoutCachedLeft =3840
                            LayoutCachedTop =6540
                            LayoutCachedWidth =5325
                            LayoutCachedHeight =6855
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =1980
                    Top =6600
                    Width =360
                    Height =360
                    TabIndex =8
                    BorderColor =10921638
                    Name ="chkCollected"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Was plant collected?"
                    GridlineColor =10921638

                    LayoutCachedLeft =1980
                    LayoutCachedTop =6600
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =6960
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2220
                            Top =6540
                            Width =1065
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCollected"
                            Caption ="Collected?"
                            ControlTipText ="Was plant collected?"
                            GridlineColor =10921638
                            LayoutCachedLeft =2220
                            LayoutCachedTop =6540
                            LayoutCachedWidth =3285
                            LayoutCachedHeight =6855
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =4200
                    Width =2700
                    Height =1680
                    FontSize =20
                    LeftMargin =72
                    TopMargin =72
                    BackColor =12835293
                    Name ="lblStem"
                    Caption ="stem"
                    ControlTipText ="Add/Edit stem characteristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =4200
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =5880
                    BackThemeColorIndex =3
                    BackShade =90.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6900
                    Top =6540
                    Width =2274
                    Height =315
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="cbxCollectedBy"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Person who collected the plant"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =6900
                    LayoutCachedTop =6540
                    LayoutCachedWidth =9174
                    LayoutCachedHeight =6855
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =300
                    Top =3345
                    Width =600
                    Height =420
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintPct"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =3345
                    LayoutCachedWidth =900
                    LayoutCachedHeight =3765
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =7260
                    Width =10140
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsg"
                    Caption ="message"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =7260
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =7575
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =5760
                    Top =7080
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsgIcon"
                    Caption ="icon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =7080
                    LayoutCachedWidth =6585
                    LayoutCachedHeight =7680
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =180
                    Top =510
                    Width =9840
                    Height =690
                    FontSize =14
                    LeftMargin =216
                    RightMargin =216
                    BackColor =32768
                    ForeColor =16777215
                    Name ="lblConfirmed"
                    Caption ="confirmed"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =510
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =1200
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3870
                    Top =6120
                    Width =5310
                    Height =315
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxBestGuess"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Best guess species"
                    GridlineColor =10921638

                    LayoutCachedLeft =3870
                    LayoutCachedTop =6120
                    LayoutCachedWidth =9180
                    LayoutCachedHeight =6435
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =1935
                    Top =6135
                    Width =1845
                    Height =240
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblBestGuess"
                    Caption ="Best Guess Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =1935
                    LayoutCachedTop =6135
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =6375
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2520
                    Top =720
                    Width =2280
                    Height =315
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxConfirmedCode"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Confirmed species lookup code"
                    GridlineColor =10921638

                    LayoutCachedLeft =2520
                    LayoutCachedTop =720
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =1035
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =1860
                    Top =720
                    Width =540
                    Height =315
                    BorderColor =8355711
                    ForeColor =15921906
                    Name ="lblConfirmedCode"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =1860
                    LayoutCachedTop =720
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =1035
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4980
                    Top =1620
                    Width =5040
                    Height =720
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSalientFeature"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter plant most salient feature"
                    GridlineColor =10921638

                    LayoutCachedLeft =4980
                    LayoutCachedTop =1620
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =2340
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4860
                            Top =1320
                            Width =1545
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblSalientFeature"
                            Caption ="Salient Feature:"
                            GridlineColor =10921638
                            LayoutCachedLeft =4860
                            LayoutCachedTop =1320
                            LayoutCachedWidth =6405
                            LayoutCachedHeight =1635
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =60
                    Width =1794
                    Height =315
                    TabIndex =13
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001800000001000000 ,
                        0x00000000fff200000000000003000000190000001c0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0063006200780050006c0061006e00740054007900700065005d002e005600 ,
                        0x61006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxPlantType"
                    RowSourceType ="Value List"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Type of plant"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =4320
                    LayoutCachedTop =60
                    LayoutCachedWidth =6114
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000170000005b00 ,
                        0x63006200780050006c0061006e00740054007900700065005d002e0056006100 ,
                        0x6c00750065003d00220022000000000000000000000000000000000000000000 ,
                        0x0000000000030000000100000000000000ffffff000200000022002200000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =1260
                    Top =3465
                    Width =3366
                    Height =555
                    TabIndex =14
                    BorderColor =10921638
                    Name ="Frame51"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3465
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =4020
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =1380
                            Top =3345
                            Width =2010
                            Height =315
                            BackColor =16051931
                            BorderColor =8355711
                            ForeColor =477336
                            Name ="lblOptgPerennialGrassType"
                            Caption ="Perennial Grass Type"
                            GridlineColor =10921638
                            LayoutCachedLeft =1380
                            LayoutCachedTop =3345
                            LayoutCachedWidth =3390
                            LayoutCachedHeight =3660
                            BackThemeColorIndex =8
                            BackTint =20.0
                            ForeThemeColorIndex =9
                            ForeTint =100.0
                            ForeShade =50.0
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =215
                            Left =1446
                            Top =3703
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optBunchgrass"
                            GridlineColor =10921638

                            LayoutCachedLeft =1446
                            LayoutCachedTop =3703
                            LayoutCachedWidth =1706
                            LayoutCachedHeight =3943
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1676
                                    Top =3675
                                    Width =1110
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblBunchgrass"
                                    Caption ="Bunchgrass"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1676
                                    LayoutCachedTop =3675
                                    LayoutCachedWidth =2786
                                    LayoutCachedHeight =3990
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =215
                            Left =3000
                            Top =3703
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optRhizomatous"
                            GridlineColor =10921638

                            LayoutCachedLeft =3000
                            LayoutCachedTop =3703
                            LayoutCachedWidth =3260
                            LayoutCachedHeight =3943
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3230
                                    Top =3675
                                    Width =1275
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblRhizomatous"
                                    Caption ="Rhizomatous"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =3230
                                    LayoutCachedTop =3675
                                    LayoutCachedWidth =4505
                                    LayoutCachedHeight =3990
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =215
                    Left =1260
                    Top =2685
                    Width =3366
                    Height =598
                    TabIndex =15
                    BorderColor =10921638
                    Name ="optgGrassType"
                    GridlineColor =10921638

                    LayoutCachedLeft =1260
                    LayoutCachedTop =2685
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =3283
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =1380
                            Top =2565
                            Width =2010
                            Height =315
                            BackColor =16051931
                            BorderColor =8355711
                            ForeColor =477336
                            Name ="lblForbGrassTypes"
                            Caption ="Forb/Grass Type"
                            GridlineColor =10921638
                            LayoutCachedLeft =1380
                            LayoutCachedTop =2565
                            LayoutCachedWidth =3390
                            LayoutCachedHeight =2880
                            BackThemeColorIndex =8
                            BackTint =20.0
                            ForeThemeColorIndex =9
                            ForeTint =100.0
                            ForeShade =50.0
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =215
                            Left =1446
                            Top =2923
                            OptionValue =1
                            BorderColor =10921638
                            Name ="optAnnual"
                            GridlineColor =10921638

                            LayoutCachedLeft =1446
                            LayoutCachedTop =2923
                            LayoutCachedWidth =1706
                            LayoutCachedHeight =3163
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1676
                                    Top =2895
                                    Width =1110
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblAnnual"
                                    Caption ="Annual"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =1676
                                    LayoutCachedTop =2895
                                    LayoutCachedWidth =2786
                                    LayoutCachedHeight =3210
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =215
                            Left =3000
                            Top =2923
                            TabIndex =1
                            OptionValue =2
                            BorderColor =10921638
                            Name ="optPerennial"
                            GridlineColor =10921638

                            LayoutCachedLeft =3000
                            LayoutCachedTop =2923
                            LayoutCachedWidth =3260
                            LayoutCachedHeight =3163
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3230
                                    Top =2895
                                    Width =1275
                                    Height =315
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="lblPerennial"
                                    Caption ="Perennial"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =3230
                                    LayoutCachedTop =2895
                                    LayoutCachedWidth =4505
                                    LayoutCachedHeight =3210
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6660
                    Top =2565
                    Width =3285
                    Height =315
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLeafType"
                    ControlTipText ="Leaf type"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =2565
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =2880
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =5580
                    Top =6540
                    Width =1230
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCollectedBy"
                    Caption ="Collected By"
                    GridlineColor =10921638
                    LayoutCachedLeft =5580
                    LayoutCachedTop =6540
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =6855
                End
                Begin Label
                    OverlapFlags =215
                    Left =5805
                    Top =3015
                    Width =720
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLeafMargin"
                    Caption ="Margin"
                    GridlineColor =10921638
                    LayoutCachedLeft =5805
                    LayoutCachedTop =3015
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =3330
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6660
                    Top =3015
                    Width =3285
                    Height =315
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLeafMargin"
                    ControlTipText ="Leaf margin"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =3015
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =3330
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =223
                    Left =5160
                    Top =3465
                    Width =1500
                    Height =540
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label68"
                    Caption ="Other Leaf Characteristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =5160
                    LayoutCachedTop =3465
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =4005
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6660
                    Top =3480
                    Width =3285
                    Height =525
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLeafOtherCharacteristics"
                    ControlTipText ="Leaf - other characteristics"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =3480
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =4005
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =223
                    Left =1320
                    Top =4260
                    Width =1425
                    Height =585
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblStemCharacteristics"
                    Caption ="Stem Characteristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =4260
                    LayoutCachedWidth =2745
                    LayoutCachedHeight =4845
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =300
                    Top =4800
                    Width =2460
                    Height =960
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxStemCharacteristics"
                    ControlTipText ="Stem characteristics"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =4800
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =5760
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =3780
                    Top =4260
                    Width =1425
                    Height =585
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFlowerCharacteristics"
                    Caption ="Flower Characteristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =3780
                    LayoutCachedTop =4260
                    LayoutCachedWidth =5205
                    LayoutCachedHeight =4845
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3060
                    Top =4860
                    Width =2100
                    Height =900
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFlowerCharacteristics"
                    ControlTipText ="Flower characteristics"
                    GridlineColor =10921638

                    LayoutCachedLeft =3060
                    LayoutCachedTop =4860
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =5760
                    BackThemeColorIndex =-1
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    Left =5340
                    Top =4200
                    Width =4680
                    Height =1680
                    FontSize =14
                    LeftMargin =144
                    TopMargin =144
                    BackColor =9950949
                    Name ="lblMicroHabitat"
                    Caption ="habitat"
                    GridlineColor =10921638
                    LayoutCachedLeft =5340
                    LayoutCachedTop =4200
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =5880
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =6540
                    Top =4290
                    Width =2880
                    Height =510
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblGeneralCharacteristics"
                    Caption ="General && Microhabitat Characteristics"
                    GridlineColor =10921638
                    LayoutCachedLeft =6540
                    LayoutCachedTop =4290
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =4800
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5460
                    Top =4860
                    Width =4485
                    Height =900
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxMicroHabitat"
                    ControlTipText ="General && microhabitat characteristics"
                    GridlineColor =10921638

                    LayoutCachedLeft =5460
                    LayoutCachedTop =4860
                    LayoutCachedWidth =9945
                    LayoutCachedHeight =5760
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    Left =4980
                    Top =720
                    Width =735
                    Height =315
                    BorderColor =8355711
                    ForeColor =15921906
                    Name ="lblConfirmedSpecies"
                    Caption ="Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =4980
                    LayoutCachedTop =720
                    LayoutCachedWidth =5715
                    LayoutCachedHeight =1035
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5880
                    Top =720
                    Width =3894
                    Height =315
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="cbxConfirmedSpecies"
                    RowSourceType ="Table/Query"
                    ControlTipText ="Confirmed species"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =5880
                    LayoutCachedTop =720
                    LayoutCachedWidth =9774
                    LayoutCachedHeight =1035
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
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
' Form:         Unknown
' Level:        Application form
' Version:      1.08
' Basis:        Dropdown form
'
' Description:  Unknown form object related properties, Unknown, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, July 5, 2016
' References:   -
' Revisions:    BLC - 7/5/2016  - 1.00 - initial version
'               BLC - 8/2/2016  - 1.01 - use Me.CallingForm
'               BLC - 8/23/2016 - 1.02 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 10/25/2016 - 1.03 - removed ButtonCaption, SeleectedID,
'                                        SelectedValue properties
'               BLC - 1/24/2017 - 1.04 - adjust to use GetContext()
'               BLC - 9/25/2017 - 1.05 - revise for NCPN_framework.Location class
'               BLC - 9/27/2017 - 1.06 - update to use Factory.NewClassXX() vs GetClass()
'               BLC - 10/2/2017 - 1.07 - add btnSpeciesSearch_Click()
'               BLC - 10/19/2017 - 1.08 - added comment length
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
    If Len(Value) > 0 Then
        m_CallingForm = Value
    Else
        RaiseEvent InvalidCallingForm(Value)
    End If
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
'   BLC - 1/24/2017 - adjust to use GetContext()
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Main"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then Me.CallingForm = Me.OpenArgs

    'minimize Calling Form
    ToggleForm Me.CallingForm, -1

    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
                'Nz(TempVars("ParkCode"), "") & Space(2) & ">" & Space(2) & _
                'Nz(TempVars("River"), "") & Space(2) & ">" & Space(2) & _
                'Nz(TempVars("SiteCode"), "") & Space(2) & ">" & Space(2) & _
                'Nz(TempVars("Feature"), "")

    Title = "Unknown Species"
    lblTitle.Caption = ""
    Directions = "Enter the unknown information and click save."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    
    lblLeaf.Caption = StringFromCodepoint(uLeafFallen)
    lblLeaf.FontSize = 20
    lblGrass.Caption = StringFromCodepoint(uGrass)
    lblGrass.FontSize = 20
    lblStem.Caption = StringFromCodepoint(uHerb)
    lblFlower.Caption = StringFromCodepoint(uFloretteWhite)
    lblFlower.FontSize = 24
    lblConfirmed.Caption = StringFromCodepoint(uCheckItem) 'uThumbsUp
    lblConfirmed.FontSize = 30
    
    btnConfirmUnknown.Caption = StringFromCodepoint(uKey)
    
    'set hover
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
    btnConfirmUnknown.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnSave.Enabled = False
    tbxUnknownCode.BackColor = lngYellow
    cbxPlantType.BackColor = lngYellow
  
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0

'Veg unknowns
'Public Const PLANT_TYPES As String = "herb,shrub,tree,grass,sedge,other"  'TEXT(50) --> TEXT(15)
'Public Const LEAF_TYPES As String = "compound/simple, arrangement" 'TEXT(50)
'Public Const FORB_GRASS_TYPES As String = "Annual,Perennial" 'TEXT(10)
'Public Const PERENNIAL_GRASS_TYPES As String = "Bunchgrass, Rhizomatous" 'TEXT(15)
'Salient feature TEXT(255)
'Leaf margin TEXT(50)
'Other leaf characteristics:  pubescence, sap, stipules TEXT(50)
'Stem characteristics: shape, pubescence, bud TEXT(50)
'Flower characteristics: color location floral formula TEXT(50)
'General and microhabitat characteristics TEXT(50)
'Perennial grass type: Bunchgrass or Rhizomatous TEXT(15)
'Collection method TEXT(50)
  
    'populate dropdowns
    Set cbxCollectedBy.Recordset = GetRecords("s_contact_list")
    
    cbxCollectedBy.BoundColumn = 1
    cbxCollectedBy.ColumnCount = 2
    cbxCollectedBy.ColumnWidths = "0;.7in"
    
    cbxPlantType.RowSource = Replace(PLANT_TYPES, ",", ";")
  
    Set cbxConfirmedSpecies.Recordset = GetRecords("s_species_by_park")
  
    cbxConfirmedSpecies.BoundColumn = 1 'bind to label (not ID)
    cbxConfirmedSpecies.ColumnCount = 5
    cbxConfirmedSpecies.ColumnHeads = True
    cbxConfirmedSpecies.ColumnWidths = "0;.7in;.2in;0;0" 'display the display column (combines label - summary)
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[Unknown form])"
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
' Source/date:  Bonnie Campbell, July 5, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/5/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[Unknown form])"
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

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[Unknown form])"
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
            "Error encountered (#" & Err.Number & " - tbxStartDate_Change[Unknown form])"
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
            "Error encountered (#" & Err.Number & " - tbxStartDate_LostFocus[Unknown form])"
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
    
    'clear values
'    tbxName.Value = ""
    tbxDescription.Value = ""
    
    btnSave.Enabled = False
    
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUndo_Click[Unknown form])"
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
' Source/date:  Bonnie Campbell, July 5, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/5/2016 - initial version
'   BLC - 9/25/2017 - revise for NCPN_framework.Location class
'   BLC - 9/27/2017 - update to use Factory.NewClassXX() vs GetClass()
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'Dim loc As New Location
    Dim loc As NCPN_framework.Location
    Set loc = Factory.NewLocation
    
    With loc
        'values passed into form
        .CollectionSourceName = "T"
        
        .CreateDate = ""
        .CreatedByID = 0
        .LastModified = ""
        .LastModifiedByID = 0
        
        '.ProtocolID = 1
        '.SiteID = 1
        
        'form values
'        .UnknownName = tbxName.Value
'        .UnknownType = "" 'cbxUnknownType.SelText
'
'        .HeadtoOrientDistance = tbxDistance.Value
'        .HeadtoOrientBearing = tbxBearing.Value
        
        .ID = tbxID.Value '0 if new, edit if > 0
        .SaveToDb
    End With
    
    'clear values & refresh display
    Me.RecordSource = ""
    
    tbxDescription.ControlSource = ""
    
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
            "Error encountered (#" & Err.Number & " - btnSave_Click[Unknown form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnConfirmUnknown_Click
' Description:  Undo button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 21, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/21/2016 - initial version
' ---------------------------------
Private Sub btnConfirmUnknown_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "ConfirmUnknown", acNormal, , , , , "Unknown|" & tbxID.Value
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnConfirmUnknown_Click[Unknown form])"
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
'   BLC - 10/19/2017 - added comment length
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "Unknown|" & tbxID.Text & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[Unknown form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSpeciesSearch_Click
' Description:  Woody Canopy Cover button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' SoSpeciesSearche/date:  Bonnie Campbell, August 2, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/2/2016 - initial version
' ---------------------------------
Private Sub btnSpeciesSearch_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "SpeciesSearch", acNormal, , , , , Me.Name
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSpeciesSearch_Click[Unknown form])"
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
            "Error encountered (#" & Err.Number & " - Form_Close[Unknown form])"
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
' Source/date:  Bonnie Campbell, July 5, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/5/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
'    If tbxDistance.Value > 0 And tbxBearing.Value <> "" Then
        isOK = True
'    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    btnSave.Enabled = isOK
    
    'refresh form
    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[Unknown form])"
    End Select
    Resume Exit_Handler
End Sub
