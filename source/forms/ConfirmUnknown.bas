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
    ItemSuffix =84
    Left =3855
    Top =2430
    Right =28545
    Bottom =15015
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x236ab60a61c3e440
    End
    Caption ="Identify Unknown Species"
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
                    Caption ="title"
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
                    Left =2100
                    Top =1080
                    Width =1740
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblPlantType"
                    Caption ="Confirmed Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =2100
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1020
                    Top =840
                    Width =1020
                    Height =555
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUnknownCode"
                    Caption ="Unknown Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =840
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1395
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =900
                    Width =720
                    ForeColor =4210752
                    Name ="btnComment"
                    Caption ="comment"
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
                    ForeColor =16777184
                    Name ="lblContext"
                    Caption ="Context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3600
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7500
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
                    Top =480
                    Width =7500
                    Height =960
                    FontSize =14
                    LeftMargin =216
                    TopMargin =288
                    BackColor =13434828
                    Name ="lblIdentified"
                    Caption ="identification"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =480
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1440
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6720
                    Top =60
                    Width =720
                    TabIndex =1
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

                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
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
                    TabIndex =2
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
                    Left =5940
                    Top =60
                    Width =720
                    TabIndex =3
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

                    LayoutCachedLeft =5940
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
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
                    Left =7560
                    Top =105
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =4
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
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =960
                    Height =315
                    FontSize =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUnknownCode"
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                End
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =3000
                    Width =7650
                    Height =4380
                    TabIndex =5
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.UnknownList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3000
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =7380
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =2880
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =2880
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =7500
                    BackThemeColorIndex =-1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3540
                    Top =600
                    Width =2214
                    Height =315
                    TabIndex =6
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxIdentifiedBy"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Person who identified the unknown"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =3540
                    LayoutCachedTop =600
                    LayoutCachedWidth =5754
                    LayoutCachedHeight =915
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
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =2640
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
                    Caption ="message"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =2640
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =2955
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =2460
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
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2460
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =3060
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
                    Top =1500
                    Width =7500
                    Height =840
                    FontSize =14
                    LeftMargin =216
                    TopMargin =72
                    RightMargin =216
                    BottomMargin =72
                    BackColor =32768
                    ForeColor =16777215
                    Name ="lblConfirmed"
                    Caption ="confirmed"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1500
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =2340
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3540
                    Top =1020
                    Width =1860
                    Height =315
                    TabIndex =7
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxIdentifyDate"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Date species was identified"
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3540
                    LayoutCachedTop =1020
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1335
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
                    Left =2640
                    Top =1020
                    Width =720
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblIdentifyDate"
                    Caption ="Date"
                    GridlineColor =10921638
                    LayoutCachedLeft =2640
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1335
                End
                Begin Label
                    OverlapFlags =215
                    Left =2280
                    Top =1560
                    Width =1590
                    Height =315
                    BorderColor =8355711
                    ForeColor =15921906
                    Name ="lblConfirmedCode"
                    Caption ="Code"
                    GridlineColor =10921638
                    LayoutCachedLeft =2280
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =1875
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =2640
                    Top =600
                    Width =720
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblIdentifiedBy"
                    Caption ="By"
                    GridlineColor =10921638
                    LayoutCachedLeft =2640
                    LayoutCachedTop =600
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =915
                End
                Begin Label
                    OverlapFlags =215
                    Left =4620
                    Top =1560
                    Width =1590
                    Height =315
                    BorderColor =8355711
                    ForeColor =15921906
                    Name ="lblConfirmedSpecies"
                    Caption ="Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =4620
                    LayoutCachedTop =1560
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =1875
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ForeShade =95.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =3774
                    Height =315
                    FontSize =9
                    TabIndex =8
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x01000000a0000000020000000100000000000000000000001b00000001000000 ,
                        0x00000000fff2000000000000030000001c0000001f0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004d006f00640061006c00530065006400530069007a006500 ,
                        0x5d002e00560061006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxConfirmedSpecies"
                    RowSourceType ="Table/Query"
                    ControlTipText ="Confirmed species"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =5814
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
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =1980
                    Width =1680
                    Height =315
                    TabIndex =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =16051931
                    Name ="tbxConfirmedCode"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1980
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =2295
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =8
                    ForeTint =20.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =223
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4620
                    Top =1980
                    Width =2700
                    Height =315
                    TabIndex =10
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =16051931
                    Name ="tbxConfirmedSpecies"
                    ConditionalFormat = Begin
                        0x0100000096000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000170000001a0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4620
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =2295
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =8
                    ForeTint =20.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x740062007800420065006100720069006e0067005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =1980
                    Top =1980
                    Width =4860
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =16051931
                    Name ="lblNotIdentified"
                    GridlineColor =10921638
                    LayoutCachedLeft =1980
                    LayoutCachedTop =1980
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =2295
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =1
                    BorderTint =100.0
                    BorderShade =65.0
                    ForeThemeColorIndex =8
                    ForeTint =20.0
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
' Version:      1.01
' Basis:        Dropdown form
'
' Description:  Unknown form object related properties, Unknown, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, September 21, 2016
' References:   -
' Revisions:    BLC - 9/21/2016  - 1.00 - initial version
'               BLC - 1/24/2017  - 1.01 - revise to use GetContext()
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
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)
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
'   BLC - 1/24/1017 - revise to use GetContext()
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Unknown"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
    
        Dim strCaller As String
    
        'identify calling form
        Select Case Split(Me.OpenArgs, "|")(0)
            Case "WoodyCanopySpecies", "UnderstoryRootedSpecies", "AllRootedSpecies"
                strCaller = "VegPlot"
            Case Else
                strCaller = Split(Me.OpenArgs, "|")(0)
        End Select
        
        Me.CallingForm = strCaller

    End If
    
    'minimize calling form
    ToggleForm Me.CallingForm, -1

    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
                'Nz(TempVars("ParkCode"), "") & Space(2) & ">" & Space(2) & _
                'Nz(TempVars("River"), "") & Space(2) & ">" & Space(2) & _
                'Nz(TempVars("SiteCode"), "") & Space(2) & ">" & Space(2) & _
                'Nz(TempVars("Feature"), "")

    Title = "Identify/Confirm Unknown Species"
    lblTitle.Caption = ""
    Directions = "Enter the identification information for the unknown and click save."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    
    lblConfirmed.Caption = StringFromCodepoint(uCheckItem) 'uThumbsUp
    lblConfirmed.FontSize = 30
        
    btnSave.ControlTipText = "Identify/Confirm Unknown"
    
    'set hover
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnSave.Enabled = False
    tbxIdentifyDate.BackColor = lngYellow
  
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
    Set cbxConfirmedSpecies.Recordset = GetRecords("s_species_by_park")
    Set cbxIdentifiedBy.Recordset = GetRecords("s_contact_list")
        
    cbxConfirmedSpecies.BoundColumn = 1
    cbxConfirmedSpecies.ColumnCount = 3
    cbxConfirmedSpecies.ColumnWidths = "0;.8in;.7in"
    
    cbxIdentifiedBy.BoundColumn = 1
    cbxIdentifiedBy.ColumnCount = 2
    cbxIdentifiedBy.ColumnWidths = "0;.7in"

    'adjust confirmed section if not confirmed
    If IsDate(tbxIdentifyDate) Then
        'confirmed
        lblConfirmed.BackColor = lngSpringGreen
        lblConfirmed.ForeColor = lngWhite
        tbxConfirmedCode.Visible = True
        tbxConfirmedSpecies.Visible = True
        lblNotIdentified.Visible = True
    Else
        'not confirmed
        lblConfirmed.BackColor = lngGray50
        lblConfirmed.ForeColor = lngLtGray
        tbxConfirmedCode.Visible = False
        tbxConfirmedSpecies.Visible = False
        lblNotIdentified.Caption = "Unknown has not yet been identified"
        lblNotIdentified.Visible = True
    End If

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
'    tbxDescription.Value = ""
    
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
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    Dim loc As New Location
    
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
    
'    tbxDescription.ControlSource = ""
    
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
    DoCmd.OpenForm "Comment", acNormal, , , , , "Unknown|" & tbxID.Text
    
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

    'Forms("Main").Form.visible = True
    
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
'    If IsDate(tbxConfirmDate.Value) Then
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
