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
    ItemSuffix =34
    Left =3855
    Top =2430
    Right =24030
    Bottom =15015
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
            Height =1695
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
                    Caption ="Features"
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
                    Caption ="directions"
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
                    Left =1260
                    Top =1380
                    Width =1245
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFeature"
                    Caption ="Feature"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =1380
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =1695
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
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =3660
                    Top =60
                    Width =4140
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblContext"
                    Caption ="Context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3660
                    LayoutCachedTop =60
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =300
                    Top =900
                    Width =855
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLocation"
                    Caption ="Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =900
                    LayoutCachedWidth =1155
                    LayoutCachedHeight =1215
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =1200
                    Top =900
                    Width =3414
                    Height =315
                    TabIndex =1
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
                    Name ="cbxLocation"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;0;0;1"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Transect location"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1200
                    LayoutCachedTop =900
                    LayoutCachedWidth =4614
                    LayoutCachedHeight =1215
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =4740
                    Top =900
                    Width =1680
                    TabIndex =2
                    ForeColor =16711680
                    Name ="btnAddLocation"
                    Caption =" Add Location"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new location"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000009070603090706080 ,
                        0x906850c0806850f0806050f0805840c070584080705040300000000000000000 ,
                        0x0000000000000000000000000000000000000000a0786060907060e0907050ff ,
                        0x706840ff906840ffb06840ff906040ff805040ff705040e07050406000000000 ,
                        0x00000000000000000000000000000000a0806060a07860f0c08060ff807830ff ,
                        0x607820ff506040ff906040ffc06840ffb06040ff905840ff705040f070504060 ,
                        0x000000000000000000000000b0807030a07860e0d08860fff08860ffd08040ff ,
                        0x608820ff308820ff606830ff706840ffb06840ffb06040ff905840ff705040e0 ,
                        0x705040300000000000000000b0807080c08870fff09860fff09060ffc08860ff ,
                        0x508040ff409820ff408830ff707040ffb06840ff806040ff606040ff605840ff ,
                        0x705840800000000000000000b08070c0e09870ffffa070ffe09070ff707850ff ,
                        0x50a840ff50a830ff809830ffd07850ff707040ff506840ff307810ff606830ff ,
                        0x805840c00000000000000000b08870f0f0b890fff0a880ff607860ff609850ff ,
                        0x60b040ff90a040ffe08860ffd08050ff507050ff408830ff309010ff408020ff ,
                        0x806050f00000000000000000b08870f0ffb8a0ffa09880ff40a850ff60b850ff ,
                        0x909860ffe09870ffe09060ffe08860ff60a030ff40a020ff409820ff408820ff ,
                        0x806850f00000000000000000b08870c0e0b8a0ff90b880ff50c870ff50c870ff ,
                        0x80a070ffc0a070ffd09870ffe09060ffa09840ff50a830ff40a020ff508830ff ,
                        0x906850c00000000000000000b0907080b0a080ff90d0a0ff70e090ff90e0a0ff ,
                        0x80e8a0ffa0d090ffb0a870fff09870ffe09060ff50a830ff50a830ff708840ff ,
                        0x907060800000000000000000b0907030b09080e090c8a0ffa0e8b0ffd0f0c0ff ,
                        0xd0f8d0fff0f8d0ffb0e8b0ff809860ffd09840ffb09040ff609840ff907050e0 ,
                        0x90706030000000000000000000000000b0908070b0a080f0b0d8b0ffd0f0d0ff ,
                        0xe0f8d0ff80b070ff80d890ff50b070ff908840ffd08860ffa07860f0a0786060 ,
                        0x0000000000000000000000000000000000000000b0908070b09080e0c0c0a0ff ,
                        0xd0e0c0ffb0e8b0ff80e0a0ff60b870ffa08850ffa07860e0a080606000000000 ,
                        0x000000000000000000000000000000000000000000000000b0908030b0908080 ,
                        0xb09080c0b09080f0c09880f0b09080c0b0807080b08070300000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4740
                    LayoutCachedTop =900
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =1260
                    PictureCaptionArrangement =5
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
            Height =9240
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6660
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
                    Left =1320
                    Top =60
                    Width =1860
                    Height =315
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFeature"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000180000001b0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074006200780046006500610074007500720065005d002e00560061006c00 ,
                        0x750065003d002200220000000000000022002200000000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000150000005b00 ,
                        0x74006200780046006500610074007500720065005d002e00560061006c007500 ,
                        0x65003d0022002200000000000000000000000000000000000000000000000000 ,
                        0x00030000000100000000000000ffffff00020000002200220000000000000000 ,
                        0x0000000000000000000000000000
                    End
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
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =2
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
                Begin Subform
                    CanShrink = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =4740
                    Width =7650
                    Height =4380
                    TabIndex =4
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.FeatureList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =4740
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =9120
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =4620
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =4620
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =9240
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
                    TabIndex =5
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
                    Left =420
                    Top =2820
                    Width =7140
                    Height =1500
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDescription"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter feature description as needed."
                    GridlineColor =10921638

                    LayoutCachedLeft =420
                    LayoutCachedTop =2820
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =4320
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =2400
                            Width =1200
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblDescription"
                            Caption ="Description:"
                            GridlineColor =10921638
                            LayoutCachedLeft =240
                            LayoutCachedTop =2400
                            LayoutCachedWidth =1440
                            LayoutCachedHeight =2715
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =375
                    Top =990
                    Width =7140
                    Height =1320
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFeatureDirections"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Enter directions to the feature."
                    GridlineColor =10921638

                    LayoutCachedLeft =375
                    LayoutCachedTop =990
                    LayoutCachedWidth =7515
                    LayoutCachedHeight =2310
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =195
                            Top =570
                            Width =1080
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblFeatureDirections"
                            Caption ="Directions:"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =570
                            LayoutCachedWidth =1275
                            LayoutCachedHeight =885
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =4380
                    Width =7860
                    Height =315
                    LeftMargin =360
                    TopMargin =7
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsg"
                    Caption ="message"
                    GridlineColor =10921638
                    LayoutCachedTop =4380
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =4695
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Top =4380
                    Width =660
                    Height =315
                    TopMargin =7
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsgIcon"
                    Caption ="icon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =4380
                    LayoutCachedWidth =660
                    LayoutCachedHeight =4695
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
' Form:         Feature
' Level:        Application form
' Version:      1.04
' Basis:        Dropdown form
'
' Description:  Feature form object related properties, Feature, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, June 27, 2016
' References:   -
' Revisions:    BLC - 6/27/2016 - 1.00 - initial version
'               BLC - 8/23/2016 - 1.01 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 10/20/2016 - 1.02 - added calling form property, removed buttoncaption,
'                                         selectedID, selectedValue properties
'               BLC - 1/10/2017 - 1.03 - revised to use GetRecords() vs. GetTemplate()
'                                        to populate location, code cleanup
'               BLC - 1/11/2017 - 1.04 - added location button/combobox
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
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
'   BLC - 10/20/2016 - revised to use callingform property, GetContext()
'   BLC - 1/10/2017 - revised to use GetRecords() vs. GetTemplate() to populate location,
'                     code cleanup
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
                 
    Title = "Feature"
    Directions = "Enter the feature details."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    lblDirections.Caption = "Enter feature details."
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    
    'set hover
    btnComment.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    tbxFeature.BackColor = lngYellow
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
    
    'get locations for park or river segment depending upon if river segment is set
    With cbxLocation
        '.RowSource = IIf(Len(Nz(TempVars("River"), "")) > 0, GetTemplate("s_location_by_park_river"), GetTemplate("s_location_by_park"))
        Set .Recordset = IIf(Len(Nz(TempVars("River"), "")) > 0, _
                        GetRecords("s_location_by_park_river"), _
                        GetRecords("s_location_by_park"))
                        
        .ColumnCount = 2
        .BoundColumn = 1
        .ColumnWidths = "0;1in"
    End With
    
    'cbxLocation.RowSource = GetTemplate("s_location_by_park", "parkcode" & PARAM_SEPARATOR & TempVars.item("ParkCode"))
  
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
            "Error encountered (#" & Err.Number & " - Form_Open[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Current[Feature form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxFeature_AfterUpdate
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
Private Sub tbxFeature_AfterUpdate()
On Error GoTo Err_Handler

    If Len(tbxFeature.Text) > 0 Then _
        ReadyForSave

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxFeature_Change[Feature form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxLocation_AfterUpdate
' Description:  AfterUpdate actions
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
Private Sub cbxLocation_AfterUpdate()
On Error GoTo Err_Handler

    If Len(cbxLocation.Text) > 0 Then _
        ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxLocation_AfterUpdate[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
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
            "Error encountered (#" & Err.Number & " - btnUndo_Click[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    UpsertRecord Me
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[Feature form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddLocation_Click
' Description:  Add location button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 11, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2017 - initial version
' ---------------------------------
Private Sub btnAddLocation_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Location", acNormal, , , , , "Transect|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddLocation_Click[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "event|" & tbxID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
'   BLC - 10/20/2016 - revised to use callingform property
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
            "Error encountered (#" & Err.Number & " - Form_Close[Feature form])"
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
' Source/date:  Bonnie Campbell, June 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/27/2016 - initial version
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: feature & location ID
    If Len(Nz(tbxFeature.Value, "")) > 0 Then
        isOK = True
    End If
    
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
            "Error encountered (#" & Err.Number & " - ReadyForSave[Feature form])"
    End Select
    Resume Exit_Handler
End Sub
