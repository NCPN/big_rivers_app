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
    ItemSuffix =37
    Left =7440
    Top =1515
    Right =15300
    Bottom =10740
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x7c316d7d90cae440
    End
    Caption ="Location (Sampling Location)"
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
            Height =1380
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
                    Caption ="Location (Sampling Location)"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =180
                    Top =420
                    Width =6840
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Enter the sampling location information and click save."
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
                    Name ="lblBearing"
                    Caption ="Bearing"
                    GridlineColor =10921638
                    LayoutCachedLeft =4260
                    LayoutCachedTop =1065
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =1380
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
                    Caption ="Distance"
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
                    Name ="lblName"
                    Caption ="Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1065
                    LayoutCachedWidth =2445
                    LayoutCachedHeight =1380
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
                    ForeColor =16777215
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
                Begin Label
                    OverlapFlags =215
                    Left =180
                    Top =720
                    Width =855
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTransect"
                    Caption ="Transect"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =720
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =1035
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =1260
                    Top =720
                    Width =2514
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
                    Name ="cbxTransect"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;0;0;1"
                    ControlTipText ="Veg transect"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1260
                    LayoutCachedTop =720
                    LayoutCachedWidth =3774
                    LayoutCachedHeight =1035
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
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7860
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6660
                    Top =60
                    Width =720
                    TabIndex =4
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
                    Name ="tbxBearing"
                    AfterUpdate ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    ControlTipText ="Bearing 0 to 360 degrees"
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001600000001000000 ,
                        0x00000000fff200000000000003000000180000001b0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800420065006100720069006e0067005d002e00560061006c00 ,
                        0x750065003d002200220000000000000022002200000000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =4140
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
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
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =75
                    Width =720
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =255
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
                    Left =5760
                    Top =60
                    Width =720
                    TabIndex =6
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
                Begin Subform
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =105
                    Top =3360
                    Width =7650
                    Height =4380
                    TabIndex =7
                    BorderColor =10921638
                    Name ="list"
                    SourceObject ="Form.LocationList"
                    GridlineColor =10921638

                    LayoutCachedLeft =105
                    LayoutCachedTop =3360
                    LayoutCachedWidth =7755
                    LayoutCachedHeight =7740
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Top =3240
                    Width =7860
                    Height =4620
                    BackColor =4144959
                    BorderColor =10921638
                    Name ="rctList"
                    GridlineColor =10921638
                    LayoutCachedTop =3240
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =7860
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
                    TabIndex =8
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
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000098000000020000000100000000000000000000001700000001000000 ,
                        0x00000000fff200000000000003000000180000001b0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800440069007300740061006e00630065005d002e0056006100 ,
                        0x6c00750065003d0022002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =60
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000160000005b00 ,
                        0x740062007800440069007300740061006e00630065005d002e00560061006c00 ,
                        0x750065003d002200220000000000000000000000000000000000000000000000 ,
                        0x000000030000000100000000000000ffffff0002000000220022000000000000 ,
                        0x00000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =360
                    Top =900
                    Width =7140
                    Height =1860
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxNotes"
                    ControlTipText ="Enter establishment comments and other notes."
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =900
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =2760
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =480
                            Width =690
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblNotes"
                            Caption ="Notes:"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =480
                            LayoutCachedWidth =870
                            LayoutCachedHeight =795
                        End
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
                    Name ="tbxName"
                    AfterUpdate ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000100000000000000000000001300000001000000 ,
                        0x00000000fff20000000000000300000014000000170000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004e0061006d0065005d002e00560061006c00750065003d00 ,
                        0x22002200000000002200220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =375
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000120000005b00 ,
                        0x7400620078004e0061006d0065005d002e00560061006c00750065003d002200 ,
                        0x2200000000000000000000000000000000000000000000000000000300000001 ,
                        0x00000000000000ffffff00020000002200220000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    TextAlign =3
                    Top =2940
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
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =3255
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =4320
                    Top =2760
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblMsgIcon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =2760
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =3360
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
' Form:         Location
' Level:        Application form
' Version:      1.05
' Basis:        Dropdown form
'
' Description:  Location form object related properties, Location, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, May 31, 2016
' References:   -
' Revisions:    BLC - 5/31/2016 - 1.00 - initial version
'               BLC - 7/29/2016 - 1.01 - revised to use UpsertRecord()
'               BLC - 8/23/2016 - 1.02 - changed ReadyForSave() to public for
'                                        mod_App_Data Upsert/SetRecord()
'               BLC - 9/1/2016  - 1.03 - btnSave_Click, btnUndo_Click code cleanup
'               BLC - 10/20/2016 - 1.04 - added CallingForm property, removed buttoncaption,
'                                         selectedID, selectedValue properties
'               BLC - 1/11/2017 - 1.05 - hide title label, CallingRecordID
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
Private m_CallingRecordID As Integer
Private m_LocationType As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidCallingRecordID(Value As Integer)
Public Event InvalidLocationType(Value As String)

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

Public Property Let CallingRecordID(Value As Integer)
        m_CallingRecordID = Value
End Property

Public Property Get CallingRecordID() As Integer
    CallingRecordID = m_CallingRecordID
End Property

Public Property Let LocationType(Value As String)
        m_LocationType = Value
End Property

Public Property Get LocationType() As String
    LocationType = m_LocationType
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
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
'   BLC - 10/20/2016 - revised to use CallingForm property, GetContext()
'   BLC - 1/11/2017 - hide second title
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

'    Dim aryOArgs() As Variant
    Dim ltype As String
    
    'default
    Me.CallingForm = "Main"
    Me.CallingRecordID = -1
    
'    If Len(Me.OpenArgs) > 0 Then
'        aryOArgs = Split(Me.OpenArgs, "|")
'
'        If IsArray(aryOArgs) Then
'
'            'set calling for & record
'            Me.CallingForm = aryOArgs(0)
'            Me.CallingRecordID = Nz(aryOArgs(1), -1)
'
'        End If
'    End If
    
    'set location type
    ltype = "" 'default
    
    Select Case Me.CallingForm
        Case "Feature"
            ltype = "F"
        Case "Transect"
            ltype = "T"
        Case "VegPlot"
            ltype = "P"
    End Select
    
    Me.LocationType = ltype
    
    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()
    
    Title = "Location (Sampling Location)"
    lblTitle.Caption = "" 'hide second title
    Directions = "Enter the sampling location information and click save."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
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
    tbxDistance.BackColor = lngYellow
    tbxBearing.BackColor = lngYellow
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
    
    'ID default -> value used only for edits of existing table values
    tbxID.Value = 0
    
    'initialize values
    ClearForm Me

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
' Sub:          tbxName_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/29/2016 - initial version
' ---------------------------------
Private Sub tbxName_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxName_AfterUpdate[Location form])"
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
' Source/date:  Bonnie Campbell, July 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/29/2016 - initial version
' ---------------------------------
Private Sub tbxDistance_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxDistance_AfterUpdate[Location form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxBearing_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 29, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/29/2016 - initial version
' ---------------------------------
Private Sub tbxBearing_AfterUpdate()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxBearing_AfterUpdate[Location form])"
    End Select
    Resume Exit_Handler
End Sub

'' ---------------------------------
'' Sub:          tbxStartDate_Change
'' Description:  Dropdown change actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 6/1/2016 - initial version
'' ---------------------------------
'Private Sub tbxStartDate_Change()
'On Error GoTo Err_Handler
'
'    ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tbxStartDate_Change[Location form])"
'    End Select
'    Resume Exit_Handler
'End Sub
'
'' ---------------------------------
'' Sub:          tbxStartDate_LostFocus
'' Description:  Dropdown change actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 6/1/2016 - initial version
'' ---------------------------------
'Private Sub tbxStartDate_LostFocus()
'On Error GoTo Err_Handler
'
'    'ReadyForSave
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - tbxStartDate_LostFocus[Location form])"
'    End Select
'    Resume Exit_Handler
'End Sub

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
'   BLC - 9/1/2016 - commented code cleanup
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
'   BLC - 7/29/2016 - revised to use UpsertRecord()
'   BLC - 9/1/2016  - commented code cleanup
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
            "Error encountered (#" & Err.Number & " - btnSave_Click[Location form])"
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
    DoCmd.OpenForm "Comment", acNormal, , , , , "location|" & tbxID.text
    
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
'   BLC - 10/20/2016 - revised to use CallingForm property
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
'   BLC - 8/23/2016 - changed ReadyForSave() to public for mod_App_Data Upsert/SetRecord()
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    If Len(tbxName.Value) > 0 And tbxDistance.Value > 0 And tbxBearing.Value <> "" Then
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
            "Error encountered (#" & Err.Number & " - ReadyForSave[Location form])"
    End Select
    Resume Exit_Handler
End Sub
