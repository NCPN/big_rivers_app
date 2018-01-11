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
    ItemSuffix =45
    Left =5850
    Top =2805
    Right =13710
    Bottom =10650
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x26c502515408e540
    End
    Caption ="Set batch photo data"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
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
            Height =1320
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
                    Caption ="Set Photo Info"
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
                    Width =7500
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Set the sampling event, photographer and type for the selected photos && update"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =420
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =735
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
                    OverlapFlags =215
                    TextAlign =3
                    Left =3360
                    Top =60
                    Width =4380
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblContext"
                    Caption ="context"
                    GridlineColor =10921638
                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =108
                    Top =960
                    Width =1755
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =6750156
                    Name ="lblRecordRefID"
                    GridlineColor =10921638
                    LayoutCachedLeft =108
                    LayoutCachedTop =960
                    LayoutCachedWidth =1863
                    LayoutCachedHeight =1275
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6540
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =1080
                    Top =2460
                    Width =4860
                    Height =3420
                    BackColor =12566463
                    BorderColor =10921638
                    Name ="rctPhoto"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =2460
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =5880
                    BackShade =75.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6840
                    Top =4320
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

                    LayoutCachedLeft =6840
                    LayoutCachedTop =4320
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =4680
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
                    Visible = NotDefault
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
                    TabIndex =4
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
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6060
                    Top =4320
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

                    LayoutCachedLeft =6060
                    LayoutCachedTop =4320
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =4680
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
                    Visible = NotDefault
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
                    ControlSource ="ID"
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
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2580
                    Top =480
                    Width =2820
                    Height =315
                    FontSize =8
                    TabIndex =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001800000001000000 ,
                        0x00000000fff200000000000003000000190000001c0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005300740061007200740044006100740065005d002e005600 ,
                        0x61006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxPhotographer"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2160"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Select photographer"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2580
                    LayoutCachedTop =480
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =795
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000170000005b00 ,
                        0x7400620078005300740061007200740044006100740065005d002e0056006100 ,
                        0x6c00750065003d00220022000000000000000000000000000000000000000000 ,
                        0x0000000000030000000100000000000000ffffff000200000022002200000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Top =6120
                    Width =7860
                    Height =315
                    FontSize =9
                    LeftMargin =360
                    TopMargin =36
                    RightMargin =360
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblMsg"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedTop =6120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =6435
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =4320
                    Top =5940
                    Width =825
                    Height =600
                    FontSize =20
                    BackColor =4144959
                    BorderColor =8355711
                    ForeColor =6750105
                    Name ="lblMsgIcon"
                    FontName ="Segoe UI"
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =5940
                    LayoutCachedWidth =5145
                    LayoutCachedHeight =6540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =215
                    MultiSelect =2
                    IMESentenceMode =3
                    Left =1320
                    Top =2940
                    Width =4380
                    Height =2820
                    FontSize =8
                    BackColor =65535
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxPhotos"
                    RowSourceType ="Table/Query"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Select photos to update"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1320
                    LayoutCachedTop =2940
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =5760
                    BackThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    Left =1080
                    Top =60
                    Width =600
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblEvent"
                    Caption ="Event"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =375
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    Left =1980
                    Top =60
                    Width =3414
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
                    Name ="cbxEvent"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;0;0;2880"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Event (sample visit)"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1980
                    LayoutCachedTop =60
                    LayoutCachedWidth =5394
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =5520
                    Top =60
                    TabIndex =7
                    ForeColor =16711680
                    Name ="btnAddEvent"
                    Caption ="í ½í·“  Add Event"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Add a new event/sampling visit"
                    GridlineColor =10921638

                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =420
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
                    OverlapFlags =215
                    TextAlign =2
                    Left =1080
                    Top =2160
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUnclassifiedPhotos"
                    Caption ="Photos"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =2160
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =2475
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1080
                    Top =480
                    Width =1350
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotographer"
                    Caption ="Photographer"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =480
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =795
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5520
                    Top =480
                    Width =720
                    TabIndex =8
                    ForeColor =4210752
                    Name ="btnAddContact"
                    Caption ="í ½í±¥"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5520
                    LayoutCachedTop =480
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =840
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
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2040
                    Top =2580
                    Width =3720
                    Height =315
                    FontSize =8
                    TabIndex =9
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001800000001000000 ,
                        0x00000000fff200000000000003000000190000001c0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005300740061007200740044006100740065005d002e005600 ,
                        0x61006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxPhotoFilter"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2160"
                    ControlTipText ="Select photo type(s) to display"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2040
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =2895
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000170000005b00 ,
                        0x7400620078005300740061007200740044006100740065005d002e0056006100 ,
                        0x6c00750065003d00220022000000000000000000000000000000000000000000 ,
                        0x0000000000030000000100000000000000ffffff000200000022002200000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1320
                    Top =2580
                    Width =570
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblPhotoFilter"
                    Caption ="Filter"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =2580
                    LayoutCachedWidth =1890
                    LayoutCachedHeight =2895
                    ForeTint =75.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =6240
                    Top =5640
                    TabIndex =10
                    ForeColor =16711680
                    Name ="btnUpdatePhotos"
                    Caption ="í ½í·“  Update Photos"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Update photos"
                    GridlineColor =10921638

                    LayoutCachedLeft =6240
                    LayoutCachedTop =5640
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =6000
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
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =1080
                    Top =1320
                    Width =6120
                    Height =720
                    FontSize =8
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16737792
                    Name ="lblPhotoTypesHint"
                    Caption ="Photo types hint"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =1320
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =2040
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2580
                    Top =900
                    Width =2820
                    Height =315
                    FontSize =8
                    TabIndex =11
                    BoundColumn =1
                    BackColor =65535
                    BorderColor =10921638
                    ForeColor =4210752
                    ConditionalFormat = Begin
                        0x010000009a000000020000000100000000000000000000001800000001000000 ,
                        0x00000000fff200000000000003000000190000001c0000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078005300740061007200740044006100740065005d002e005600 ,
                        0x61006c00750065003d0022002200000000002200220000000000
                    End
                    Name ="cbxPhotoType"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="1440;1440;1440"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Select desired photo type"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2580
                    LayoutCachedTop =900
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1215
                    BackThemeColorIndex =-1
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000100000000000000fff20000170000005b00 ,
                        0x7400620078005300740061007200740044006100740065005d002e0056006100 ,
                        0x6c00750065003d00220022000000000000000000000000000000000000000000 ,
                        0x0000000000030000000100000000000000ffffff000200000022002200000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1080
                    Top =900
                    Width =1350
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPhotoType"
                    Caption ="Photo Type"
                    ControlTipText ="Select desired photo type"
                    GridlineColor =10921638
                    LayoutCachedLeft =1080
                    LayoutCachedTop =900
                    LayoutCachedWidth =2430
                    LayoutCachedHeight =1215
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
' Form:         PhotoBatchUpdate
' Level:        Application form
' Version:      1.02
' Basis:        Dropdown form
'
' Description:  Batch photo form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, December 8, 2017
' References:   -
' Revisions:    BLC - 12/8/2017  - 1.00 - initial version
'               BLC - 12/11/2017 - 1.01 - renamed PhotoBatchUpdate vs PhotoBulkUpdate
'               BLC - 1/3/2018   - 1.02 - add PhotoType recordset, SelPhoto & SelPhotos properties
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

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

Private m_SelPhotos As Collection
Private m_SelPhoto As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

Public Event InvalidSelPhoto(Value As Long)

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

Public Property Let SelPhotos(Value As Collection)
'    If  Then
        Set m_SelPhotos = Value
'    Else
'        RaiseEvent InvalidSelPhotos(Value)
'    End If
End Property

Public Property Get SelPhotos() As Collection
    Set SelPhotos = m_SelPhotos
End Property

Public Property Let SelPhoto(Value As Long)
    If IsNumeric(Value) Then
        m_SelPhoto = Value
    Else
        RaiseEvent InvalidSelPhoto(Value)
    End If

    'check if value is already present
    Dim InCollection As Boolean
    InCollection = False
    Dim i As Long
    
    For i = 1 To Me.SelPhotos.Count
        If SelPhotos.item(i) = Value Then
            InCollection = True
            Exit For
        End If
    Next
    
    If InCollection = False Then
        'add to the collection
        Me.SelPhotos.Add Value
    End If
    
End Property

Public Property Get SelPhoto() As Long
    SelPhoto = m_SelPhoto
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  OpenArgs passes only the calling form name
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
'   BLC - 1/3/2018  - add PhotoType recordset, SelPhotos collection property
'                     revised default calling form (Photos vs Main)
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
    Me.CallingForm = "Photos"
    
    If Len(Nz(Me.OpenArgs, "")) > 0 And _
        Len(Nz(Me.OpenArgs, "")) <> Replace(Nz(Me.OpenArgs, ""), "|", "") Then
        
        'set the referencing table & record
        Me.CallingForm = Split(Me.OpenArgs, "|")(0)
        lblRecordRefID.Caption = Me.CallingForm _
                                & " ID # " & Split(Me.OpenArgs, "|")(1)
    End If

    'minimize Main
    ToggleForm Me.CallingForm, -1
    
    'set context - based on TempVars
    lblContext.ForeColor = lngLime
    lblContext.Caption = GetContext()

    Title = "Set Photo Info"
    Directions = "Set the sampling event, photographer and type for the selected photos."
    tbxIcon.Value = StringFromCodepoint(uBullet)
    lblDirections.ForeColor = lngLtBlue
    lblPhotoTypesHint.Caption = "Photo Types: " & vbCrLf & _
                                "F-feature" & Space(2) & "T-transect" & Space(2) & "O-overview" & Space(2) & "R-reference" & Space(2) & "U-unclassified " & vbCrLf & _
                                "OA-OtherAnimal" & Space(2) & "OP-Plant" & Space(2) & "OC-Cultural" & Space(2) & "OD-Disturbance" & Space(2) & "OF-Field Work" & Space(2) & "OS-Scenic" & Space(2) & "OW-Weather" & Space(2) & "OO-Other"
    lblPhotoTypesHint.ForeColor = lngBlue
    btnComment.Caption = StringFromCodepoint(uComment)
    btnComment.ForeColor = lngBlue
    btnAddEvent.Caption = StringFromCodepoint(uCalendarSpiral) & Space(2) & "Add Event"
    btnAddEvent.ForeColor = lngBlue
    btnAddContact.Caption = StringFromCodepoint(uUsers) & Space(2) & "Add Photographer"
    btnAddContact.ForeColor = lngBlue
    btnUpdatePhotos.Caption = StringFromCodepoint(uPicFramed) & Space(2) & "Update Photos"
    btnUpdatePhotos.ForeColor = lngBlue
    
    lblRecordRefID.ForeColor = lngLtLime
        
    'set hover
    btnAddEvent.HoverColor = lngGreen
    btnAddContact.HoverColor = lngGreen
    btnUpdatePhotos.HoverColor = lngGreen
    btnComment.HoverColor = lngGreen
    btnSave.HoverColor = lngGreen
    btnUndo.HoverColor = lngGreen
      
    'hidden (unused) controls
    btnSave.Visible = False
    btnUndo.Visible = False
    lblPhotoFilter.Visible = False
    cbxPhotoFilter.Visible = False
      
    'defaults
    tbxIcon.ForeColor = lngRed
    btnComment.Enabled = False
    btnSave.Enabled = False
    lblMsgIcon.Caption = ""
    lblMsg.Caption = ""
  
    'ID default -> value used only for edits of existing table values
    tbxID.DefaultValue = 0
    
    'initialize values << place here before initial call to Form_Current()
    '                     driven by setting record sources
    Dim col As New Collection
    Me.SelPhotos = New Collection 'col
    
    'clear form datasource in case it was saved (to keep unbound)
    Me.RecordSource = ""
    
    'set data sources
    Set cbxPhotographer.Recordset = GetRecords("s_contact_list")
    'lbxPhotos:  ID, PhotoType, PhotoPath, PhotoFilename, PhotoDate,  Event_ID, Photographer_ID
    Set lbxPhotos.Recordset = GetRecords("s_usys_temp_photo_list")
    
    SetTempVar "EnumType", "PhotoType"
    Set cbxPhotoType.Recordset = GetRecords("s_app_enum_list")
    cbxPhotoType.ColumnHeads = True
    cbxPhotoType.ColumnCount = 3            'ID, type abbrev, type name
    cbxPhotoType.BoundColumn = 2            'type abbrev
    cbxPhotoType.ColumnWidths = "0;0;1;"    'display only type name
    
    'determine what level to populate
    Dim efilter As String
    
    'site is default
    efilter = "s_events_by_site"
    cbxEvent.ColumnCount = 6
    cbxEvent.ColumnWidths = "0;0;0;2in;0;0"

    Select Case TempVars("ParkCode")
        Case "BLCA" 'feature level if set
            If Not TempVars("Feature") Is Nothing Then _
                efilter = "s_events_by_feature"
        Case "CANY" 'site level
        Case "DINO" 'no transects/plots
    End Select
        
    'populate events
    Set cbxEvent.Recordset = GetRecords(efilter)
    cbxEvent.BoundColumn = 1
    'cbxEvent.ColumnCount = 5
    'cbxEvent.ColumnWidths = "0in;0in;0in;0in;2in"
    
    'populate photo types
'    cbxPhotoType.BoundColumn = 1
'    cbxPhotoType.ColumnCount = 2
'    cbxPhotoType.ColumnWidths = "1in;3in"
    
    'populate photo list
    ' ID, PhotoType, PhotoPath, PhotoFilename, PhotoDate,  Event_ID, Photographer_ID
    'Event_ID , Photographer_ID, PhotoPath, PhotoFilename, PhotoDate, PhotoType
    lbxPhotos.ColumnHeads = True
    lbxPhotos.BoundColumn = 1
    lbxPhotos.ColumnCount = 7
    lbxPhotos.ColumnWidths = "0;0.3in;0;2.4in;1in;0;0;"
    'lbxPhotos.MultiSelect = True
    'lbxPhotos.MultiSelect = '"Extended"
    
    'set list data source
    
'    Set Me.list.Form.Recordset = GetRecords("s_record_action_by_refID", Params)
'    Me.list.Form.Requery
    
    'initialize values
    ClearForm Me
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Load[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
              
'      If tbxID > 0 Then btnComment.Enabled = True
    'MsgBox tbxID, vbCritical, "Current"
    'If tbxID > 0 Then ReadyForSave

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeUpdate
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
              
    If Not m_SaveOK Then
        Cancel = True
    End If
    'Cancel = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxPhotographer_AfterUpdate
' Description:  Dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub cbxPhotographer_AfterUpdate()
On Error GoTo Err_Handler

'    Me.PhotographerID = cbxPhotographer.Value
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPhotographer_AfterUpdate[SetPhotographerPhotographer form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxEvent_AfterUpdate
' Description:  Combobox after update actions
' Assumptions:  Event combobox contains the following columns:
'                   column(0)= event ID                 column(3)= event date - site name (sitecode)
'                   column(1)= event date               column(4)= site code
'                   column(2)= event date - site code   column(5)= park code
'               Column 1 (event date) will be used to determine the proper MSS year
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub cbxEvent_AfterUpdate()
On Error GoTo Err_Handler
    
    'check if ready
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxEvent_AfterUpdate[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxPhotoType_AfterUpdate
' Description:  Dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 11, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/11/2017 - initial version
' ---------------------------------
Private Sub cbxPhotoType_AfterUpdate()
On Error GoTo Err_Handler
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPhotoType_AfterUpdate[SetPhotoTypePhotoType form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lbxPhotos_AfterUpdate
' Description:  Dropdown after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Microsoft (office 365 dev), June 12, 2017
'   https://msdn.microsoft.com/en-us/vba/access-vba/articles/listbox-itemsselected-property-access
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
'   BLC - 1/3/2018  - set SelPhotos from selected images
' ---------------------------------
Private Sub lbxPhotos_AfterUpdate()
On Error GoTo Err_Handler
    
    'clear the overall collection
    Me.SelPhotos = New Collection
    
    Dim item As Variant 'items selected are variants
    
    For Each item In lbxPhotos.ItemsSelected
        'add photo to selected photos collection
        Me.SelPhoto = lbxPhotos.ItemData(item)
Debug.Print lbxPhotos.ItemData(item)
    Next
    
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxPhotos_AfterUpdate[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
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
            "Error encountered (#" & Err.Number & " - btnUndo_Click[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddEvent_Click
' Description:  Add event button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub btnAddEvent_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "Events", acNormal, , , , , Me.Name
    
    'refresh cbx (done from event form close)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEvent_Click[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddContact_Click
' Description:  Add Contact button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub btnAddContact_Click()
On Error GoTo Err_Handler
    
    'open form
    DoCmd.OpenForm "Contact", acNormal, , , , , Me.Name
    
    'refresh cbx (done from Contact form close)
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddContact_Click[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub btnSave_Click()
On Error GoTo Err_Handler
    
    'set enable btnSave_Click save
    m_SaveOK = True
        
    UpsertRecord Me
    
    Me![list].Form.Requery
    
    'revert to disable non-btnSave_Click save
    m_SaveOK = False
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSave_Click[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnUpdatePhotos_Click
' Description:  Update photos button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 11, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/11/2017 - initial version
' ---------------------------------
Private Sub btnUpdatePhotos_Click()
On Error GoTo Err_Handler
    
'    'set enable btnSave_Click save
'    m_SaveOK = True
       
    Dim i As Integer
    
    With lbxPhotos
        If .ItemsSelected.Count > 0 Then
        Debug.Print "#selected " & .ItemsSelected.Count
            For i = 1 To .ItemsSelected.Count
                'UpsertRecord Me
                Debug.Print .ItemData(i)
                
                UpsertRecord Me
            Next
        End If
    End With
    
'    UpsertRecord Me
'
'    Me![list].Form.Requery
'
'    'revert to disable non-btnSave_Click save
'    m_SaveOK = False
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUpdatePhotos_Click[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Private Sub btnComment_Click()
On Error GoTo Err_Handler
    
    'open comment form
    DoCmd.OpenForm "Comment", acNormal, , , , , "photoBatchupdate|" & tbxID & "|255"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnComment_Click[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
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
            "Error encountered (#" & Err.Number & " - Form_Close[PhotoBatchUpdate form])"
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
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
'   BLC - 1/3/2018  - revised to pass Int(x) values to IsInt to avoid
'                     false results due to strings (vs numerics)
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    'set color of icon depending on if values are set
    'requires: Event_ID, Photographer_ID, PhotoType, Photos selected
    'cbxPhotoType Columns: 0 - ID, 1 - type abbrev, 3 - type name
    If cbxEvent > 0 _
        And IsInt(Int(cbxPhotographer)) _
        And IsInt(Int(cbxPhotoType.Column(0))) _
        And lbxPhotos.ItemsSelected.Count > 0 _
        Then
            isOK = True
    End If
    
    tbxIcon.ForeColor = IIf(isOK = True, lngDkGreen, lngRed)
    'btnSave.Enabled = isOK
    btnUpdatePhotos.Enabled = isOK
    
    'refresh form
'    Me.Requery
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          RunReadyForSave
' Description:  Run ready for save check from another form (public method)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017 - initial version
' ---------------------------------
Public Sub RunReadyForSave()
On Error GoTo Err_Handler

    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunReadyForSave[PhotoBatchUpdate form])"
    End Select
    Resume Exit_Handler
End Sub
